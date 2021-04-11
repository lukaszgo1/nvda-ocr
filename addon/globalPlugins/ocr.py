# -*- coding: UTF-8 -*-
"""NVDA OCR plugin
This plugin uses Tesseract for OCR: https://github.com/tesseract-ocr
@author: James Teh <jamie@nvaccess.org>
@author: Rui Batista <ruiandrebatista@gmail.com>
@copyright: 2011-2020 NV Access Limited, Rui Batista, Łukasz Golonka
@license: GNU General Public License version 2.0
"""

import os
import tempfile
import subprocess
from xml.parsers import expat
from collections import namedtuple
from copy import copy
import wx
import globalPluginHandler
import gui
import api
import languageHandler
import addonHandler
import textInfos.offsets
import ui
import locationHelper
import scriptHandler
import config
from io import StringIO
from configobj import ConfigObj
from logHandler import log

addonHandler.initTranslation()


qualityInfo = namedtuple("qualityInfo", ("translatedName", "dirName"))

# qualities of tessdata and OCR result
OCR_QUALITIES = {
	# Translators: OCR optimization label in settings
	0: qualityInfo(_("speed"), "fast"),
	# Translators: OCR optimization label in settings
	1: qualityInfo(_("quality"), "best")
}

POSSIBLE_QUALITIES = ''.join('"{}", '.format(q.dirName) for q in OCR_QUALITIES.values())

priorityInfo = namedtuple("priorityInfo", ("translatedName", "priorityConstant"))

# priorities with which Tesseract.exe could be started to
OCR_PRIORITIES = {
	# Translators: OCR priority in settings
	"below_normal": priorityInfo(_("below normal"), 0x00004000),
	# Translators: OCR priority in settings
	"normal": priorityInfo(_("normal"), 0x00000020),
	# Translators: OCR priority in settings
	"above_normal": priorityInfo(_("above normal"), 0x00008000),
	# Translators: OCR priority in settings
	"high": priorityInfo(_("high"), 0x00000080),
	# Translators: OCR priority in settings
	"real_time": priorityInfo(_("real-time"), 0x00000100)
}

POSSIBLE_PRIORITIES = ''.join('"{}", '.format(prioName) for prioName in OCR_PRIORITIES.keys())

PLUGIN_DIR = os.path.dirname(__file__)
TESSERACT_EXE = os.path.join(PLUGIN_DIR, "tesseract", "bin", "tesseract.exe")
TESSDATA_BASEDIR = os.path.join(PLUGIN_DIR, "tesseract", "tessdata")

IMAGE_RESIZE_FACTOR = 2

OcrWord = namedtuple("OcrWord", ("offset", "left", "top"))


class LanguageInfo:

	"""Provides information about a single language supported by Tesseract."""

	# Tesseract identifies languages using  their ISO 639-2 language codes
	# whereas NVDA locales are identifed by ISO 639-1 codes.
	# Below dictionaries provide mapping from one representation to the other.
	NVDALocalesToTesseractLangs = {
		"bg": "bul",
		"ca": "cat",
		"cs": "ces",
		"zh_CN": "chi_tra",
		"da": "dan",
		"de": "deu",
		"el": "ell",
		"en": "eng",
		"fi": "fin",
		"fr": "fra",
		"hu": "hun",
		"id": "ind",
		"it": "ita",
		"ja": "jpn",
		"ko": "kor",
		"lv": "lav",
		"lt": "lit",
		"nl": "nld",
		"nb_NO": "nor",
		"pl": "pol",
		"pt": "por",
		"ro": "ron",
		"ru": "rus",
		"sk": "slk",
		"sl": "slv",
		"es": "spa",
		"sr": "srp",
		"sv": "swe",
		"tg": "tgl",
		"tr": "tur",
		"uk": "ukr",
		"vi": "vie"
	}

	tesseractLangsToNVDALocales = {v: k for k, v in NVDALocalesToTesseractLangs.items()}

	TesseractLocalesToWindowsLocalizedLangNames = dict()
	WindowsLocalizedLangNamesToTesseractLocales = dict()

	FALLBACK_LANGUAGE = "eng"

	def __init__(self, NVDALocaleName=None, TesseractLocaleName=None, localizedName=None):
		self._NVDALocaleName = NVDALocaleName
		self._TesseractLocaleName = TesseractLocaleName
		self._localizedName = localizedName
		if self._NVDALocaleName and self._TesseractLocaleName is None:
			self._TesseractLocaleName = self.NVDALocalesToTesseractLangs[self._NVDALocaleName]
		elif self._TesseractLocaleName and self._NVDALocaleName is None:
			self._NVDALocaleName = self.tesseractLangsToNVDALocales[self._TesseractLocaleName]
		if(
			self._TesseractLocaleName
			and self._TesseractLocaleName in self.TesseractLocalesToWindowsLocalizedLangNames
		):
			self._localizedName = self.TesseractLocalesToWindowsLocalizedLangNames[self._TesseractLocaleName]
		if(
			self._localizedName
			and self._TesseractLocaleName is None
			and self._localizedName in self.WindowsLocalizedLangNamesToTesseractLocales
		):
			self._TesseractLocaleName = self.WindowsLocalizedLangNamesToTesseractLocales[self._localizedName]

	@staticmethod
	def availableTesseractLanguageFiles(quality):
		for file in os.listdir(os.path.join(PLUGIN_DIR, "tesseract", "tessdata", quality)):
			if file.endswith(".traineddata"):
				yield os.path.splitext(file)[0]

	@classmethod
	def fromAvailableLanguages(cls, quality):
		for langFN in cls.availableTesseractLanguageFiles(quality):
			yield cls(TesseractLocaleName=langFN)

	@classmethod
	def fromConfiguredLanguage(cls):
		return cls(TesseractLocaleName=config.conf["ocr"]["language"])

	@classmethod
	def fromFallbackLanguage(cls):
		return cls(TesseractLocaleName=LanguageInfo.FALLBACK_LANGUAGE)

	@classmethod
	def fromCurrentNVDALanguage(cls):
		currentNVDALang = languageHandler.getLanguage()
		for possibleLocaleName in (currentNVDALang, currentNVDALang.split("_")[0], cls.FALLBACK_LANGUAGE):
			try:
				return cls(NVDALocaleName=possibleLocaleName)
			except KeyError:
				continue

	@property
	def localizedName(self):
		"""Returns localized name of the language with which this object was initialized."""
		res = self._localizedName
		if res is None:
			res = languageHandler.getLanguageDescription(self._NVDALocaleName)
			if res:
				self.__class__.WindowsLocalizedLangNamesToTesseractLocales[res] = self._TesseractLocaleName
				self.__class__.TesseractLocalesToWindowsLocalizedLangNames[self._TesseractLocaleName] = res
			else:
				# If there is no localized name for the given locale just return a language code.
				# This is better than no name at all.
				res = self._NVDALocaleName
		return res

	@property
	def TesseractLocaleName(self):
		return self._TesseractLocaleName


class HocrParser(object):

	def __init__(self, xml, leftCoordOffset, topCoordOffset):
		self.leftCoordOffset = leftCoordOffset
		self.topCoordOffset = topCoordOffset
		parser = expat.ParserCreate("utf-8")
		parser.StartElementHandler = self._startElement
		parser.EndElementHandler = self._endElement
		parser.CharacterDataHandler = self._charData
		self._textList = []
		self.textLen = 0
		self.lines = []
		self.words = []
		self._hasBlockHadContent = False
		parser.Parse(xml)
		self.text = "".join(self._textList)
		del self._textList

	def _startElement(self, tag, attrs):
		if tag in ("p", "div"):
			self._hasBlockHadContent = False
		elif tag == "span":
			cls = attrs["class"]
			if cls == "ocr_line":
				self.lines.append(self.textLen)
			elif cls == "ocrx_word":
				# Get the coordinates from the bbox info specified in the title attribute.
				title = attrs.get("title")
				# cut non-bbox info if present
				titleBbox = title.split(";")[0]
				prefix, l, t, r, b = titleBbox.split(" ")
				self.words.append(OcrWord(self.textLen,
					self.leftCoordOffset + int(l) / IMAGE_RESIZE_FACTOR,
					self.topCoordOffset + int(t) / IMAGE_RESIZE_FACTOR))

	def _endElement(self, tag):
		pass

	def _charData(self, data):
		if data.isspace():
			if not self._hasBlockHadContent:
				# Strip whitespace at the start of a block.
				return
			# All other whitespace should be collapsed to a single space.
			data = " "
			if self._textList and self._textList[-1] == data:
				return
		self._hasBlockHadContent = True
		self._textList.append(data)
		self.textLen += len(data)


class OcrTextInfo(textInfos.offsets.OffsetsTextInfo):

	def __init__(self, obj, position, parser):
		self._parser = parser
		super(OcrTextInfo, self).__init__(obj, position)

	def copy(self):
		return self.__class__(self.obj, self.bookmark, self._parser)

	def _getTextRange(self, start, end):
		return self._parser.text[start:end]

	def _getStoryLength(self):
		return self._parser.textLen

	def _getLineOffsets(self, offset):
		start = 0
		for end in self._parser.lines:
			if end > offset:
				return (start, end)
			start = end
		return (start, self._parser.textLen)

	def _getWordOffsets(self, offset):
		start = 0
		for word in self._parser.words:
			if word.offset > offset:
				return (start, word.offset)
			start = word.offset
		return (start, self._parser.textLen)

	def _getPointFromOffset(self, offset):
		for nextWord in self._parser.words:
			if nextWord.offset > offset:
				break
			word = nextWord
		else:
			# No matching word, so use the top left of the object.
			return locationHelper.Point(int(self._parser.leftCoordOffset), int(self._parser.topCoordOffset))
		return locationHelper.Point(int(word.left), int(word.top))


configSpecString = f"""
	language = string(default={LanguageInfo.fromCurrentNVDALanguage().TesseractLocaleName})
	quality = option({POSSIBLE_QUALITIES} default="fast")
	priority = option({POSSIBLE_PRIORITIES} default="high")
"""
confspec = ConfigObj(StringIO(configSpecString), list_values=False, encoding="UTF-8")
confspec.newlines = "\r\n"

config.conf.spec["ocr"] = confspec


class OCRSettingsPanel(gui.SettingsPanel):

	# Translators: Title of the OCR settings dialog in the NVDA settings.
	title = _("OCR settings")

	def makeSettings(self, settingsSizer):
		sHelper = gui.guiHelper.BoxSizerHelper(self, sizer = settingsSizer)
		# Translators: Label of a radiobox used to optimize recognition for speed/quality
		recogQualityLabel = _("During OCR, prefer")
		self.recogQualityRB = sHelper.addItem(
			wx.RadioBox(
				self, label=recogQualityLabel, choices=[x.translatedName for x in OCR_QUALITIES.values()]
			)
		)
		select = [k for k, v in OCR_QUALITIES.items() if v.dirName == config.conf["ocr"]["quality"]][0]
		self.recogQualityRB.SetSelection(select)
		self.recogQualityRB.Bind(wx.EVT_RADIOBOX, self.onQualityChange)
		# Translators: Label of a  combobox used to choose a recognition language
		recogLanguageLabel = _("Recognition &language")
		self.recogLanguageCB = sHelper.addLabeledControl(
			recogLanguageLabel,
			wx.Choice,
			choices=[],
			style = wx.CB_SORT
		)
		self.updateCB()
		# Translators: Label of a  combobox used to choose a recognition priority
		recogPriorityLabel = _("Recognition &priority")
		self.recogPriorityCB = sHelper.addLabeledControl(
			recogPriorityLabel,
			wx.Choice,
			choices=[x.translatedName for x in OCR_PRIORITIES.values()]
		)
		priorityID = config.conf["ocr"]["priority"]
		select = self.recogPriorityCB.FindString(OCR_PRIORITIES[priorityID].translatedName)
		self.recogPriorityCB.SetSelection(select)

	def onQualityChange(self, evt):
		self.updateCB()

	def updateCB(self):
		self.recogLanguageCB.Set(
			[lang.localizedName for lang in LanguageInfo.fromAvailableLanguages(
				OCR_QUALITIES[self.recogQualityRB.GetSelection()].dirName
			)]
		)
		select = self.recogLanguageCB.FindString(LanguageInfo.fromConfiguredLanguage().localizedName)
		if select == wx.NOT_FOUND:
			select = self.recogLanguageCB.FindString(LanguageInfo.fromFallbackLanguage().localizedName)
		self.recogLanguageCB.SetSelection(select)

	def onSave(self):
		qualityIndex = self.recogQualityRB.GetSelection()
		config.conf["ocr"]["quality"] = OCR_QUALITIES[qualityIndex].dirName
		ocrLanguage = LanguageInfo(localizedName=self.recogLanguageCB.GetStringSelection())
		config.conf["ocr"]["language"] = ocrLanguage.TesseractLocaleName
		priorityString = self.recogPriorityCB.GetStringSelection()
		priorityID = [k for k, v in OCR_PRIORITIES.items() if v.translatedName == priorityString][0]
		config.conf["ocr"]["priority"] = priorityID


class GlobalPlugin(globalPluginHandler.GlobalPlugin):

	def __init__(self):
		super(globalPluginHandler.GlobalPlugin, self).__init__()
		gui.NVDASettingsDialog.categoryClasses.append(OCRSettingsPanel)

	def terminate(self):
		gui.NVDASettingsDialog.categoryClasses.remove(OCRSettingsPanel)

	@scriptHandler.script(
		# Translators: Input help mode message for the script used to recognize current navigator object.
		description = _("Recognizes current navigator object using Tesseract OCR. After recognition is done thext can be reviewed with review cursor commands."),
		gesture="kb:NVDA+r"
		)
	def script_ocrNavigatorObject(self, gesture):
		nav = api.getNavigatorObject()
		# Translators: Announced when object on which recognition is performed is not visible.
		cannotRecognizeMSG = _("Object is not visible.")
		try:
			left, top, width, height = nav.location
		except TypeError:
			ui.message(cannotRecognizeMSG)
			return
		if left < 0 or top < 0 or width <= 0 or height <= 0:
			ui.message(cannotRecognizeMSG)
			return
		bmp = wx.EmptyBitmap(width, height)
		mem = wx.MemoryDC(bmp)
		mem.Blit(0, 0, width, height, wx.ScreenDC(), left, top)
		img = bmp.ConvertToImage()
		# Tesseract copes better if we convert to black and white...
		img = img.ConvertToGreyscale()
		# and increase the size.
		img = img.Rescale(
			width * IMAGE_RESIZE_FACTOR,
			height * IMAGE_RESIZE_FACTOR,
			quality=wx.IMAGE_QUALITY_BICUBIC
		)
		baseFile = os.path.join(tempfile.gettempdir(), "nvda_ocr")
		try:
			imgFile = baseFile + ".bmp"
			img.SaveFile(imgFile)
			# Translators: Announced when recognition starts.
			ui.message(_("Running OCR"))
			ocrLang = config.conf["ocr"]["language"]
			ocrQualityDir = config.conf["ocr"]["quality"]
			langArg = '/'.join([ocrQualityDir, ocrLang])
			priorityID = config.conf["ocr"]["priority"]
			priorityArg = OCR_PRIORITIES[priorityID].priorityConstant
			# Hide the Tesseract window.
			si = subprocess.STARTUPINFO()
			si.dwFlags = subprocess.STARTF_USESHOWWINDOW
			si.wShowWindow = subprocess.SW_HIDE
			# If NVDA is attached to a console window Tesseract release info its written to this console.
			# Stdout cannot be unconditionally redirected to null however, as it breaks  when focused window is not a console.
			if api.getFocusObject().windowClassName == "ConsoleWindowClass":
				redirecStdoutTo = subprocess.DEVNULL
			else:
				redirecStdoutTo = None
			try:
				subprocess.check_call(
					(TESSERACT_EXE, imgFile, baseFile, "--tessdata-dir", TESSDATA_BASEDIR, "-l", langArg, "hocr"),
					startupinfo=si,
					stdout=redirecStdoutTo,
					creationflags=priorityArg
				)
			except subprocess.CalledProcessError as e:
				log.info(e)
				# Translators: error message when OCR fails
				ui.message(_("Error during OCR, please see log (chosen language files are present?)"))
		finally:
			try:
				os.remove(imgFile)
			except OSError:
				pass
		try:
			hocrFile = baseFile + ".hocr"
			parser = HocrParser(open(hocrFile, encoding='utf8').read(), left, top)
		finally:
			try:
				os.remove(hocrFile)
			except OSError:
				pass
		if parser.textLen == 0:
			# Translators: Announced when OCR process succeeded, but no text was recognized.
			ui.message(_("No text found."))
			return
		# Let the user review the OCR output.
		# TextInfo of the navigator object cannot be overwritten dirrectly as this makes it impossible to navigate with the caret in edit fields.
		# Create a shallow copy of the navigator object and overwrite there.
		objWithResults = copy(nav)
		objWithResults.makeTextInfo = lambda position: OcrTextInfo(objWithResults, position, parser)
		api.setReviewPosition(objWithResults.makeTextInfo(textInfos.POSITION_FIRST))
		# Translators: Announced when recognition is finished.
		ui.message(_("Done"))

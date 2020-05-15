# -*- coding: UTF-8 -*-
"""NVDA OCR plugin
This plugin uses Tesseract for OCR: https://github.com/tesseract-ocr
It also uses Pillow: https://python-pillow.org/
@author: James Teh <jamie@nvaccess.org>
@author: Rui Batista <ruiandrebatista@gmail.com>
@copyright: 2011-2020 NV Access Limited, Rui Batista, Łukasz Golonka
@license: GNU General Public License version 2.0
"""

import sys
import os
import tempfile
import subprocess
from xml.parsers import expat
from collections import namedtuple
from copy import copy
import wx
import config
import globalPluginHandler
import gui
import api
from logHandler import log
import languageHandler
import addonHandler
addonHandler.initTranslation()
import textInfos.offsets
import ui
import locationHelper
import scriptHandler

PLUGIN_DIR = os.path.dirname(__file__)
TESSERACT_EXE = os.path.join(PLUGIN_DIR, "tesseract", "tesseract.exe")

# Pillow requires pathlib which is not bundled with NVDA
# Therefore place it in the plugin directory and add it temporarily to PYTHONPATH
sys.path.append(PLUGIN_DIR)
from .PIL import ImageGrab
from .PIL import Image
del sys.path[-1]

IMAGE_RESIZE_FACTOR = 2

OcrWord = namedtuple("OcrWord", ("offset", "left", "top"))

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
			elif cls == "ocr_word":
				# Get the coordinates from the bbox info specified in the title attribute.
				title = attrs.get("title")
				prefix, l, t, r, b = title.split(" ")
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

class OCRSettingsPanel(gui.SettingsPanel):

	# Translators: Title of the OCR settings dialog in the NVDA settings.
	title = _("OCR settings")

	def makeSettings(self, settingsSizer):
		sHelper = gui.guiHelper.BoxSizerHelper(self, sizer = settingsSizer)
		# Translators: Label of a  combobox used to choose a recognition language
		recogLanguageLabel = _("Recognition &language")
		self.availableLangs = {languageHandler.getLanguageDescription(tesseractLangsToLocales[lang]) or tesseractLangsToLocales[lang] : lang for lang in getAvailableTesseractLanguages()}
		self.recogLanguageCB = sHelper.addLabeledControl(
			recogLanguageLabel,
			wx.Choice,
			choices = list(self.availableLangs.keys()),
			style = wx.CB_SORT
		)
		tessLangsToDescs = {v : k  for k, v in self.availableLangs.items()}
		curlang = config.conf["ocr"]["language"]
		try:
			select = tessLangsToDescs[curlang]
		except ValueError:
			select = tessLangsToDescs['eng']
		select = self.recogLanguageCB.FindString(select)
		self.recogLanguageCB.SetSelection(select)

	def onSave (self):
		lang = self.availableLangs[self.recogLanguageCB.GetStringSelection()]
		config.conf["ocr"]["language"] = lang

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
		left, top, width, height = nav.location
		img = ImageGrab.grab(bbox=(left, top, left + width, top + height))
		# Tesseract copes better if we convert to black and white...
		img = img.convert(mode='L')
		# and increase the size.
		img = img.resize((width * IMAGE_RESIZE_FACTOR, height * IMAGE_RESIZE_FACTOR), Image.BICUBIC)
		baseFile = os.path.join(tempfile.gettempdir(), "nvda_ocr")
		try:
			imgFile = baseFile + ".bmp"
			img.save(imgFile)
			# Translators: Announced when recognition starts.
			ui.message(_("Running OCR"))
			lang = config.conf["ocr"]["language"]
			# Hide the Tesseract window.
			si = subprocess.STARTUPINFO()
			si.dwFlags = subprocess.STARTF_USESHOWWINDOW
			si.wShowWindow = subprocess.SW_HIDE
			subprocess.check_call(
				(TESSERACT_EXE, imgFile, baseFile, "-l", lang, "hocr"),
				startupinfo=si,
				stdout=subprocess.DEVNULL
			)
		finally:
			try:
				os.remove(imgFile)
			except OSError:
				pass
		try:
			hocrFile = baseFile + ".html"
			parser = HocrParser(open(hocrFile,encoding='utf8').read(),
				left, top)
		finally:
			try:
				os.remove(hocrFile)
			except OSError:
				pass
		# Let the user review the OCR output.
		# TextInfo of the navigator object cannot be overwritten dirrectly as this makes it impossible to navigate with the caret in edit fields.
		# Create a shallow copy of the navigator object and overwrite there.
		objWithResults = copy(nav)
		objWithResults.makeTextInfo = lambda position: OcrTextInfo(nav, position, parser)
		api.setReviewPosition(objWithResults.makeTextInfo(textInfos.POSITION_FIRST))
		# Translators: Announced when recognition is finished, note that it is not guaranteed that some text has been found.
		ui.message(_("Done"))

localesToTesseractLangs = {
"bg" : "bul",
"ca" : "cat",
"cs" : "ces",
"zh_CN" : "chi_tra",
"da" : "dan",
"de" : "deu",
"el" : "ell",
"en" : "eng",
"fi" : "fin",
"fr" : "fra",
"hu" : "hun",
"id" : "ind",
"it" : "ita",
"ja" : "jpn",
"ko" : "kor",
"lv" : "lav",
"lt" : "lit",
"nl" : "nld",
"nb_NO" : "nor",
"pl" : "pol",
"pt" : "por",
"ro" : "ron",
"ru" : "rus",
"sk" : "slk",
"sl" : "slv",
"es" : "spa",
"sr" : "srp",
"sv" : "swe",
"tg" : "tgl",
"tr" : "tur",
"uk" : "ukr",
"vi" : "vie"
}
tesseractLangsToLocales = {v : k for k, v in localesToTesseractLangs.items()}

def getAvailableTesseractLanguages():
	dataDir = os.path.join(os.path.dirname(__file__), "tesseract", "tessdata")
	dataFiles = [file for file in os.listdir(dataDir) if file.endswith('.traineddata')]
	return [os.path.splitext(file)[0] for file in dataFiles]

def getDefaultLanguage():
	lang = languageHandler.getLanguage()
	if lang not in localesToTesseractLangs and "_" in lang:
		lang = lang.split("_")[0]
	return localesToTesseractLangs.get(lang, "eng")

configspec = {
	"language" : f"string(default={getDefaultLanguage()})"
}

config.conf.spec["ocr"] = configspec

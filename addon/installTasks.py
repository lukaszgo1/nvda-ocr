import re
import os
import config
import globalVars

def onInstall():
	oldOCRConfig = os.path.join(globalVars.appArgs.configPath, "ocr.ini")
	try:
		with open(oldOCRConfig, "r", encoding = "utf-8") as f:
			configContent = f.read()
		langFinder = r'^language \= (.+)$'
		if re.match(langFinder, configContent):
			OCRLang = re.match(langFinder, configContent)[1]
			# Ensure that language is written to the default configuration
			if "ocr" not in config.conf.profiles[0]:
				config.conf.profiles[0]["ocr"] = {}
			config.conf.profiles[0]["ocr"]["language"] = OCRLang
		os.remove(oldOCRConfig)
	except FileNotFoundError:
		pass

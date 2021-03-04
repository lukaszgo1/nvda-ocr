"""Download and extract dependencies.
"""

import os
import urllib.request
import fnmatch
import shutil
import subprocess

DEPS_URLS = {
	"7zip": "https://7-zip.org/a/7z1900.msi",
	"tesseract": "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w32-setup-v5.0.0-alpha.20201127.exe"
}
ROOT_DIR = os.path.abspath(os.path.dirname(__file__))
DEPS_DIR = os.path.join(ROOT_DIR, "deps")
PLUGIN_DIR = os.path.join(ROOT_DIR, "addon", "globalPlugins")

depFiles = set()
dep7Zip = None

def downloadDeps():
	try:
		os.mkdir(DEPS_DIR)
	except OSError:
		pass
	print("Downloading dependencies")
	for prg,url in DEPS_URLS.items():
		fn = os.path.basename(url)
		localPath = os.path.join(DEPS_DIR, fn)
		depFiles.add(localPath)
		if os.path.isfile(localPath):
			print(("%s already downloaded" % fn))
			continue
		print("Downloading %s" % fn)
		# Download to a temporary path in case the download aborts.
		tempPath = localPath + ".tmp"
		urllib.request.urlretrieve(url, tempPath)
		os.rename(tempPath, localPath)

def extract7Zip():
	for zfn in depFiles:
		if fnmatch.fnmatch(zfn, "*/7z*.msi"):
			break
	else:
		assert False
	msiDir = os.path.join(DEPS_DIR, "7zip")
	print("Extracting 7Zip")
	shutil.rmtree(msiDir, ignore_errors=True)
	process = subprocess.Popen(["msiexec.exe", "/a", zfn, "/qn", "TARGETDIR=%s"%msiDir])
	process.wait()
	exeDir = os.path.join(msiDir, "Files", "7-Zip")
	global dep7Zip
	dep7Zip = os.path.join(exeDir, "7z.exe")
	if not os.path.isfile(dep7Zip):
		assert False

TESSERACT_FILES = {
	"bin": ["tesseract.exe", "*.dll"],
	"doc": ["doc/AUTHORS", "doc/LICENSE"],
	"tessdata": ["tessdata/configs/*", "tessdata/tessconfigs/*", "tessdata/*.traineddata", "tessdata/*.user-*", "tessdata/pdf.ttf"]
}

def extractTesseract():
	for zfn in depFiles:
		if fnmatch.fnmatch(zfn, "*/tesseract-ocr-*32-*"):
			break
	else:
		assert False
	tessDir = os.path.join(PLUGIN_DIR, "tesseract")
	shutil.rmtree(tessDir, ignore_errors=True)
	for outDir,includeFiles in TESSERACT_FILES.items():
		print("Extracting Tesseract %s"%outDir)
		includeSwitches = ['-i!%s'%f for f in includeFiles]
		if outDir == "bin":
			outDirPath = os.path.join(tessDir, outDir)
		else:
			outDirPath = tessDir
		process = subprocess.Popen([dep7Zip, "x", *includeSwitches, "-aoa", "-bd", "-bso0", "-spe", "-o%s"%outDirPath, zfn])
		process.wait()
	print("Adjusting traineddata folders...")
	tessdataDir = os.path.join(tessDir, "tessdata")
	os.mkdir(os.path.join(tessdataDir, "fast"))
	files = os.listdir(tessdataDir)
	fastDir = os.path.join(tessdataDir, "fast")
	for file in fnmatch.filter(files, "*.traineddata"):
		if not file.startswith("osd."):
			filePath = os.path.join(tessdataDir, file)
			shutil.move(filePath, fastDir)
	os.mkdir(os.path.join(tessdataDir, "best"))

def main():
	downloadDeps()
	extract7Zip()
	extractTesseract()
	print("Done!")

if __name__ == "__main__":
	main()

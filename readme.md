# OCR

* Authors: NV Access Limited & other contributors. Currently maintained by Łukasz Golonka <lukasz.golonka@mailbox.org>
* NVDA compatibility: 2019.3 and beyond
* Download [Stable version][1]
* Download [version compatible with NVDA 2019.2 and older][2]

Important: if you are using NVDA 2017.3 or later on Windows 10, please consider using buit-in Windows 10 OCR.
Performs optical character recognition (OCR) to extract text from an object which is inaccessible. The Tesseract OCR engine is used. To perform OCR, move to the object in question using object navigation and press NVDA+r. You can set the OCR recognition language by going to the NVDA settings panel  and selecting OCR settings. The keyboard shortcut can be reassigned from NVDA input gestures dialog in the "Miscellaneous" category.

## Change log:

### Changes for 2.1:

* When recognition is performed on an edit field it is now possible to navigate it with the system caret without moving focus from it and back again.
* Tesseract release info is no longer written to the focused console.
* When no text is recognized this fact is announced to the user.
* It is no longer possible to attempt recognition on an invisible object.
* WX is now used to capture images instead of Pillow.

### Changes for 2.0:

* Compatibility with NVDA 2019.3 and later.
* It is possible to set different recognition languages for different configuration profiles.
* OCR settings can now be  changed from NVDA settings dialog rather than from a separate dialog in the tools menu.




[1]: https://addons.nvda-project.org/files/get.php?file=ocr
[2]: https://www.nvaccess.org/files/nvda-addons/ocr_0.20120529.01.nvda-addon


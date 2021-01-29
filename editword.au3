#include <MsgBoxConstants.au3>
#include <Word.au3>
#include <ScreenCapture.au3>

Run("notepad.exe")
; Wait for Notepad to become active - by waiting explicitly on a Window with the specified title
Local $notepadHWnd = WinWaitActive("Untitled - Notepad")

Local $tempScreencaptureFileName = @MyDocumentsDir & "\screenshot.jpg"
; Capture full screen
; _ScreenCapture_Capture($tempScreencaptureFileName )
; add coordinates to specify a region : starting x,y and width and height, for example  , 0, 0, 200, 500)
; to capture a specific window:
_ScreenCapture_CaptureWnd ( $tempScreencaptureFileName ,$notepadHWnd )

; create a new Word application object
Local $oWord = _Word_Create()
; Add a new empty document
Local $oDoc = _Word_DocAdd($oWord)
; position ourselves at the end of the document (slightly redundant in a fresh, new and empty document)
Local $oRange = _Word_DocRangeSet($oDoc, -1, Default, 4, Default, 4)
; add screenshot from file into new Word document at the "range" location
_Word_DocPictureAdd($oDoc, $tempScreencaptureFileName , Default, Default, $oRange)
; remove the temporary screenshot file
FileDelete (  $tempScreencaptureFileName )
; inform the user about our success!
MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocAdd Example", "A new empty document has successfully been added.")

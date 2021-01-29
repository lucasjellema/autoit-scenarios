; Run the Notepad application
Run("notepad.exe")

; Wait for Notepad to become active - by waiting explicitly on a Window with the specified title
WinWaitActive("Untitled - Notepad")

; Now that the Notepad window is active type some text
Send("Hello World{ENTER}Live from Notepad{ENTER}")
Sleep(2000)

; press Ctrl + Shift + S - the short cut for the File | Save As menu option
Send("^S")
; type ALT N (!n) to activate the file name entry field
; followed by the file name and add ALT+S (the ! causes the ALT key to be pressed)to activate the Save button (with short cut key ALT S)
Send("!nhelloworld.txt!S")
Sleep(2000)
; close Notepad.exe - by closing the window; note: the window can no longer be accessed through its original title
; there are several ways to identify a window - the title is but one; however, when multiple Notepad windows are running, this way would not suffice
WinClose("[CLASS:Notepad]", "")
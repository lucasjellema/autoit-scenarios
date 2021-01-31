; invoke this script with two command line parameters:
; 1: channel indication (for example BBC ONE, NPO 1) - default: BBC ONE
; 2: duration of recording in minutes - default 1
#include <AutoItConstants.au3>
#include <File.au3>
#include <ScreenCapture.au3>

Local $logFileHandle = FileOpen(@ScriptDir & "\recordZiggoChannel.log", 1)

Main() 

Func Main()
   Local $defaultRecordingTime = 1
   Local $recordingTime
   Local $defaultChannel = "BBC ONE"
   Local $channel
   If $CmdLine[0] > 0 Then
      lg("First commandline parameter " & $CmdLine[1] )
      lg("Second commandline parameter " & $CmdLine[2] )
      $channel  = $CmdLine[1]
   Else
      $channel = $defaultChannel
   EndIf

   If $CmdLine[0] > 1 Then
      $recordingTime  = $CmdLine[2]
   Else
      $recordingTime = $defaultRecordingTime
   EndIf

   lg("runSandbox to record from " & $channel & " during " & $recordingtime & " minute(s)")
   Local $sandboxWindowHandle = runSandbox()
   sc("SandboxRunning")
   lg("after runSandbox - handle to window " & hWnd)
   Sleep(6000)
   lg("Run Edge")
   runEdge()
   sc("Edge Running")
   Sleep(5000)
   lg("runZiggo")
   runZiggo($channel)
   sc("ziggo running")
   ; scale Sandbox Window to almost my entire screen
   Sleep(3000)
   lg("Scale Windows Sandbox")
   WinMove ( $sandboxWindowHandle, "", 5, 5, 1500, 950)
   ; maximize
   Sleep(500)
   MouseClick($MOUSE_CLICK_LEFT, 1423, 18)
   ; restore
   Sleep(500)
   MouseClick($MOUSE_CLICK_LEFT, 985 , 16)
   sc("SandboxMaximized")
   ; time to run the screen recorder for the window "Windows Sandbox"
   lg("Run OBS Recorder")
   Sleep(500)
   runOBSRecorder( $recordingTime)
   ; close the Windows Sandbox
   lg("after OBS Recorder; close Sandbox")
   WinClose($sandboxWindowHandle)
   Local $confirmationWnd = WinWait("Windows Sandbox", "Are you sure you want to close Windows Sandbox?")
   ; confirm closing the window - click on OK
   MouseClick($MOUSE_CLICK_LEFT, 823, 504)
   lg("Sandbox closed")
   ; release logfile
   FileClose($logFileHandle) ; Close the filehandle to release the file.
EndFunc

Func runOBSRecorder( $recordingTimeInMinutes)
   Local $obsAppLink = "C:\Program Files\obs-studio\bin\64bit\obs64.exe"
   Local $obsWorkingDir = "C:\Program Files\obs-studio\bin\64bit"
   ; Run OBS
   Run($obsAppLink, $obsWorkingDir)

   ; wait for OBS window
   Local $obsWindowTitle = "OBS Studio 26.1.1 (64-bit, windows) - Profile: Untitled - Scenes: Untitled"
   Local $obsWnd = WinWait($obsWindowTitle )
   ; select scene 7 - predefined with window capture of Windows Sandbox Window
   MouseClick($MOUSE_CLICK_LEFT, 28,728)
   Sleep(500)

   ; start recording
   MouseClick($MOUSE_CLICK_LEFT, 1353 , 665)

   ; sleep for as long as the program lasts
   Sleep(1000 * 60 *  $recordingTimeInMinutes)

   ; stop recording
   MouseClick($MOUSE_CLICK_LEFT, 1353 , 665)
   ; close OBS
   Sleep(2000)
   WinClose($obsWnd)
EndFunc


Func runSandbox()
   Const $sandboxAppId =  "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsSandbox.exe"

   ; run the Windows Sandbox
;   ShellExecute('shell:Appsfolder\' & $sandboxAppId )
;Run(@ComSpec & ' /c start "" "shell:Appsfolder\' & $sandboxAppId & '"', '', @SW_HIDE)

   ShellExecute('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Windows Sandbox.lnk')

   Local $sandboxWindowTitle = "Windows Sandbox"
   lg("wait for window " & $sandboxWindowTitle)
   Sleep(10000)
   Local $hWnd = WinWait($sandboxWindowTitle)
   lg("Window grabbed")
   return $hWnd
EndFunc

Func runEdge()
   ; run Edge
   ; mouse to Edge shortcut and double click
   Sleep(8000)
   MouseClick($MOUSE_CLICK_LEFT, 200,250, 2)
   ; maximize Edge
   Sleep(2000)
   MouseClick($MOUSE_CLICK_LEFT, 638, 132)
EndFunc




Func loginZiggo()
   ; login Ziggo
   MouseClick($MOUSE_CLICK_LEFT, 1234,206)
   ; wait for login window
   Sleep(1800)
   Send("username{TAB}")
   Sleep(1400)
   Send("password") ; note: escape special characters such as +, {,}, ! and ^ (see https://www.autoitscript.com/autoit3/docs/appendix/SendKeys.htm)
   Sleep(2500)
   ; click login button
   MouseClick($MOUSE_CLICK_LEFT, 756,520)
   ; let login take place
   Sleep(2000)
EndFunc

Func runZiggo($channel)
   Local $ziggoURL = "https://www.ziggogo.tv/nl.html{enter}"
   ; enter URL
   Sleep(1500)
   ; click on location bar
   MouseClick($MOUSE_CLICK_LEFT, 400,160)
   Send("{DEL}")
   Sleep(500)
   Send($ziggoURL)
   Sleep(5000)
   ; remove cookies warning
   MouseClick($MOUSE_CLICK_LEFT, 1332, 298)
   Sleep(300)
   loginZiggo()
   MouseClick($MOUSE_CLICK_LEFT, 756,520)
   ; get rid of edge password popup
   Sleep(2300)
   MouseClick($MOUSE_CLICK_LEFT, 1100, 640)

   ; I had problems with the first login attempt. With a second login sequence, that problem seemed to go away. It is a bit clumsy, but for now it does the job

   ; TV & Replay
   Sleep(4000)
   MouseClick($MOUSE_CLICK_LEFT, 204 , 384)
   Sleep(2000)
   MouseClick($MOUSE_CLICK_LEFT, 810,451)

   Sleep (3500)

   MouseClick($MOUSE_CLICK_LEFT, 400,160)
   Send("{DEL}")
   Sleep(500)
   Send($ziggoURL)
   Sleep(5000)

   Sleep(3000)
   loginZiggo()
   ; remove cookies warning
   MouseClick($MOUSE_CLICK_LEFT, 1332, 298)

   ; TV & Replay
   Sleep(3400)
   MouseClick($MOUSE_CLICK_LEFT, 204 , 384)
   Sleep(4000)
   ; DETERMINE the CHANNEL TO BRING UP
   ; click in channel search
   MouseClick($MOUSE_CLICK_LEFT, 1134, 307)
   Sleep(400)
   ; type in the channel search field the $channel 
   Send($channel)
   ; allow some time for the channel selection to be loaded
   Sleep(2000)
   ; click on first selected channel
   MouseClick($MOUSE_CLICK_LEFT, 442,451)

   ; it takes a while before the channel is playing; when we expect it to be, we want to go to full screen mode
   Sleep (13500)
   ; click anywhere to bring up controls
   MouseClick($MOUSE_CLICK_LEFT, 1060,440)
   Sleep(800)
   MouseClick($MOUSE_CLICK_LEFT, 1332,208)
   Sleep(800)
   ; press okay in response to Edge question on full screen; we are certainly happy with that
   MouseClick($MOUSE_CLICK_LEFT, 1029, 684)

EndFunc



Func lg($logText)
   _FileWriteLog($logFileHandle, $logText)
EndFunc

Func sc($captureName)
   Local $tempScreencaptureFileName = @MyDocumentsDir & "\" & $captureName & ".jpg"
   ; Capture full screen and write to file system
  _ScreenCapture_Capture($tempScreencaptureFileName )
EndFunc


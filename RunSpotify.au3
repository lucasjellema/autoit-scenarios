# run Spotify app
# https://www.autoitscript.com/forum/topic/187123-win10-run-windows-store-apps/
#include <AutoItConstants.au3>

Local $spotifyAppId =  "SpotifyAB.SpotifyMusic_zpdnekdrzrea0!Spotify"
Local $defaultPlaylist = "something quiet"
Local $playlist
If $CmdLine[0] > 0 Then
    $playlist = $CmdLine[1]
Else
    $playlist = $defaultPlaylist
EndIf

# two ways
ShellExecute('shell:Appsfolder\' & $spotifyAppId)

#Run(@ComSpec & ' /c start "" "shell:Appsfolder\' & $spotifyAppId & '"', '', @SW_HIDE)

Local $spotifyWindowTitle = "Spotify Premium"
Local $hWnd = WinWaitActive($spotifyWindowTitle)

# mouse to search field
Sleep(2000)
MouseClick($MOUSE_CLICK_LEFT, 308,22)

# type  name playlist
Send($playlist)

# wait for result to appear then mouse to playlist and click to play
Sleep(1500)
#MouseMove(307,215)
MouseClick($MOUSE_CLICK_LEFT, 307,215)




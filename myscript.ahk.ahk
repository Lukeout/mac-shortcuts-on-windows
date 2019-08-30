#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
*tab:: 
{   
    if (GetKeyState("LAlt", "P") AND GetKeyState("LShift", "P") = false) {     
        Send {LControl up}{LAlt down}{tab}
        KeyWait, tab  
    } else if (GetKeyState("LAlt", "P") AND GetKeyState("LShift", "P")) {     
        Send {LControl up}{LShift down}{LAlt down}{tab}
        KeyWait, tab
    } else if (GetKeyState("LCtrl", "P") AND GetKeyState("LShift", "P") = false) {     
        Send {LAlt up}{LCtrl down}{tab}
        KeyWait, tab
    } else if (GetKeyState("LCtrl", "P") AND GetKeyState("LShift", "P")) {  
        Send {LAlt up}{LShift down}{LCtrl down}{tab}
        KeyWait, tab
    } else if (GetKeyState("LWin", "P") AND GetKeyState("LShift", "P") = false) {     
        Send {LWin down}{tab}
        KeyWait, tab
    } else if (GetKeyState("LWin", "P") AND GetKeyState("LShift", "P")) {  
        Send {LShift down}{LWin down}{tab}
        KeyWait, tab
    } else {   
        send {tab}
    }      
    return
}

~LAlt Up::
{   
    send {LAlt up}
    return
}

~LCtrl Up::
{   
    send {LCtrl up}
    return
}

LAlt::LCtrl 
LCtrl::LAlt

;Maps ctrl-q to close
^q::Send !{F4}
return

;This is for long click screen shot to screen shot folder

#F18:: 
IfWinExist ahk_class screenClass
{
Send {esc}
return
}
IfWinActive ahk_exe Powerpnt.exe
{
Send {F5}
return
}
else
{
send {LWin Down}{PrintScreen}{LWin Up} 
sleep 500 
return
}

;This is for double click, screen shots snippet to OneNote2016. In Porpoint it goes forward
#F19::

IfWinExist ahk_class screenClass
{
Send {Left} 
return
}
else
{
 Send, {escape down}{escape up}
     SetKeyDelay, 0 
     SendInput, #+s
return
}

;This is for single click, Opens One Note 2016 and brings it the front if already open. Goes backwards in power point.
#F20::
IfWinExist ahk_class screenClass
{
Send {Space} 
return
}
IfWinNotExist ahk_exe OneNote.exe
{
Run C:\Program Files (x86)\Microsoft Office\root\Office16\OneNote.exe
return
}
IfWinActive ahk_exe OneNote.exe
{
sleep 50
    send {alt}   ;Activate menu hotkeys
    sleep 70   ;Waits for menu to activate, you may need to tweak this
    send d   ;Select draw
    sleep 10   ;Wait
    send p   ;Select pens
    return
}
IfWinExist ahk_exe OneNote.exe
{
WinActivate
return
}




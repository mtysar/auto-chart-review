#Requires AutoHotkey v1.1
#SingleInstance Force
SetTitleMatchMode, 2

; Win+B to trigger the workflow
#b::
    ; Wait for user to physically release Win key
    KeyWait, LWin
    KeyWait, RWin
    Sleep, 200

    ; 1) Clear clipboard, then copy selected cell in Excel
    Clipboard :=
    Sleep, 100
    SendPlay, ^c
    ClipWait, 2
    if (ErrorLevel)
    {
        ; Fallback to regular Send if SendPlay didn't work for copy
        Send, ^c
        ClipWait, 2
    }
    if (ErrorLevel)
    {
        MsgBox, Failed to copy from Excel. Try again.
        return
    }
    Sleep, 200

    ; Store MRN from clipboard before switching
    MRN := Clipboard

    ; 2) Switch to EPIC (Hyperspace via Citrix)
    if WinExist("HYPERSPACE")
        WinActivate
    else
    {
        MsgBox, Could not find EPIC/Hyperspace window. Please open EPIC and try again.
        return
    }
    WinWaitActive,,3
    Sleep, 1000

    ; Click the window to ensure Citrix has keyboard focus
    Click
    Sleep, 500

    ; 3) Add patient (Ctrl+4) - try SendPlay for Citrix
    SendPlay, ^4
    Sleep, 2000

    ; 4) Paste MRN into the patient search field
    Clipboard := MRN
    Sleep, 100
    SendPlay, ^v
    Sleep, 500

    ; 5) Double Enter to confirm/search
    SendPlay, {Enter}
    Sleep, 800
    SendPlay, {Enter}
    Sleep, 2000

    ; 6) Ctrl+Shift + type "oncology" (navigator search)
    SendPlay, ^+oncology
    Sleep, 500

return

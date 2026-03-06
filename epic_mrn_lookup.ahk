#Requires AutoHotkey v1.1
#SingleInstance Force
#InstallKeybdHook
#UseHook
SetTitleMatchMode, 2

; Win+B to trigger the workflow
#b::
    KeyWait, LWin
    KeyWait, RWin
    KeyWait, b
    Sleep, 200

    ; 1) Copy from Excel
    ToolTip, Step 1: Copying MRN... (1s)
    Clipboard :=
    Sleep, 100
    Send, ^c
    ClipWait, 2
    if (ErrorLevel)
    {
        ToolTip
        MsgBox, Failed to copy from Excel.
        return
    }
    MRN := Clipboard
    ToolTip, Step 1 DONE: MRN = %MRN%
    Sleep, 500

    ; 2) Switch to EPIC
    ToolTip, Step 2: Switching to EPIC...
    WinActivate, HYPERSPACE
    Sleep, 500
    ToolTip, Step 2: Clicking for focus...
    Click
    Sleep, 500
    ToolTip, Step 2 DONE

    ; 2b) Ctrl+W to close last patient
    Sleep, 300
    ToolTip, Step 2b: Closing last patient (Ctrl+W)...
    DllCall("keybd_event", "UChar", 0x11, "UChar", 0x1D, "UInt", 0, "UPtr", 0)
    Sleep, 50
    DllCall("keybd_event", "UChar", 0x57, "UChar", 0x11, "UInt", 0, "UPtr", 0)
    Sleep, 50
    DllCall("keybd_event", "UChar", 0x57, "UChar", 0x11, "UInt", 0x02, "UPtr", 0)
    Sleep, 50
    DllCall("keybd_event", "UChar", 0x11, "UChar", 0x1D, "UInt", 0x02, "UPtr", 0)
    ToolTip, Step 2b: Waiting... (1s)
    Sleep, 1000

    ; 3) Ctrl+4 Add patient
    ToolTip, Step 3: Sending Ctrl+4... (2s)
    DllCall("keybd_event", "UChar", 0x11, "UChar", 0x1D, "UInt", 0, "UPtr", 0)
    Sleep, 50
    DllCall("keybd_event", "UChar", 0x34, "UChar", 0x05, "UInt", 0, "UPtr", 0)
    Sleep, 50
    DllCall("keybd_event", "UChar", 0x34, "UChar", 0x05, "UInt", 0x02, "UPtr", 0)
    Sleep, 50
    DllCall("keybd_event", "UChar", 0x11, "UChar", 0x1D, "UInt", 0x02, "UPtr", 0)
    ToolTip, Step 3: Ctrl+4 sent. Waiting... (2s)
    Sleep, 1000
    ToolTip, Step 3: Ctrl+4 sent. Waiting... (1s)
    Sleep, 1000

    ; 4) Paste MRN
    ToolTip, Step 4: Pasting MRN... (1s)
    Clipboard := MRN
    Sleep, 200
    DllCall("keybd_event", "UChar", 0x11, "UChar", 0x1D, "UInt", 0, "UPtr", 0)
    Sleep, 50
    DllCall("keybd_event", "UChar", 0x56, "UChar", 0x2F, "UInt", 0, "UPtr", 0)
    Sleep, 50
    DllCall("keybd_event", "UChar", 0x56, "UChar", 0x2F, "UInt", 0x02, "UPtr", 0)
    Sleep, 50
    DllCall("keybd_event", "UChar", 0x11, "UChar", 0x1D, "UInt", 0x02, "UPtr", 0)
    Sleep, 500

    ; 5) Double Enter
    ToolTip, Step 5: Enter #1... (1s)
    DllCall("keybd_event", "UChar", 0x0D, "UChar", 0x1C, "UInt", 0, "UPtr", 0)
    Sleep, 50
    DllCall("keybd_event", "UChar", 0x0D, "UChar", 0x1C, "UInt", 0x02, "UPtr", 0)
    Sleep, 800
    ToolTip, Step 5: Enter #2...
    DllCall("keybd_event", "UChar", 0x0D, "UChar", 0x1C, "UInt", 0, "UPtr", 0)
    Sleep, 50
    DllCall("keybd_event", "UChar", 0x0D, "UChar", 0x1C, "UInt", 0x02, "UPtr", 0)
    ToolTip, Step 5: Patient loading... (5s)
    Sleep, 1000
    ToolTip, Step 5: Patient loading... (4s)
    Sleep, 1000
    ToolTip, Step 5: Patient loading... (3s)
    Sleep, 1000
    ToolTip, Step 5: Patient loading... (2s)
    Sleep, 1000
    ToolTip, Step 5: Patient loading... (1s)
    Sleep, 1000

    ; 6) Ctrl+Space to open navigator, then type "oncology" + Enter
    WinActivate, HYPERSPACE
    Sleep, 300
    Click
    Sleep, 500

    ToolTip, Step 6: Sending Ctrl+Space...
    DllCall("keybd_event", "UChar", 0x11, "UChar", 0x1D, "UInt", 0, "UPtr", 0)
    Sleep, 50
    DllCall("keybd_event", "UChar", 0x20, "UChar", 0x39, "UInt", 0, "UPtr", 0)
    Sleep, 50
    DllCall("keybd_event", "UChar", 0x20, "UChar", 0x39, "UInt", 0x02, "UPtr", 0)
    Sleep, 50
    DllCall("keybd_event", "UChar", 0x11, "UChar", 0x1D, "UInt", 0x02, "UPtr", 0)

    ToolTip, Step 6: Waiting for navigator... (2s)
    Sleep, 1000
    ToolTip, Step 6: Waiting for navigator... (1s)
    Sleep, 1000

    ; Type "oncology" as plain keystrokes
    ToolTip, Step 6: Typing oncology...
    vks := [0x4F, 0x4E, 0x43, 0x4F, 0x4C, 0x4F, 0x47, 0x59]
    Loop, 8
    {
        vk := vks[A_Index]
        DllCall("keybd_event", "UChar", vk, "UChar", 0, "UInt", 0, "UPtr", 0)
        Sleep, 30
        DllCall("keybd_event", "UChar", vk, "UChar", 0, "UInt", 0x02, "UPtr", 0)
        Sleep, 30
    }

    ; Press Enter to confirm
    Sleep, 300
    DllCall("keybd_event", "UChar", 0x0D, "UChar", 0x1C, "UInt", 0, "UPtr", 0)
    Sleep, 50
    DllCall("keybd_event", "UChar", 0x0D, "UChar", 0x1C, "UInt", 0x02, "UPtr", 0)

    ToolTip, DONE!
    Sleep, 1500
    ToolTip

return

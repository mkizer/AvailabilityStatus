; ----------------------------------------------------------------------------
; Written using AutoHotKey version: 1.1.32.00
; Author: Michael Kizer <michael.kizer@perspecta.com>
; Date:   2020-06-28
;
; Script Function:
; Prompts user for an availability status and sets it for Skype, Teams, and Chime.
; Also accepts command line arguments to set status.
;
; Todo:
; Optional input box for a status message for Skype/Teams.
; Also, add canned messages in the ini file. Use a combo box to select them.
;
; Changelog:
; v1.0 - 2020-06-30 - Initial release
; ----------------------------------------------------------------------------

#SingleInstance force

IniRead, UseSkype, AvailabilityStatus.ini, services, skype, true
IniRead, UseTeams, AvailabilityStatus.ini, services, teams, false
IniRead, UseChime, AvailabilityStatus.ini, services, chime, false

; Retrieve and validate command line parameters
NumParam = %0%
; MsgBox Number of parameters is %NumParam%

if NumParam > 0
{
    param1 = %1%
    ; MsgBox param1 is %param1%

    switch param1
    {
    case "/available":  Selection = 1
    case "/busy":       Selection = 2
    case "/dnd":        Selection = 3
    case "/away":       Selection = 4
    case "/brb":        Selection = 5
    case "/off":        Selection = 6
    default:
        MsgBox Invalid command line parameter. Valid values are: /available, /busy, /dnd, /away, /brb, /off
        ExitApp
    }
}

; If no command line parameters, show the GUI
if NumParam = 0
{
    Gui, New
    Gui, Add, Text,, Skype/Teams/Chime Availability Status
    Gui, Add, Text,, Select your availability:
    Gui, Add, Radio, vSelection Checked, Available
    Gui, Add, Radio, , Busy
    Gui, Add, Radio, , Do Not Disturb
    Gui, Add, Radio, , Away
    Gui, Add, Radio, , Be Right Back
    Gui, Add, Radio, , Off Work
    Gui, Add, Text, 
    Gui Add, Button, h25 w184 Default, OK
    Gui, Show,  w206 center
    return
}

ButtonOk:
if NumParam = 0
{
    Gui,Submit ;,Nohide      ;Remove Nohide if you want the GUI to hide.
}

; Set status section

; Skype
if (WinExist("ahk_class CommunicatorMainWindowClass") and %UseSkype% = true)
{
    WinActivate		; Automatically uses the window found above.

    switch Selection
    {
    case 1:
        Send !fmv   ; Available - Alt + f, then m, then v
        StatusText = Available
    case 2:
        Send !fmb   ; Busy - Alt + f, then m, then b
        StatusText = Busy
    case 3:
        Send !fmd   ; Do Not Disturb - Alt + f, then m, then d
        StatusText = Do Not Disturb
    case 4:
        Send !fma   ; Away - Alt + f, then m, then a
        StatusText = Away
    case 5:
        Send !fme   ; Be Right Back - Alt + f, then m, then e
        StatusText = Be Right Back
    case 6:
        Send !fmw   ; Off Work - Alt + f, then m, then w
        StatusText = Off Work
    }
    Send {Esc}
    TrayTip, Availability Status, Your Skype availability status was set to '%StatusText%', 20, 17
}	

; Teams
if (WinExist("ahk_exe Teams.exe") and %UseTeams% = true)
{
    WinActivate		; Automatically uses the window found above.
    Sleep, 1000
    ; Activate Search box
    Send ^e
    ; Select all previously entered text and delete
    Send ^a
    Send {Delete}
    Sleep, 1000

    switch Selection
    {
    case 1:
        SendInput /available
        Sleep, 1000
        SendInput {Enter}
        StatusText = Available
    case 2:
        SendInput /busy
        Sleep, 1000
        SendInput {Enter}
        StatusText = Busy
    case 3:
        SendInput /dnd
        Sleep, 1000
        SendInput {Enter}
        StatusText = Do Not Disturb
    case 4, 6:
        SendInput /away
        Sleep, 1000
        SendInput {Enter}
        StatusText = Away
    case 5:
        SendInput /brb
        Sleep, 1000
        SendInput {Enter}
        StatusText = Be Right Back
    }

    TrayTip, Availability Status, Your Teams availability status was set to '%StatusText%', 20, 17
}	

; Chime
; available, busy, do not disturb - the rest are automatic
; 
if (WinExist("ahk_exe Chime.exe") and %UseChime% = true)
{
    WinActivate		; Automatically uses the window found above.
    Send {Esc}
    CoordMode, Mouse, Window
    MouseClick, Left, 100, 100

    switch Selection
    {
    case 1:
        Send {Tab}{Enter}{Tab}{Tab}{Enter}
        StatusText = Available
    case 2:
        Send {Tab}{Enter}{Tab}{Tab}{Tab}{Enter}
        StatusText = Busy
    case 3:
        Send {Tab}{Enter}{Tab}{Tab}{Tab}{Tab}{Enter}
        StatusText = Do Not Disturb
    case 4, 5, 6:
        ; Chime doesn't have these options, so just reset the presences status to Automatic
        Send {Tab}{Enter}{Tab}{Enter}
        StatusText = Automatic
    }
    
    TrayTip, Availability Status, Your Chime availability status was set to '%StatusText%', 20, 17
}	

GuiEscape:
GuiClose:
ExitApp
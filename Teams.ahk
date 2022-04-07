SetTitleMatchMode, 2 ; 2 = a partial match on the title

; MICROSOFT TEAMS - Toggle Mute
+^!M::
WinGet, winid, ID, A	; Save the current window ID
if WinExist("(Meeting)") ;Yes, every Teams meeting has that in the title bar - even if it's not visible to you
{
	WinActivate ; Without any parameters this activates the previously retrieved window - in this case your meeting
	Send, ^+M   ; Teams' native Mute shortcut
        WinActivate ahk_id %winid% ; Restore previous window focus
	return
}


; MICROSOFT TEAMS - Accept Video Call
; May not actually work due to pop up/overlay rather than window
+^!A::
WinGet, winid, ID, A	; Save the current window ID
if WinExist("(Meeting)") ;Yes, every Teams meeting has that in the title bar - even if it's not visible to you
{
	WinActivate ; Without any parameters this activates the previously retrieved window - in this case your meeting
	Send, ^+A   ; Teams' native Accept Video Call shortcut
        WinActivate ahk_id %winid% ; Restore previous window focus
	return
}

; MICROSOFT TEAMS - Turn on WebCam
+^!O::
WinGet, winid, ID, A	; Save the current window ID
if WinExist("(Meeting)") ;Yes, every Teams meeting has that in the title bar - even if it's not visible to you
{
	WinActivate ; Without any parameters this activates the previously retrieved window - in this case your meeting
	Send, ^+O   ; Teams' native Webcam enable shortcut
        WinActivate ahk_id %winid% ; Restore previous window focus
	return
}

; MICROSOFT TEAMS - Share Desktop #1
+^!E::
WinGet, winid, ID, A	; Save the current window ID
if WinExist("(Meeting)") ;Yes, every Teams meeting has that in the title bar - even if it's not visible to you
{
	WinActivate ; Without any parameters this activates the previously retrieved window - in this case your meeting
	Send, ^+E   ; Teams' native Screen share shortcut
	Send {enter}
        WinActivate ahk_id %winid% ; Restore previous window focus
	return
}

; MICROSOFT TEAMS - Leave call/meeting
+^!H::
WinGet, winid, ID, A	; Save the current window ID
if WinExist("(Meeting)") ;Yes, every Teams meeting has that in the title bar - even if it's not visible to you
{
	WinActivate ; Without any parameters this activates the previously retrieved window - in this case your meeting
	Send, ^+H   ; Teams' native End call shortcut
        WinActivate ahk_id %winid% ; Restore previous window focus
	return
}
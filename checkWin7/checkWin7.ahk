; Recommended defaults
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
;SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; Hotkey to activate script
^`::

; Key delay after striking hotkey
SetKeyDelay, 500

; This script searches for asset tags from a spreadsheet to look for the operating system in ServiceNow

; To start, need both browser windows open on start page (maybe can automate that part later)
; Make spreadsheet window active
; Click on column where comments are located first

; Scrolls to the last filled entry in column
	Send, {ctrl down}
	Send, {down 2}
	Send, {ctrl up}

; Goes to the next blank comment
Send, {down}

/*
; Initialize loop success tally
loopSuccess := 0

; ----- GUI -----

Gui, new, , Win7 Inspector
Gui, Show, w200 h100
Gui +AlwaysOnTop -MaximizeBox -MinimizeBox
Gui, Show, x0 y0 ; Top left
Gui, Show, NA ; Shows without activating
Gui, Add, Text, vloopSuccess, Sucessful Runs: 
*/

Loop ,
{	
	; Intialize comment to paste into spreadsheet at end of assessment
	Comment := ""

	; Scrolls over to leftmost column (might take >1 if missing values, hence 5)
	Send, {ctrl down}
	Send, {left 5}

	; Copies asset tag	
	Send, c									; could bracket c... maybe... idk		; set key delay
	Send, {ctrl up}

	; Switches to ServiceNow window
	Send, {alt down}
	Send, {tab} 
	Send, {alt up}
	
	; Finds Asset Management page
	Loop,
	{
		ImageSearch, X, Y, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, Imgs\asset_management_link.png
		If (ErrorLevel = 0)
		{
			Click, %X%, %Y%
				break
		}
	}
	
	; Wait 0.4 seconds to load
	Sleep, 400

	; Finds search bar, then click into it
	Loop,
	{
		ImageSearch, X, Y, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, Imgs\search_box.png
		If (ErrorLevel = 0)
		{
			Click, %X%, %Y%
				break
		}
	}

	; Pastes asset tag into search bar
	Send, {ctrl down}
	Send, v
	Send, {ctrl up}

	; Searches for asset tag
	Send, {enter}

	; Wait 0.4 seconds to load
	Sleep, 400

	; Finds link to click, then clicks it. If it can't find the image, ... meh
	lease := ""
	Loop,
	{
		ImageSearch, X, Y, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, Imgs\lease_link.png
		
		If (ErrorLevel = 0)
		{
			Y := Y + 90 ; maybe between 80-100 ???
			X := X + 5
			Click, %X%, %Y%
				break
		}
		
		If (ErrorLevel = 0)
		{
			Click, %X%, %Y%
			lease := "not lease"
			MsgBox,,, "Not a Lease!"
				break
		}
	}

	; Wait 2 seconds to load
	Sleep, 2000

	; If in stock and quarantined, make a note of it
	state := ""
	Loop, 5
	{
		ImageSearch, X, Y, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, Imgs\state_in_stock_quarantined.png
		
		If (ErrorLevel = 0)
		{
			state := "quarantined redeployed"
			MsgBox,,, "Quarantined!"
				break
		}
	}

	Send, {pgdn}
	Sleep, 2500 ; sleep 2.5 secs
	Send, {pgdn}

	; Scrolls down to ensure the hardware tab is selected
	Loop, 5
	{

		ImageSearch, X, Y, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, Imgs\hardware_tab.png
		
		If (ErrorLevel = 0)
		{
			Y := Y + 5
			X := X + 5
			Click, %X%, %Y%
				break
		}
	}

	; Clicks the asset tag link to view hardware
	Loop,
	{
		ImageSearch, X, Y, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, Imgs\hardware_asset_tag.png
		
		If (ErrorLevel = 0)
		{
			Y := Y + 65	; 50-60 px
			X := X + 60 ; 50-60 px
			Click, %X%, %Y%
				break
		}
	}

	Sleep, 1000 ; sleep a sec

	; Check if Win7
	os := "win7"

	Loop, 5
	{
		ImageSearch, X, Y, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, Imgs\os_win7.png
		
		If (ErrorLevel = 0)
		{
			os := "win7"
			MsgBox,,, "It's Win7. Ew."
				break
		}
	}

	; Check if Win10
	Loop, 5
	{
		ImageSearch, X, Y, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, Imgs\os_win10.png
		
		If (ErrorLevel = 0)
		{
			os := "win10"
				break
		}
	}

	Comment = %Comment% %lease% %state% %os% -mas
	MsgBox,,, %Comment%

	; Switches back to spreadsheet
	Send, {alt down}
	Send, {tab}
	Send, {alt up}

	; Scrolls over to rightmost column (might take >1 if missing values, hence 5)
	Send, {ctrl down}
	Send, {right 5}
	Send, {ctrl up}

	; Set Clipboard to Comment
	Clipboard = %Comment%

	; Pastes status into comment column
	Send, {ctrl down}
	Send, v
	Send, {ctrl up}

	; Move down to next comment
	Send, {down}
	
	loopSuccess += 1
	GuiControl,, loopSuccess, Successful Runs: %loopSuccess%
} Until Clipboard = ""
/*
Works up to this point! :) Testing complete.
*/
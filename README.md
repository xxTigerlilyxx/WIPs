How to use Data EXCELerator: 

This Excel utility is useful for when you must manually copy and paste things from the internet or other documents into Excel. Instead of clicking back and forth between your several open windows, just use this to cut out that annoying middle step. The Range Highlighter feature helps you see where your active cell is currently located when you are not actively in Excel. The Range Viewer feature helps you double check what cells you are selecting and the data you are populating into it.

|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
|||||||||||||||||||||||||||||||||||||   Excel Controls   ||||||||||||||||||||||||||||||||||||||||				
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

ALT+C = Copy currently highlighted selection and paste into Excel as plain text into single cells

ALT+V = Copy currently highlighted selection and paste array into Excel as rich text into multiple cells

ALT+UP = Move up and select next cell(s) above active cell(s)

ALT+DOWN = Move down and select next cell(s) below active cell(s)

ALT+LEFT = Move left and select next cell(s) the the left of active cell(s)

ALT+RIGHT = Move right and select next cell(s) the right of active cell(s)

ALT+ENTER = Opens an input box you can type in specific text to send to the active cell(s)

|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
|||||||||||||||||||||||||||||||||||||||||    Features     ||||||||||||||||||||||||||||||||||||||||||				
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

ALT+SHIFT+R = Toggles on and off Range Viewer feature, which displays Range Address, Value, and Sum next to mouse

ALT+SHIFT+H  = Toggles on and off Range Highlighter feature, which helps you see what cell you are selecting when active in a different window (this feature temporarily disables Excel's undo capability when ON)

|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
|||||||||||||||||||||||||||||||||||||   Shutting Down  ||||||||||||||||||||||||||||||||||||||	
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

ALT+END = Closes down Data EXCELerator within 5 seconds

|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
|||||||||||||||||||||||||||||||||||||||||||     Notes     ||||||||||||||||||||||||||||||||||||||||||||				
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

• Copy/Paste operations cannot be undone with Data EXCELerator - make backups!
• After use, Copy/Paste operations will not allow you to use Excel's Undo feature to go back in time
• Range Highlighter may leave traces of yellow fill if you click around with it ON, especially with large selections
• Do not type directly into Excel when Range Highlighter is on, it may turn the cell yellow (but you may like that)
• It's hard to explain the different between the two Copy/Paste operations, but they are both useful
• Planning to add more features in the future, like Paste Formatting/Values only and highlighting multiple cells like you would with SHIFT+ARROWKEY. I'm open to any suggestions you may have!

• Thanks for using Data EXCELerator! Have fun playing around with it!



;[][][][][][][][][][][][][][][][] Data EXCELerator [][][][][][][][][][][][][][][][]; 

;[][][][][][][][][][][][][][][][] Auto Execute Section [][][][][][][][][][][][][][][][]; 

#SingleInstance, force ; Doesn't allow the script to run multiple instances at once.
#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases. 
SendMode Input ; Recommended for new scripts due to its superior speed and reliability. 
DetectHiddenWindows, On ; Detects Hidden Windows that are invisible running in background
RangeViewer := 0 ; 0 = Display Active Cell Range & Value next to mouse (ON), 1 = Don’t display (OFF) RetryCheck := 0
RetryCheck := 0 ; Resets Retry Timer variable to 0 tries - 15 tries = close invisible Excel instance
SelectionHighlight := 0 ; 0 = Highlights the active cell when not in Excel window (ON), 1 = Selection Highlighter feature OFF

;[][][][][][][][][][][][][][][][] Connect To Excel [][][][][][][][][][][][][][][][]; 

MsgBox, 0x2040, Data EXCELerator,
(
How to use Data EXCELerator: 

This Excel utility is useful for when you must manually copy and paste things from the internet or other documents into Excel. Instead of clicking back and forth between your several open windows, just use this to cut out that annoying middle step. The Range Highlighter feature helps you see where your active cell is currently located when you are not actively in Excel. The Range Viewer feature helps you double check what cells you are selecting and the data you are populating into it.

|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
|||||||||||||||||||||||||||||||||||||   Excel Controls   ||||||||||||||||||||||||||||||||||||||||				
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

ALT+C = Copy currently highlighted selection and paste into Excel as plain text into single cells

ALT+V = Copy currently highlighted selection and paste array into Excel as rich text into multiple cells

ALT+UP = Move up and select next cell(s) above active cell(s)

ALT+DOWN = Move down and select next cell(s) below active cell(s)

ALT+LEFT = Move left and select next cell(s) the the left of active cell(s)

ALT+RIGHT = Move right and select next cell(s) the right of active cell(s)

ALT+ENTER = Opens an input box you can type in specific text to send to the active cell(s)

|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
|||||||||||||||||||||||||||||||||||||||||    Features     ||||||||||||||||||||||||||||||||||||||||||				
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

ALT+SHIFT+R = Toggles on and off Range Viewer feature, which displays Range Address, Value, and Sum next to mouse

ALT+SHIFT+H  = Toggles on and off Range Highlighter feature, which helps you see what cell you are selecting when active in a different window (this feature temporarily disables Excel's undo capability when ON)

|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
|||||||||||||||||||||||||||||||||||||   Shutting Down  ||||||||||||||||||||||||||||||||||||||	
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

ALT+END = Closes down Data EXCELerator within 5 seconds

|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
|||||||||||||||||||||||||||||||||||||||||||     Notes     ||||||||||||||||||||||||||||||||||||||||||||				
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

• Copy/Paste operations cannot be undone with Data EXCELerator - make backups!
• After use, Copy/Paste operations will not allow you to use Excel's Undo feature to go back in time
• Range Highlighter may leave traces of yellow fill if you click around with it ON, especially with large selections
• Do not type directly into Excel when Range Highlighter is on, it may turn the cell yellow (but you may like that)
• It's hard to explain the different between the two Copy/Paste operations, but they are both useful
• Planning to add more features in the future, like Paste Formatting/Values only and highlighting multiple cells like you would with SHIFT+ARROWKEY. I'm open to any suggestions you may have!

• Thanks for using Data EXCELerator! Have fun playing around with it!
)

;[][][][][][][][][][][][][][][][] Connect To Excel [][][][][][][][][][][][][][][][]; 

{
RetryCheck: ; Code below uses this "label" to troubleshoot hidden Excel windows
Loop, ; Repeats the code below it indefinitely in a loop
{


try { Xl := ComObjActive("Excel.Application") ; Try to connect to Active Excel instance
WinGet, Xl_pId, PID, % "ahk_id " Xl.HWND ; Get unique process ID for Active Excel instance
Xl_Name := Xl.ActiveWorkbook.Name ; Sets workbook name to variable for troubleshooting hidden Excel instances

if (SelectionHighlight = 0) { ; If Selection Highlighter feature is on... Highlight the selection without ruining cell formatting
Xl_Selection := Xl.Selection, NewAddress := Xl_Selection.Address[0,0], OriginalFormat := Xl.Range(NewAddress).Interior.ColorIndex, Xl.Range(NewAddress).Interior.ColorIndex := 36, Xl.Range(NewAddress).Interior.ColorIndex := OriginalFormat 
}
if  (Xl_Name != "") { ; If workbook name DOES NOT equal an empty string... Reset Retry Timer
		if (RangeViewer = 0) { ; If Range Viewer is ON...
		ToolTip % "Range: " Xl.Selection.Address[0,0] "`n`nValue: " Xl.Selection.Text "`n`nSum: " Round(Xl.WorksheetFunction.Sum(Xl.Selection), 2) ; Display Range & Value	
		}
RetryCheck := 0 ; Resets Retry Timer variable back to 0

} else if  (Xl_Name = "") && (RetryCheck < 15) { ; If workbook name doesn’t exist, try for 15 more seconds
ToolTip ; Stop displaying Range & Value next to mouse
Sleep, 1000 ; Wait 1,000 milliseconds
RetryCheck += 1 ; Add 1 second to retry timer variable
Goto, RetryCheck ; Try to connect to Excel again (will try 15 times over 15 seconds)

} else if  (Xl_Name = "") && (RetryCheck = 15) { ; if unsuccessful after 15 tries and name doesn’t exist
Process, Close, %Xl_pId% ; close invisible Excel instance hiding in background
RetryCheck := 0 ; Resets Retry Timer variable to 0 tries
} } catch e { ; if an error happens in this process, ignore it instead of popping up an error message
} }  ; Close the ends of every right-facing curly bracket with the same amount of left facing curlies

;[][][][][][][][][][][][][][][][] Hotkey Section [][][][][][][][][][][][][][][][]; 

!c:: ; 	Sets the COPY/PASTE Text to ONE cell Hotkey as ALT+C - everything below this line of code will execute
{ ClipSaved := ClipboardAll ; Saves data currently stored in your clipboard so you can restore it later
Clipboard = ; Clears the contents of your clipboard to ensure better reliability with ClipWait
Send, ^c ; Makes your keyboard type CTRL+C to copy highlighted text to your clipboard
ClipWait ; Waits until data is fully loaded into clipboard
try { Xl.Selection.Value := Clipboard ; Pastes your loaded clipboard into your current selection in Excel
} catch e { ; if an error happens in this process, ignore it instead of popping up an error message
} Clipboard := ClipSaved ; Restores the data that was stored in your clipboard prior to this operation
} return ; Tells the Hotkey to step executing lines of code right here

!v::
{ ClipSaved := ClipboardAll ; Saves data currently stored in your clipboard so you can restore it later
Clipboard = ; Clears the contents of your clipboard to ensure better reliability with ClipWait
Send, ^c ; Makes your keyboard type CTRL+C to copy highlighted text to your clipboard
ClipWait ; Waits until data is fully loaded into clipboard
try { Xl.Selection.PasteSpecial() ; Pastes your loaded clipboard into your current selection in Excel
} catch e { ; if an error happens in this process, ignore it instead of popping up an error message
} Clipboard := ClipSaved ; Restores the data that was stored in your clipboard prior to this operation
} return ; Tells the Hotkey to step executing lines of code right here

!UP:: ; Sets the Move Up Hotkey as ALT+UP ARROW - everything below this line of code will execute
{ try { Xl.Selection.Offset(-1,0).Select ; Try to select the cell immediately above the active cell
} catch e { ; if an error happens in this process, ignore it instead of popping up an error message 
}} return ; Tells the Hotkey to step executing lines of code right here

!DOWN:: ; Sets the Move Down Hotkey as ALT+DOWN ARROW - everything below this line of code will execute
{ try { Xl.Selection.Offset(1,0).Select ; Try to select the cell immediately below the active cell
} catch e { 
}} return

!LEFT:: ; Sets the Move Left Hotkey as ALT+LEFT ARROW - everything below this line of code will execute
{ try { Xl.Selection.Offset(0,-1).Select ; Try to select the cell immediately left of the active cell
} catch e { 
}} return

!RIGHT:: ; Sets the Move Right Hotkey as ALT+RIGHT ARROW - everything below this line of code will execute
{ try { Xl.Selection.Offset(0,1).Select ; Try to select the cell immediately right of the active cell
} catch e { 
}} return

!ENTER:: ; Sets the Enter Value Hotkey as ALT+ENTER - everything below this line of code will execute
{ InputBox,TypeToExcel, What would you like to enter into Excel?, Type whatever you want to go into the active range in excel: ; Displays input box - Stores submission into a variable to send to Excel
if ErrorLevel {  ; If user presses Cancel or...
} else if (TypeToExcel = "") { ; if user enters nothing and presses OK, script will do nothing
} else { ; If user types in something and presses OK...
try { Xl.Selection.Value := TypeToExcel ; try to put what user typed into the Active Cell in Excel
} catch e { ; If there is an error transferring the data to Excel...
MsgBox Data transfer failed - either you need to exit from edit mode (you may be editing inside of a cell) or Excel is not open. ; Open a Message Dialog Box to notify the User why and next steps
} } } return ; Tells the Hotkey to step executing lines of code right here	

!+r:: ; Sets the Toggle Range Viewer Hotkey as ALT+SHIFT+T - everything below this line will execute
{ if (RangeViewer = 0) { ; If Range Viewer is ON...
RangeViewer := 1 ; Then set it’s variable to OFF aka “1”
ToolTip ; Turn off Range Viewer
} else if (RangeViewer = 1) ; If Range Viewer is OFF...
RangeViewer := 0 ; Then set it’s variable to ON aka “0”
} return ; Tells the Hotkey to step executing lines of code right here

!+h:: ; Sets the Toggle Range Highlighter Hotkey as ALT+SHIFT+H - everything below this line will execute
{ if (SelectionHighlight = 0) { ; If Range Highlighter is ON...
SelectionHighlight := 1 ; Then set it’s variable to OFF aka “1”
} else if (SelectionHighlight = 1) ; If Range Highlighter is OFF...
SelectionHighlight := 0 ; Then set it’s variable to ON aka “0”
} return ; Tells the Hotkey to step executing lines of code right here

!End:: ; Sets the Close Program Hotkey as ALT+END - everything below this line will execute
{ ToolTip ; Turn off Range Viewer feature
MsgBox, 0x2030, Shutting Down The Data EXCELerator, Data EXCELerator will close within 5 seconds., 5 ; Notify user that Data EXCELerator.ahk is closing
ExitApp  ; Shut down / Close / Exit Data EXCELerator.ahk program
}

}

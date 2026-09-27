# Read from Excel
tell application "Microsoft Excel"
	if exists sheet "PraatSettings" then
		# Read the settings
		tell worksheet "PraatSettings" of active workbook
			set praatOption1 to value of cell "C1"
			set praatOption2 to value of cell "C2"
			set praatOption3 to value of cell "C3"
			set excelOption1 to value of cell "C4"
			set excelOption2 to value of cell "C5"
			set excelOption3 to value of cell "C6"
			set excelOption4 to value of cell "C7"
			# Sound file path
			set pathToSound to praatOption1
			# TextGrid file path
			set pathToTextGrid to praatOption2
			# Sound object type
			if praatOption3 is "LongSound" then
				set longSound to 1
			else if praatOption3 is "Sound" then
				set longSound to 0
			end if
			# Data columns (basename, start, end)
			set basenameColID to my convertAlphabetToNumber(excelOption1)
			set startTimeColID to my convertAlphabetToNumber(excelOption2)
			set endTimeColID to my convertAlphabetToNumber(excelOption3)
			if basenameColID is startTimeColID or basenameColID is endTimeColID or startTimeColID is endTimeColID then
				display alert "An error has occurred." message "Columns for basename, start time, and end time must be different from each other in the \"PraatSettings\" sheet." as critical
				error number -128
			end if
			# Row color after launch
			set rowColor to excelOption4
		end tell
	else
		# Create the settings sheet
		set theResponse to display dialog "The \"PraatSettings\" sheet is not found. Would you like to create it?" buttons {"Cancel", "Create"} default button "Create" with title "Praat Launcher for Mac"
		set theResponse to button returned of theResponse
		if theResponse is "Create" then
			make new worksheet at the beginning of active workbook
			set name of active sheet to "PraatSettings"
			tell active sheet
				set column width of range "A:A" to 10
				set column width of range "B:B" to 20
				set column width of range "C:C" to 70
				set color of interior object of range "A1:B3" to {255, 205, 205}
				set color of interior object of range "A4:B7" to {30, 110, 70}
				set color of interior object of range "C1:C7" to {255, 255, 205}
				set color of font object of range "A4:B7" to {255, 255, 255}
				set font style of font object of range "A1,A4" to "Bold"
				set horizontal alignment of range "B1:B7" to horizontal align right
				set horizontal alignment of range "C1:C7" to horizontal align left
				set vertical alignment of range "A1:C7" to vertical alignment top
				set locked of range "C1:C7" to false
				set wrap text of range "C1:C7" to true
				set value of cell "A1" to "Praat"
				set value of cell "A4" to "Excel"
				set value of cell "B1" to "Sound file:"
				set value of cell "B2" to "TextGrid file:"
				set value of cell "B3" to "Sound Object type:"
				set value of cell "B4" to "Column for basename:"
				set value of cell "B5" to "Column for start time:"
				set value of cell "B6" to "Column for end time:"
				set value of cell "B7" to "Color of launched row:"
				set value of cell "C1" to "/Users/xxx/<basename>.wav"
				set value of cell "C2" to "/Users/xxx/<basename>.TextGrid"
				set value of cell "C3" to "Sound"
				set value of cell "C4" to "A"
				set value of cell "C5" to "B"
				set value of cell "C6" to "C"
				set value of cell "C7" to "Blue"
				border around (range "A1:C3") line style double
				border around (range "A4:C7") line style double
				add data validation (validation of range "C3") type validate list formula1 "Sound,LongSound"
				add data validation (validation of range "C4:C6") type validate list formula1 "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
				add data validation (validation of range "C7") type validate list formula1 "Blue,Green,Gray,Red,Unchanged"
				set randomPassword to ""
				repeat 10 times
					set randomPassword to randomPassword & some item of "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
				end repeat
				protect worksheet password randomPassword
			end tell
		end if
		error number -128
	end if
	# Read the selected row
	if name of active sheet of active workbook is not "PraatSettings" then
		set selectedRowID to the first row index of active cell
		# Get the workbook folder (for resolving relative paths)
		set workbookFolder to missing value
		try
			set wbFull to full name of active workbook
			if wbFull is not missing value then
				set AppleScript's text item delimiters to "/"
				set pathParts to text items of (wbFull as text)
				if (count of pathParts) > 1 then set workbookFolder to (items 1 thru -2 of pathParts) as text
				set AppleScript's text item delimiters to ""
			end if
		end try
		tell active sheet
			# Get the basename and times
			set basename to value of cell selectedRowID of column basenameColID of active sheet
			# Convert numeric basename to text
			if class of basename is in {integer, real} then
				if basename = (basename div 1) then
					set basename to (basename div 1) as text
				else
					set basename to basename as text
				end if
			end if
			set startTime to value of cell selectedRowID of column startTimeColID of active sheet
			set endTime to value of cell selectedRowID of column endTimeColID of active sheet
			# Insert the basename into the paths
			set pathToSound to my findAndReplaceInText(pathToSound, "<basename>", basename)
			set pathToTextGrid to my findAndReplaceInText(pathToTextGrid, "<basename>", basename)
			# Resolve relative paths against the workbook folder
			set pathToSound to my resolvePath(pathToSound, workbookFolder)
			set pathToTextGrid to my resolvePath(pathToTextGrid, workbookFolder)
			# Check whether the files exist
			set openSound to 0
			set openTextGrid to 0
			try
				do shell script "test -f " & quoted form of pathToSound
				set openSound to 1
			end try
			try
				do shell script "test -f " & quoted form of pathToTextGrid
				set openTextGrid to 1
			end try
			if openSound is 0 and openTextGrid is 0 then
				display alert "An error has occurred." message "\"" & pathToSound & "\" not found." & "\"" & pathToTextGrid & "\" not found." as critical
				error number -128
			end if
			if startTime is "" and endTime is "" then
				set startTime to 0
				set endTime to 0
			else if startTime is "" then
				set startTime to endTime
			else if endTime is "" then
				set endTime to startTime
			end if
			try
				set startTime to startTime as number
				set endTime to endTime as number
				if startTime is greater than endTime then
					set reversedStartTime to endTime
					set reversedEndTime to startTime
					set startTime to reversedStartTime
					set endTime to reversedEndTime
				end if
			on error
				display alert "An error has occurred." message "Invalid time information." as critical
				error number -128
			end try
		end tell
	else
		display alert "Select a different sheet." message "Praat cannot be launched from the \"PraatSettings\" sheet." as critical
		error number -128
	end if
	# Color the launched row
	if rowColor is "Red" then
		set color of font object of entire row of active cell to {255, 0, 0}
	else if rowColor is "Green" then
		set color of font object of entire row of active cell to {0, 255, 0}
	else if rowColor is "Blue" then
		set color of font object of entire row of active cell to {0, 0, 255}
	else if rowColor is "Gray" then
		set color of font object of entire row of active cell to {130, 130, 130}
	end if
end tell

on convertAlphabetToNumber(input)
	set output to offset of input in "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
end convertAlphabetToNumber

on findAndReplaceInText(theText, theSearchString, theReplaceString)
	set AppleScript's text item delimiters to theSearchString
	set theTextItems to every text item of theText
	set AppleScript's text item delimiters to theReplaceString
	set theText to theTextItems as string
	set AppleScript's text item delimiters to ""
	return theText
end findAndReplaceInText

on resolvePath(thePath, baseFolder)
	if baseFolder is missing value then return thePath
	# Leave absolute (/...) and home (~...) paths alone; otherwise treat as
	# relative to the folder of the active workbook
	if thePath does not start with "/" and thePath does not start with "~" then
		if thePath starts with "./" then set thePath to text 3 thru -1 of thePath
		if baseFolder does not end with "/" then set baseFolder to baseFolder & "/"
		set thePath to baseFolder & thePath
	end if
	return thePath
end resolvePath

# Build the Praat script
set psEsc to my findAndReplaceInText(pathToSound, "\"", "\\\"")
set tgEsc to my findAndReplaceInText(pathToTextGrid, "\"", "\\\"")

set scriptText to "open_sound = " & openSound & "
path_to_sound$ = \"" & psEsc & "\"
long_sound = " & longSound & "
open_textgrid = " & openTextGrid & "
path_to_textgrid$ = \"" & tgEsc & "\"
start_time = " & startTime & "
end_time = " & endTime & "

if open_sound = 1
	if long_sound = 1
		sound = Open long sound file: path_to_sound$
	else
		sound = Read from file: path_to_sound$
	endif
	file_end = Get end time
endif
if open_textgrid = 1
	textGrid = Read from file: path_to_textgrid$
	file_end = Get end time
endif
editor_object = 0
if open_sound = 1 and open_textgrid = 1
	selectObject: sound, textGrid
	View & Edit
	editor_object = textGrid
elsif open_sound = 1 and open_textgrid = 0
	selectObject: sound
	if long_sound = 1
		View
	elsif long_sound = 0
		View & Edit
	endif
	editor_object = sound
elsif open_sound = 0 and open_textgrid = 1
	selectObject: textGrid
	View & Edit alone
	editor_object = textGrid
else
	exitScript()
endif
editor: editor_object
	Select: start_time, end_time
	if start_time <> end_time
		duration = end_time - start_time
		Zoom: start_time - duration/8, end_time + duration/8
	elsif start_time = end_time
		if start_time <> 0 and end_time <= file_end
			Zoom: start_time - 0.5, end_time + 0.5
		endif
	endif
endeditor
"

# Write it to a temporary file
set tmpFilePath to "/tmp/tmp.praat"
try
	set tmpFileRef to open for access POSIX file tmpFilePath with write permission
	set eof tmpFileRef to 0
	write scriptText to tmpFileRef starting at eof
	close access tmpFileRef
on error errMsg
	try
		close access POSIX file tmpFilePath
	end try
	display dialog errMsg
	error number -128
end try

# Launch Praat
# If Praat is not running yet, have LaunchServices start it with the script as
# its argument. "--send" runs the script in an existing Praat when there is one
# and otherwise becomes the GUI instance itself, running the script once its
# windows are up -- so nothing here has to wait for Praat to be ready. (The
# earlier approach, starting Praat and polling System Events for its windows,
# could stall for its full 15-second timeout when that polling was refused.)
# "open" hands the launch to LaunchServices, so the new Praat lives in the GUI
# session whatever process happens to host this script.
set praatWasRunning to (application "Praat" is running)
if not praatWasRunning then
	do shell script "open -a /Applications/Praat.app --args --send " & quoted form of tmpFilePath
	return
end if
tell application "Praat" to activate

# Hand the script over to the running Praat, by way of launchd.
# The detour matters: as an Automator Quick Action this script is hosted by
# com.apple.automator.runner, an XPC service with its own bootstrap namespace.
# A Praat started directly from there does see the running Praat ("An instance
# of Praat that is not me is already running.") but cannot reach it, and the
# script is dropped without a word. Asking launchd to start it puts it in the
# GUI session, where the connection works. The job deletes itself once the
# message has been delivered, so it is not run again.
set praatBin to "/Applications/Praat.app/Contents/MacOS/Praat"
set sendLabel to "com.praatlauncher.send"
set errFilePath to "/tmp/tmp.praat.log"
set sendCommand to praatBin & " --send " & quoted form of tmpFilePath & "; launchctl remove " & sendLabel
do shell script "rm -f " & quoted form of errFilePath & "; launchctl remove " & sendLabel & " > /dev/null 2>&1; launchctl submit -l " & sendLabel & " -o " & quoted form of errFilePath & " -e " & quoted form of errFilePath & " -- /bin/sh -c " & quoted form of sendCommand
delay 1
set praatMessage to ""
try
	set praatMessage to do shell script "grep -v 'is already running' " & quoted form of errFilePath
end try
if praatMessage is not "" then
	display alert "Praat reported a problem." message praatMessage as critical
end if

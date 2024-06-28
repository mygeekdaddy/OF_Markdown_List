(*
File: OmniFocus_Due_List.scpt
Summary: Create .md file for list of tasks due and deferred +/- 7d from current date.
-------------------------------------------------------------------------------------------------
Script based on Justin Lancy (@veritrope) from Veritrope.com
http://veritrope.com/code/write-todays-completed-tasks-in-omnifocus-to-a-text-file
-------------------------------------------------------------------------------------------------
Revision: 5.1
Revised: 2024-06-25
Rev Notes: Cleaned up old code comments and legacy date functions
-------------------------------------------------------------------------------------------------
*)

--Date string calcs to get current date in ISO format
set strYear to year of (current date) as integer

if (month of (current date) as integer) < 10 then
	set strMonth to "0" & (month of (current date) as integer)
else
	set strMonth to month of (current date) as integer
end if

if (day of (current date) as integer) < 10 then
	set strDay to "0" & (day of (current date) as integer)
else
	set strDay to day of (current date) as integer
end if

set str_date to "" & strYear & "-" & strMonth & "-" & strDay

-------------------------------------------------------------------------------------------------

--Set File/Path name of output Markdown file

-- Set output folder to Desktop
--set theFilePath to ((path to desktop folder) as string) & "Open Action Items - " & str_date & ".md"

-- Set output folder to Documents 
set theFilePath to ((path to documents folder) as string) & "Open Action Items - " & str_date & ".md"

-------------------------------------------------------------------------------------------------

--Get OmniFocus task list
set due_Tasks to my OmniFocus_TaskList()

--Output .MD text file
my write_File(theFilePath, due_Tasks)

--Set OmniFocus Due Task List
on OmniFocus_TaskList()
	set endDate to (current date) + (7 * days)
	set startDate to (current date) - (7 * days)
	set CurrDate to date (short date string of (startDate))
	-- set CurrDatetxt to short date string of date (short date string of (current date))
	set endDatetxt to date (short date string of (endDate))
	tell application "OmniFocus"
		tell default document
			set refDueTaskList to a reference to (flattened tasks where (due date > startDate and due date < endDatetxt and completed = false))
			set {lstName, lstProject, lstContext, lstDueDate} to {name, name of its containing project, name of its primary tag, due date} of refDueTaskList
			set strDueText to "Tasks Due this week: " & return & return
			repeat with iTask from 1 to count of lstName
				set {varTaskName, varProject, varContext, varDueDate} to {item iTask of lstName, item iTask of lstProject, item iTask of lstContext, item iTask of lstDueDate}
				if (varDueDate < (current date)) then
					set strDueDate to "<span style=\"color:red\">" & short date string of varDueDate & "</span>"
				else
					--set strDueDate to short date string of varDueDate
					set strDueDate to "<span style=\" font-weight: 450\">" & short date string of varDueDate & "</span>"
				end if
				set strDueText to strDueText & "▢ " & varTaskName & " " & strDueDate
				set strDueText to strDueText & return
			end repeat
		end tell
		
		tell default document
			set refDeferTaskList to a reference to (flattened tasks where (defer date < endDatetxt and (due date < CurrDate or due date is missing value) and completed = false))
			set {lstDeferName, lstDeferProject, lstDeferTag, lstDeferDate} to {name, name of its containing project, name of its primary tag, defer date} of refDeferTaskList
			set strDeferText to "<br />" & "<hr>" & space & return & space & return & "Tasks Deferred to start this week:" & return & return
			repeat with iDeferTask from 1 to count of lstDeferName
				set {varDeferTaskName, varDeferProj, varDeferTag, varDeferDate} to {item iDeferTask of lstDeferName, item iDeferTask of lstDeferProject, item iDeferTask of lstDeferTag, item iDeferTask of lstDeferDate}
				if (varDeferDate < (current date)) then
					set strDeferTask to "<span style=\"color:blue\">" & short date string of varDeferDate & "</span>"
				else
					set strDeferTask to "<span style=\" font-weight: 450\">" & short date string of varDeferDate & "</span>"
				end if
				set strDeferText to strDeferText & "▢ " & varDeferTaskName & " " & strDeferTask
				set strDeferText to strDeferText & return
			end repeat
		end tell
	end tell
	set strGlobalList to strDueText & strDeferText
	strGlobalList
end OmniFocus_TaskList

-------------------------------------------------------------------------------------------------

-- Writes collected task list .MD file
on write_File(theFilePath, due_Tasks)
	set theText to due_Tasks
	-- This following line will allow existing file to be overwritten if file already exists
	tell application "System Events" to if (exists file theFilePath) then delete file theFilePath
	set theFileReference to open for access theFilePath with write permission
	write theText to theFileReference as «class utf8»
	close access the theFileReference
end write_File

-------------------------------------------------------------------------------------------------

(*
Comment out remaining lines if you do not use Marked 2 as your MD viewer
*)

tell application "Marked 2"
	open file theFilePath
end tell

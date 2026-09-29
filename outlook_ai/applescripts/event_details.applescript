-- outlook-ai: full details for calendar events looked up by Outlook id (used
-- for recurring series that have occurrences in the requested window).
-- Usage: osascript event_details.applescript <id,id,...>
-- Records: EVT (same layout as list_events.applescript) | ERR | id | error text

on run argv
	my requireOutlook()
	set outList to {}
	repeat with idText in my splitText(item 1 of argv, ",")
		if (contents of idText) is not "" then
			set evId to (contents of idText) as integer
			try
				tell application "Microsoft Outlook"
					set ev to calendar event id evId
					get subject of ev
				end tell
				set end of outList to my eventRecord(ev, "")
			on error errMsg number errNum
				set end of outList to my makeRecord({"ERR", evId as text, my clean(errMsg) & " (" & errNum & ")"})
			end try
		end if
	end repeat
	return my joinList(outList, my sepRecord())
end run

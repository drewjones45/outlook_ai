-- outlook-ai: calendar events for a window, plus what Python needs to expand
-- recurring meetings. Outlook returns a recurring series as its "master" event
-- only (start = first occurrence), never the individual occurrences, so each
-- series' iCalendar text is returned for Python to expand. Full details for
-- the series that do occur in the window come from event_details.applescript.
-- Usage: osascript list_events.applescript <startLocalISO> <endLocalISO> [calendar names, one per line]
-- Records:
--   EVT | id | subject | start | end | allDay | location | organizer | calendar |
--         isRecurring | freeBusy | attendees | body | ical | isOccurrence |
--         recurrenceId | masterId            (events overlapping the window)
--   SER | id | calendar | ical            (every recurring series)
--   XCP | masterId | recurrenceId          (an edited occurrence whose original slot is in the window)
-- attendees: "name FS address FS type FS status" items joined by GS; type,
-- status and freeBusy are Outlook enum values as text (Python decodes them).

on run argv
	my requireOutlook()
	set startD to my dateFromISO(item 1 of argv)
	set endD to my dateFromISO(item 2 of argv)
	set wanted to {}
	if (count of argv) > 2 then
		if (item 3 of argv) is not "" then set wanted to paragraphs of (item 3 of argv)
	end if
	tell application "Microsoft Outlook"
		set cals to every calendar
	end tell
	set outList to {}
	repeat with c in cals
		set cal to contents of c
		set calName to ""
		tell application "Microsoft Outlook"
			try
				set calName to name of cal
			end try
		end tell
		if (count of wanted) is 0 or wanted contains calName then
			set outList to outList & my calendarRecords(cal, my clean(calName), startD, endD)
		end if
	end repeat
	return my joinList(outList, my sepRecord())
end run

on calendarRecords(cal, calName, startD, endD)
	set inWindow to {}
	set seriesList to {}
	set movedList to {}
	tell application "Microsoft Outlook"
		with timeout of 600 seconds
			try
				set inWindow to (every calendar event of cal whose start time < endD and end time > startD)
			end try
			try
				set seriesList to (every calendar event of cal whose is recurring is true and is occurrence is false)
			on error
				try
					set seriesList to (every calendar event of cal whose is recurring is true)
				end try
			end try
			try
				set movedList to (every calendar event of cal whose is occurrence is true and recurrence id is greater than or equal to startD and recurrence id < endD)
			end try
		end timeout
	end tell
	set recs to {}
	repeat with e in inWindow
		set end of recs to my eventRecord(contents of e, calName)
	end repeat
	repeat with e in seriesList
		set end of recs to my seriesRecord(contents of e, calName)
	end repeat
	repeat with e in movedList
		set end of recs to my exceptionRecord(contents of e)
	end repeat
	return recs
end calendarRecords

-- Just enough to expand a series: two Apple Events instead of dozens.
on seriesRecord(ev, calName)
	set evId to ""
	set icalText to ""
	tell application "Microsoft Outlook"
		with timeout of 120 seconds
			try
				set evId to (id of ev) as text
			end try
			try
				set icalText to icalendar data of ev
			end try
		end timeout
	end tell
	return my makeRecord({"SER", evId, calName, my clean(icalText)})
end seriesRecord

on exceptionRecord(ev)
	set masterId to ""
	set recId to ""
	tell application "Microsoft Outlook"
		try
			set mst to master of ev
			set masterId to (id of mst) as text
		end try
		try
			set recId to my isoDate(recurrence id of ev)
		end try
	end tell
	return my makeRecord({"XCP", masterId, recId})
end exceptionRecord

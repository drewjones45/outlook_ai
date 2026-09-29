-- outlook-ai: list every message in a folder as (id, time, read state).
-- Uses bulk property reads (one Apple Event per property for the whole
-- folder) instead of a per-message loop, so it stays fast on big folders.
-- Usage: osascript index_folder.applescript <inbox|sent> <accountName or ""> <folderName>
-- Records:
--   IDX | id | time (received for inbox, sent for sent) | isRead(1/0)

on run argv
	my requireOutlook()
	set folderRole to item 1 of argv
	set accountName to item 2 of argv
	set folderName to item 3 of argv
	set theFolder to my resolveFolder(folderRole, accountName, folderName)

	-- Script-object properties keep list access O(1) in the loop below.
	script o
		property idList : {}
		property idListAgain : {}
		property timeList : {}
		property readList : {}
		property outList : {}
	end script

	tell application "Microsoft Outlook"
		with timeout of 1800 seconds
			set o's idList to id of every message of theFolder
			if folderRole is "sent" then
				set o's timeList to time sent of every message of theFolder
			else
				set o's timeList to time received of every message of theFolder
			end if
			set o's readList to is read of every message of theFolder
			-- Re-read the ids: if mail arrived mid-index the lists may not line up.
			set o's idListAgain to id of every message of theFolder
		end timeout
	end tell

	set n to count of o's idList
	if n is not (count of o's timeList) or n is not (count of o's readList) or (o's idList) is not (o's idListAgain) then
		error "Folder changed while indexing" number 1002
	end if

	repeat with i from 1 to n
		set end of o's outList to my makeRecord({"IDX", (item i of o's idList) as text, my isoDate(item i of o's timeList), my boolText(item i of o's readList)})
	end repeat
	return my joinList(o's outList, my sepRecord())
end run

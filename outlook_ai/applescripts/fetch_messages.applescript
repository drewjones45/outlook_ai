-- outlook-ai: full details for a batch of messages, looked up by Outlook id.
-- Usage: osascript fetch_messages.applescript <id,id,...> <maxBodyChars>
-- Records:
--   MSG | id | subject | sender | received | sent | isRead | to | cc | flag |
--         priority | categories | attachments | headers | replied | truncated | body
--   ERR | id | error text
-- sender/to/cc are "name FS address" items joined by GS. flag and priority are
-- Outlook enum values as text (Python decodes them). headers is the raw header
-- block; Python keeps only the few it needs.

on run argv
	my requireOutlook()
	set maxChars to (item 2 of argv) as integer
	set outList to {}
	repeat with idText in my splitText(item 1 of argv, ",")
		if (contents of idText) is not "" then
			set end of outList to my messageRecord((contents of idText) as integer, maxChars)
		end if
	end repeat
	return my joinList(outList, my sepRecord())
end run

on messageRecord(msgId, maxChars)
	try
		tell application "Microsoft Outlook"
			with timeout of 120 seconds
				set m to message id msgId
				set subj to subject of m
			end timeout
		end tell
	on error errMsg number errNum
		return my makeRecord({"ERR", msgId as text, my clean(errMsg) & " (" & errNum & ")"})
	end try

	set senderText to ""
	set recvText to ""
	set sentText to ""
	set readText to "1"
	set toText to ""
	set ccText to ""
	set flagText to ""
	set prioText to ""
	set catList to {}
	set attList to {}
	set hdrText to ""
	set repliedText to ""
	set bodyText to ""
	set truncText to "0"

	tell application "Microsoft Outlook"
		with timeout of 120 seconds
			try
				set snd to sender of m
				set senderText to my addrPair(snd)
			end try
			try
				set recvText to my isoDate(time received of m)
			end try
			try
				set sentText to my isoDate(time sent of m)
			end try
			try
				set readText to my boolText(is read of m)
			end try
			try
				set toText to my recipientList(to recipients of m)
			end try
			try
				set ccText to my recipientList(cc recipients of m)
			end try
			try
				set flagValue to todo flag of m
				set flagText to my clean(flagValue)
			end try
			try
				set prioValue to priority of m
				set prioText to my clean(prioValue)
			end try
			try
				repeat with c in (categories of m)
					set end of catList to my clean(name of c)
				end repeat
			end try
			try
				repeat with a in (attachments of m)
					set end of attList to my clean(name of a)
				end repeat
			end try
			try
				set hdrText to my clean(headers of m)
			end try
			try
				set repliedText to "0"
				if (replied to of m) or (replied to all of m) then set repliedText to "1"
			on error
				set repliedText to ""
			end try
			try
				set bodyText to plain text content of m
			end try
			if bodyText is missing value then set bodyText to ""
			if bodyText is "" then
				-- Fall back to the HTML body; Python strips the markup.
				try
					set bodyText to content of m
				end try
			end if
		end timeout
	end tell

	set bodyText to my clean(bodyText)
	if maxChars > 0 and (length of bodyText) > maxChars then
		set bodyText to text 1 thru maxChars of bodyText
		set truncText to "1"
	end if
	return my makeRecord({"MSG", msgId as text, my clean(subj), senderText, recvText, sentText, readText, toText, ccText, flagText, prioText, my joinList(catList, my sepItem()), my joinList(attList, my sepItem()), hdrText, repliedText, truncText, bodyText})
end messageRecord

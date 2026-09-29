-- outlook-ai: save a reply to a message as a DRAFT. This script never sends.
-- Usage: osascript create_reply_draft.applescript <messageId> <htmlBody> <plainBody>
--          <replyAll 1/0> <openWindow 1/0> [my addresses, one per line]
-- Records:
--   DRAFT | draftId | subject | to | method | folder
-- Uses Outlook's own "reply to" (keeps threading and the quoted original). If
-- that fails or hangs, builds an equivalent new message instead.

on run argv
	my requireOutlook()
	set msgId to (item 1 of argv) as integer
	set htmlBody to item 2 of argv
	set plainBody to item 3 of argv
	set replyAll to ((item 4 of argv) is "1")
	set openIt to ((item 5 of argv) is "1")
	set myAddresses to {}
	if (count of argv) > 5 then
		if (item 6 of argv) is not "" then set myAddresses to paragraphs of (item 6 of argv)
	end if

	tell application "Microsoft Outlook"
		set original to message id msgId
		get subject of original -- fails fast (-1728) if the id is stale
	end tell

	set draftMsg to missing value
	set method to "reply"
	try
		tell application "Microsoft Outlook"
			with timeout of 60 seconds
				if replyAll then
					set draftMsg to reply to original with reply to all without opening window
				else
					set draftMsg to reply to original without opening window
				end if
			end timeout
		end tell
		if draftMsg is missing value then error "reply to returned nothing"
	on error errMsg number errNum
		set method to "new message (Outlook's reply command failed: " & errMsg & " " & errNum & ")"
		set draftMsg to my manualReply(original, htmlBody, replyAll, myAddresses)
	end try

	if method is "reply" then
		-- Setting content replaces the whole body, so prepend to what Outlook
		-- generated (the quoted original) instead of overwriting it.
		set existing to ""
		tell application "Microsoft Outlook"
			try
				set existing to content of draftMsg
			end try
		end tell
		if existing is missing value then set existing to ""
		set newContent to my composeBody(existing, htmlBody, plainBody)
		tell application "Microsoft Outlook"
			set content of draftMsg to newContent
		end tell
	end if

	-- Read everything back before opening a window: once open, Outlook
	-- refuses property reads on the draft.
	set draftId to ""
	set draftSubject to ""
	set toText to ""
	set folderName to ""
	tell application "Microsoft Outlook"
		try
			set draftId to (id of draftMsg) as text
		end try
		try
			set draftSubject to subject of draftMsg
		end try
		try
			set toText to my recipientList(to recipients of draftMsg)
		end try
		try
			set folderName to name of (folder of draftMsg)
		end try
		if openIt then open draftMsg
	end tell
	return my makeRecord({"DRAFT", draftId, my clean(draftSubject), toText, my clean(method), my clean(folderName)})
end run

-- Insert the new text right after <body ...> so Outlook's quoted original
-- stays below it. Plain-text drafts get the plain version prepended.
on composeBody(existing, htmlBody, plainBody)
	if existing is "" then return htmlBody
	set bodyPos to offset of "<body" in existing
	if bodyPos > 0 then
		set tailText to text bodyPos thru -1 of existing
		set closePos to offset of ">" in tailText
		if closePos > 0 then
			set cutPos to bodyPos + closePos - 1
			if cutPos is greater than or equal to (length of existing) then return existing & htmlBody
			return (text 1 thru cutPos of existing) & htmlBody & (text (cutPos + 1) thru -1 of existing)
		end if
	end if
	if existing contains "<html" or existing contains "<div" or existing contains "<p" or existing contains "<br" then
		return htmlBody & existing
	end if
	return plainBody & existing
end composeBody

on manualReply(original, htmlBody, replyAll, myAddresses)
	tell application "Microsoft Outlook"
		set subj to ""
		try
			set subj to subject of original
		end try
		if subj is missing value then set subj to ""
		if not (subj starts with "RE:") then set subj to "RE: " & subj
		set acct to missing value
		try
			set acct to account of original
		end try
		if acct is missing value then
			set newMsg to make new outgoing message with properties {subject:subj, content:htmlBody}
		else
			set newMsg to make new outgoing message with properties {subject:subj, content:htmlBody, account:acct}
		end if
		set senderName to ""
		set senderAddr to ""
		try
			set snd to sender of original
			try
				set senderName to name of snd
			end try
			try
				set senderAddr to address of snd
			end try
		end try
		if senderAddr is not "" then
			make new to recipient at newMsg with properties {email address:{name:senderName, address:senderAddr}}
		end if
		if replyAll then
			repeat with r in (to recipients of original)
				my copyRecipient(contents of r, newMsg, "to", senderAddr, myAddresses)
			end repeat
			repeat with r in (cc recipients of original)
				my copyRecipient(contents of r, newMsg, "cc", senderAddr, myAddresses)
			end repeat
		end if
	end tell
	return newMsg
end manualReply

on copyRecipient(r, newMsg, kind, senderAddr, myAddresses)
	tell application "Microsoft Outlook"
		try
			set ea to email address of r
			set rName to ""
			set rAddr to ""
			try
				set rName to name of ea
			end try
			try
				set rAddr to address of ea
			end try
			if rAddr is "" or rAddr is senderAddr or myAddresses contains rAddr then return
			if kind is "cc" then
				make new cc recipient at newMsg with properties {email address:{name:rName, address:rAddr}}
			else
				make new to recipient at newMsg with properties {email address:{name:rName, address:rAddr}}
			end if
		end try
	end tell
end copyRecipient

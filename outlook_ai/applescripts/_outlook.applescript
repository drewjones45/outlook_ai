-- ===== outlook-ai Outlook helpers (appended to scripts that talk to Outlook) =====
-- Every Outlook term used in these scripts was checked against Outlook's
-- published scripting dictionary (Office 2011, 2016 and 2019 copies).

-- Fail fast instead of launching Outlook: a launch can sit behind a first-run
-- or sign-in dialog and every Apple Event would then time out.
on requireOutlook()
	if not (application "Microsoft Outlook" is running) then error "Microsoft Outlook isn't running. Open it and try again." number -600
end requireOutlook

-- Outlook "email address" record {name:..., address:..., type:...} -> "name FS address".
-- Record fields must be read inside the Outlook tell block: their keys are
-- Outlook dictionary terms.
on addrPair(rec)
	set n to ""
	set a to ""
	tell application "Microsoft Outlook"
		try
			set n to name of rec
		end try
		try
			set a to address of rec
		end try
	end tell
	return my clean(n) & my sepPart() & my clean(a)
end addrPair

-- A list of recipients or attendees -> "name FS address" items joined by GS.
on recipientList(recipientRefs)
	set outList to {}
	repeat with r in recipientRefs
		try
			set oneRef to contents of r
			tell application "Microsoft Outlook"
				set ea to email address of oneRef
			end tell
			set end of outList to my addrPair(ea)
		end try
	end repeat
	return my joinList(outList, my sepItem())
end recipientList

-- Find an account by display name or email address.
on findAccount(accountName)
	tell application "Microsoft Outlook"
		set candidates to {}
		try
			set candidates to candidates & (every exchange account)
		end try
		try
			set candidates to candidates & (every imap account)
		end try
		try
			set candidates to candidates & (every pop account)
		end try
		repeat with a in candidates
			try
				if (name of a) is accountName then return contents of a
			end try
			try
				if (email address of a) is accountName then return contents of a
			end try
		end repeat
	end tell
	error "No Outlook account named '" & accountName & "'. Run `outlook-ai doctor` to list accounts." number 1001
end findAccount

-- folderRole: "inbox" or "sent". With no account configured, use the default
-- account's folders; otherwise the named account's own inbox / sent items,
-- falling back to a folder looked up by name.
on resolveFolder(folderRole, accountName, folderName)
	tell application "Microsoft Outlook"
		if accountName is "" then
			if folderRole is "sent" then return sent items
			return inbox
		end if
	end tell
	set acct to my findAccount(accountName)
	tell application "Microsoft Outlook"
		try
			if folderRole is "sent" then
				set f to sent items of acct
			else
				set f to inbox of acct
			end if
			get name of f
			return f
		end try
		set namesToTry to {folderName}
		if folderRole is "sent" then set namesToTry to namesToTry & {"Sent Items", "Sent", "Sent Messages", "Sent Mail"}
		repeat with n in namesToTry
			try
				set f to mail folder (contents of n) of acct
				get name of f
				return f
			end try
		end repeat
	end tell
	error "No " & folderRole & " folder '" & folderName & "' in Outlook account '" & accountName & "'. Check [outlook] in your config." number 1001
end resolveFolder

-- One EVT record (see list_events.applescript for the layout). calName may be
-- "" to look the calendar up.
on eventRecord(ev, calName)
	set evId to ""
	set subj to ""
	set startText to ""
	set endText to ""
	set allDay to "0"
	set loc to ""
	set org to ""
	set recurring to "0"
	set busyText to ""
	set attList to {}
	set bodyText to ""
	set isOcc to "0"
	set recId to ""
	set masterId to ""
	tell application "Microsoft Outlook"
		with timeout of 120 seconds
			try
				set evId to (id of ev) as text
			end try
			try
				set subj to subject of ev
			end try
			if calName is "" then
				try
					set calName to name of (calendar of ev)
				end try
			end if
			try
				set startText to my isoDate(start time of ev)
			end try
			try
				set endText to my isoDate(end time of ev)
			end try
			try
				set allDay to my boolText(all day flag of ev)
			end try
			try
				set loc to location of ev
			end try
			try
				set org to organizer of ev
			end try
			try
				set recurring to my boolText(is recurring of ev)
			end try
			try
				set busyValue to free busy status of ev
				set busyText to my clean(busyValue)
			end try
			try
				repeat with a in (attendees of ev)
					if (count of attList) is greater than or equal to 60 then exit repeat
					set att to contents of a
					set ea to email address of att
					set aType to ""
					set aStatus to ""
					try
						set aType to type of att
					end try
					try
						set aStatus to status of att
					end try
					set end of attList to my addrPair(ea) & my sepPart() & my clean(aType) & my sepPart() & my clean(aStatus)
				end repeat
			end try
			try
				set bodyText to plain text content of ev
			end try
			try
				set isOcc to my boolText(is occurrence of ev)
			end try
			if isOcc is "1" then
				try
					set recId to my isoDate(recurrence id of ev)
				end try
				try
					set mst to master of ev
					set masterId to (id of mst) as text
				end try
			end if
		end timeout
	end tell
	return my makeRecord({"EVT", evId, my clean(subj), startText, endText, allDay, my clean(loc), my clean(org), my clean(calName), recurring, busyText, my joinList(attList, my sepItem()), my cleanLimit(bodyText, 4000), "", isOcc, recId, masterId})
end eventRecord

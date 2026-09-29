-- outlook-ai: describe what this Outlook exposes (version, accounts, folders, calendars).
-- Usage: osascript diagnose.applescript [accountName]
-- Records:
--   APP    | version
--   ACCT   | kind | name | email | isDefault(1/0) | fullName | microsoftOnline(1/0/"")
--   FOLDER | account | folderName | unreadCount | messageCount
--   CAL    | calendarName | account
--   INBOX  | account | folderName        (the folder outlook-ai will index)
--   NOTE   | text                       (problems worth showing the user)

on run argv
	my requireOutlook()
	set accountName to ""
	try
		if (count of argv) > 0 then set accountName to item 1 of argv
	end try
	set outList to {}
	tell application "Microsoft Outlook"
		set appVersion to ""
		try
			set appVersion to (get version)
		end try
		set end of outList to my makeRecord({"APP", my clean(appVersion)})

		set defaultName to ""
		try
			set defaultName to name of default account
		end try

		set acctList to {}
		try
			repeat with a in (every exchange account)
				set end of acctList to {"exchange", contents of a}
			end repeat
		end try
		try
			repeat with a in (every imap account)
				set end of acctList to {"imap", contents of a}
			end repeat
		end try
		try
			repeat with a in (every pop account)
				set end of acctList to {"pop", contents of a}
			end repeat
		end try
		if (count of acctList) is 0 then set end of outList to my makeRecord({"NOTE", "Outlook reported no mail accounts. If Outlook is in New Outlook mode, switch to Legacy Outlook (Outlook menu > Legacy Outlook)."})

		repeat with pair in acctList
			set acctKind to item 1 of pair
			set acct to item 2 of pair
			set acctName to ""
			set acctEmail to ""
			set acctFull to ""
			set msOnline to ""
			try
				set acctName to name of acct
			end try
			if acctKind is "exchange" then
				try
					set msOnline to my boolText(is microsoft online of acct)
				end try
			end if
			try
				set acctEmail to email address of acct
			end try
			try
				set acctFull to full name of acct
			end try
			set isDefault to "0"
			if acctName is not "" and acctName is defaultName then set isDefault to "1"
			set end of outList to my makeRecord({"ACCT", acctKind, my clean(acctName), my clean(acctEmail), isDefault, my clean(acctFull), msOnline})
			try
				repeat with f in (every mail folder of acct)
					set fName to ""
					set fUnread to ""
					set fCount to ""
					try
						set fName to name of f
					end try
					try
						set fUnread to (unread count of f) as text
					end try
					try
						set fCount to (count of messages of f) as text
					end try
					set end of outList to my makeRecord({"FOLDER", my clean(acctName), my clean(fName), fUnread, fCount})
				end repeat
			end try
		end repeat

		try
			repeat with c in (every calendar)
				set cName to ""
				set cAcct to ""
				try
					set cName to name of c
				end try
				try
					set cAcct to name of (account of c)
				end try
				set end of outList to my makeRecord({"CAL", my clean(cName), my clean(cAcct)})
			end repeat
		end try
	end tell

	try
		set theFolder to my resolveFolder("inbox", accountName, "Inbox")
		tell application "Microsoft Outlook"
			set inboxAcct to ""
			try
				set inboxAcct to name of (account of theFolder)
			end try
			set end of outList to my makeRecord({"INBOX", my clean(inboxAcct), my clean(name of theFolder)})
		end tell
	on error errMsg
		set end of outList to my makeRecord({"NOTE", "Couldn't open the inbox: " & errMsg})
	end try
	return my joinList(outList, my sepRecord())
end run

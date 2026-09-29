-- ===== outlook-ai shared helpers (appended to every script at run time) =====
-- Output format: records separated by ASCII 30, fields by ASCII 31,
-- list items by ASCII 29, item parts by ASCII 28. clean() strips those
-- characters from every value so data can never break the framing.

on sepRecord()
	return character id 30
end sepRecord

on sepField()
	return character id 31
end sepField

on sepItem()
	return character id 29
end sepItem

on sepPart()
	return character id 28
end sepPart

on replaceText(s, findText, replaceWith)
	if s does not contain findText then return s
	set oldTIDs to AppleScript's text item delimiters
	set AppleScript's text item delimiters to findText
	set parts to text items of s
	set AppleScript's text item delimiters to replaceWith
	set s to parts as text
	set AppleScript's text item delimiters to oldTIDs
	return s
end replaceText

on clean(v)
	if v is missing value then return ""
	try
		set s to v as text
	on error
		return ""
	end try
	repeat with n in {28, 29, 30, 31}
		set s to my replaceText(s, character id (n as integer), " ")
	end repeat
	return s
end clean

on cleanLimit(v, maxChars)
	set s to my clean(v)
	if maxChars > 0 and (length of s) > maxChars then set s to text 1 thru maxChars of s
	return s
end cleanLimit

on joinList(theList, sep)
	set oldTIDs to AppleScript's text item delimiters
	set AppleScript's text item delimiters to sep
	set s to theList as text
	set AppleScript's text item delimiters to oldTIDs
	return s
end joinList

on splitText(s, delim)
	set oldTIDs to AppleScript's text item delimiters
	set AppleScript's text item delimiters to delim
	set parts to text items of s
	set AppleScript's text item delimiters to oldTIDs
	return parts
end splitText

on makeRecord(fieldList)
	return my joinList(fieldList, my sepField())
end makeRecord

on boolText(b)
	try
		if b then return "1"
	end try
	return "0"
end boolText

on pad2(n)
	set s to (n as integer) as text
	if (length of s) < 2 then set s to "0" & s
	return s
end pad2

-- AppleScript date -> "YYYY-MM-DDTHH:MM:SS" (local time; locale independent)
on isoDate(d)
	if d is missing value then return ""
	try
		if class of d is not date then return ""
		set t to time of d
		return ((year of d) as integer as text) & "-" & my pad2(month of d as integer) & "-" & my pad2(day of d) & "T" & my pad2(t div 3600) & ":" & my pad2((t mod 3600) div 60) & ":" & my pad2(t mod 60)
	on error
		return ""
	end try
end isoDate

-- "YYYY-MM-DDTHH:MM:SS" (local time) -> AppleScript date; locale independent
on dateFromISO(s)
	set d to current date
	set day of d to 1
	set year of d to (text 1 thru 4 of s) as integer
	set month of d to (text 6 thru 7 of s) as integer
	set day of d to (text 9 thru 10 of s) as integer
	set time of d to ((text 12 thru 13 of s) as integer) * 3600 + ((text 15 thru 16 of s) as integer) * 60 + ((text 18 thru 19 of s) as integer)
	return d
end dateFromISO

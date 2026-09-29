-- outlook-ai: post a macOS notification (used by `outlook-ai notify`).
-- Usage: osascript notify.applescript <title> <subtitle> <message>

on run argv
	display notification (item 3 of argv) with title (item 1 of argv) subtitle (item 2 of argv)
	return ""
end run

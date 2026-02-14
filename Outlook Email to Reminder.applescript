-- ============================================================================
-- Outlook Email to Apple Reminder
-- ============================================================================
-- Creates an Apple Reminder from the currently selected Microsoft Outlook email.
-- The reminder title is set to the email subject, and the notes field contains
-- a deep link back to the original email in Outlook.
--
-- Designed for use with Keyboard Maestro but can also be run standalone
-- from Script Editor or via osascript.
-- ============================================================================

-- ============================================================================
-- Section 1: Verify Microsoft Outlook is running
-- ============================================================================
tell application "System Events"
	set outlookRunning to (exists (processes where name is "Microsoft Outlook"))
end tell

if not outlookRunning then
	display notification "Microsoft Outlook is not running." with title "Reminder Not Created" sound name "Basso"
	return
end if

-- ============================================================================
-- Section 2: Resolve selected Outlook message
-- ============================================================================
-- Keyboard Maestro can trigger quickly after app focus changes, so we wait
-- briefly, then try several Outlook dictionary entry points in order:
--   1) active explorer selection (most reliable for highlighted row)
--   2) selected objects of front window
--   3) selection
--   4) current messages
-- We keep the first object that looks like a mail message.
-- ============================================================================
set theMessage to missing value

try
	tell application "Microsoft Outlook"
		delay 0.35
		
		-- Method 1: highlighted messages in the message list
		if theMessage is missing value then
			try
				set candidateItems to selected objects of active explorer
				set theMessage to my firstMailMessageFromList(candidateItems)
			end try
		end if
		
		-- Method 2: selected objects from the front window
		if theMessage is missing value then
			try
				set candidateItems to selected objects of front window
				set theMessage to my firstMailMessageFromList(candidateItems)
			end try
		end if
		
		-- Method 3: generic selection property
		if theMessage is missing value then
			try
				set candidateItems to selection
				set theMessage to my firstMailMessageFromList(candidateItems)
			end try
		end if
		
		-- Method 4: fallback for some Outlook builds
		if theMessage is missing value then
			try
				set candidateItems to current messages
				set theMessage to my firstMailMessageFromList(candidateItems)
			end try
		end if
	end tell
on error errMsg
	display notification "Could not read Outlook selection: " & errMsg with title "Reminder Not Created" sound name "Basso"
	return
end try

if theMessage is missing value then
	display notification "No email is selected in Outlook." with title "Reminder Not Created" sound name "Basso"
	return
end if

-- ============================================================================
-- Section 3: Extract fields needed for the reminder
-- ============================================================================
set theSubject to "(No Subject)"
set theSender to "Unknown Sender"
set internetMessageId to ""
set numericMessageId to ""

try
	tell application "Microsoft Outlook"
		set theSubject to subject of theMessage
		if theSubject is missing value or theSubject is "" then set theSubject to "(No Subject)"
		
		try
			set theSender to name of sender of theMessage
		on error
			set theSender to "Unknown Sender"
		end try
		
		-- RFC Message-ID header (preferred for message:// links)
		try
			set internetMessageId to message id of theMessage
		on error
			set internetMessageId to ""
		end try
		
		-- Outlook internal item id (fallback link)
		try
			set numericMessageId to (id of theMessage) as text
		on error
			set numericMessageId to ""
		end try
	end tell
on error errMsg
	display notification "Could not read email details: " & errMsg with title "Reminder Not Created" sound name "Basso"
	return
end try

-- ============================================================================
-- Section 4: Build deep link URL back to the Outlook message
-- ============================================================================
-- Preferred format: message://%3C<internet-message-id>%3E
-- If Outlook does not expose a message-id header, fallback to
-- outlook://open?itemid=<internal-id>
-- ============================================================================
set outlookLink to ""

if internetMessageId is not "" then
	set cleanMessageId to my trimText(internetMessageId)	
	if cleanMessageId is not "" then
		if cleanMessageId does not start with "<" then set cleanMessageId to "<" & cleanMessageId
		if cleanMessageId does not end with ">" then set cleanMessageId to cleanMessageId & ">"
		set encodedMessageId to my encodeMessageIdForURL(cleanMessageId)
		set outlookLink to "message://" & encodedMessageId
	end if
end if

if outlookLink is "" and numericMessageId is not "" then
	set outlookLink to "outlook://open?itemid=" & numericMessageId
end if

if outlookLink is "" then
	display notification "Could not generate an Outlook link for this email." with title "Reminder Not Created" sound name "Basso"
	return
end if

-- ============================================================================
-- Section 5: Create reminder note content
-- ============================================================================
set reminderNote to "From: " & theSender & linefeed & linefeed & "Open in Outlook:" & linefeed & outlookLink

-- ============================================================================
-- Section 6: Create reminder in the default Reminders list
-- ============================================================================
try
	tell application "Reminders"
		set defaultList to default list
		make new reminder in defaultList with properties {name:theSubject, body:reminderNote}
	end tell
on error errMsg
	display notification "Could not create reminder: " & errMsg with title "Reminder Not Created" sound name "Basso"
	return
end try

-- ============================================================================
-- Section 7: Success notification
-- ============================================================================
display notification "\"" & theSubject & "\"" with title "Reminder Created" sound name "Glass"

-- ============================================================================
-- Helpers
-- ============================================================================

on firstMailMessageFromList(itemList)
	if itemList is missing value then return missing value
	if itemList is {} then return missing value
	
	repeat with oneItem in itemList
		try
			if class of oneItem is mail message then return oneItem
		on error
			-- Ignore non-mail objects.
		end try
	end repeat
	
	return missing value
end firstMailMessageFromList

on trimText(theText)
	if theText is missing value then return ""
	set t to (theText as text)
	set ws to {space, tab, return, linefeed}
	
	repeat while t is not "" and (character 1 of t) is in ws
		set t to text 2 thru -1 of t
	end repeat
	
	repeat while t is not "" and (character -1 of t) is in ws
		set t to text 1 thru -2 of t
	end repeat
	
	return t
end trimText

on encodeMessageIdForURL(rawMessageId)
	-- Minimal URL encoding sufficient for Message-ID values.
	set outText to rawMessageId
	set outText to my replaceText("%", "%25", outText)
	set outText to my replaceText("<", "%3C", outText)
	set outText to my replaceText(">", "%3E", outText)
	set outText to my replaceText(" ", "%20", outText)
	set outText to my replaceText("\"", "%22", outText)
	set outText to my replaceText("#", "%23", outText)
	return outText
end encodeMessageIdForURL

on replaceText(findText, replaceWith, sourceText)
	set AppleScript's text item delimiters to findText
	set textItems to every text item of sourceText
	set AppleScript's text item delimiters to replaceWith
	set newText to textItems as text
	set AppleScript's text item delimiters to ""
	return newText
end replaceText

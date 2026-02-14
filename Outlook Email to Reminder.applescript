-- ============================================================================
-- Outlook Email to Apple Reminder
-- ============================================================================
-- Creates an Apple Reminder from the currently selected Microsoft Outlook email.
-- The reminder title is set to the email subject, and the notes field contains
-- a clickable deep link (message://) back to the original email in Outlook.
--
-- Designed for use with Keyboard Maestro but can also be run standalone
-- from Script Editor or via osascript.
-- ============================================================================

-- ============================================================================
-- Section 1: Verify Microsoft Outlook is running
-- ============================================================================
-- Check that Outlook is open before attempting to read any mail data.
-- If Outlook isn't running we show a notification and exit early rather
-- than launching the app or throwing a cryptic scripting error.
-- ============================================================================
tell application "System Events"
	set outlookRunning to (exists (processes where name is "Microsoft Outlook"))
end tell

if not outlookRunning then
	display notification "Microsoft Outlook is not running." with title "Reminder Not Created" sound name "Basso"
	return
end if

-- ============================================================================
-- Section 2: Get the currently selected email in Outlook
-- ============================================================================
-- Outlook's AppleScript dictionary exposes "current messages" at the
-- application level, which returns a list of messages the user has
-- highlighted in the message list. We grab the first item from that list.
-- If nothing is selected (empty list) we notify the user and bail out.
-- ============================================================================
try
	tell application "Microsoft Outlook"
		set selectedMessages to current messages

		if selectedMessages is {} then
			display notification "No email is selected in Outlook." with title "Reminder Not Created" sound name "Basso"
			return
		end if

		set theMessage to item 1 of selectedMessages

		-- ====================================================================
		-- Section 3: Extract email properties
		-- ====================================================================
		-- Pull the subject line and the unique message ID from the message.
		-- The message ID is used to construct the deep link URL.
		-- We also grab the sender name for extra context in the reminder note.
		-- ====================================================================
		set theSubject to subject of theMessage
		set theMessageId to id of theMessage
		set theSender to (name of sender of theMessage)
	end tell

on error errMsg
	display notification "Could not read the selected email: " & errMsg with title "Reminder Not Created" sound name "Basso"
	return
end try

-- ============================================================================
-- Section 4: Build the Outlook deep link URL
-- ============================================================================
-- Microsoft Outlook on macOS supports the "message://" URL scheme, which
-- opens a specific email when clicked. The format is:
--     message://%3C<message-id>%3E
-- where %3C and %3E are URL-encoded angle brackets (< and >).
--
-- However, the id property from AppleScript is a numeric internal ID,
-- not the RFC Message-ID header. For the numeric ID, Outlook supports:
--     outlook://open?itemid=<numeric-id>
-- This scheme reliably opens the message from Reminders or any other app.
-- ============================================================================
set outlookLink to "outlook://open?itemid=" & theMessageId

-- ============================================================================
-- Section 5: Compose the reminder note body
-- ============================================================================
-- Include the sender and a clearly labelled link so the user can click
-- through from Reminders back to the email in Outlook.
-- ============================================================================
set reminderNote to "From: " & theSender & linefeed & linefeed & "Open in Outlook:" & linefeed & outlookLink

-- ============================================================================
-- Section 6: Create the reminder in Apple Reminders
-- ============================================================================
-- We add the reminder to the default list. Apple's Reminders app will
-- automatically recognise the URL in the note body and make it clickable.
-- ============================================================================
try
	tell application "Reminders"
		set defaultList to default list

		set newReminder to make new reminder in defaultList with properties {name:theSubject, body:reminderNote}
	end tell

on error errMsg
	display notification "Could not create reminder: " & errMsg with title "Reminder Not Created" sound name "Basso"
	return
end try

-- ============================================================================
-- Section 7: Success notification
-- ============================================================================
-- Give the user clear feedback that the reminder was created, including
-- the subject line so they can confirm it's the right email.
-- ============================================================================
display notification "\"" & theSubject & "\"" with title "Reminder Created" sound name "Glass"

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
-- When invoked via Keyboard Maestro the app switch can race the
-- AppleScript engine, so we give Outlook a moment to settle.
-- We then try two methods to obtain the selected message:
--   1. "selection" – returns whatever is highlighted in the UI
--   2. "current messages" – an older/alternative property
-- If neither yields a message we notify the user and bail out.
-- ============================================================================
try
	tell application "Microsoft Outlook"
		-- Small delay so Outlook registers the selection after an app switch
		delay 0.3

		-- Method 1: "selection" returns a list of selected objects
		set theMessage to missing value
		try
			set sel to selection
			if sel is not {} then
				set theMessage to item 1 of sel
			end if
		end try

		-- Method 2: fall back to "current messages"
		if theMessage is missing value then
			try
				set msgs to current messages
				if msgs is not {} then
					set theMessage to item 1 of msgs
				end if
			end try
		end if

		if theMessage is missing value then
			display notification "No email is selected in Outlook." with title "Reminder Not Created" sound name "Basso"
			return
		end if

		-- ====================================================================
		-- Section 3: Extract email properties
		-- ====================================================================
		-- Pull the subject line and the unique message ID from the message.
		-- The message ID is used to construct the deep link URL.
		-- We also grab the sender name for extra context in the reminder note.
		-- ====================================================================
		set theSubject to subject of theMessage
		set theMessageId to id of theMessage
		try
			set theSender to (name of sender of theMessage)
		on error
			set theSender to "Unknown"
		end try
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

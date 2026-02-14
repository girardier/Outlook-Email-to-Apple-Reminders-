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
-- Recent versions of Outlook for Mac (Classic v16.75+ and New Outlook)
-- have a known bug where the AppleScript "selection" and "current messages"
-- properties return empty results. We work around this by using JXA
-- (JavaScript for Automation) where selectedObjects() still functions.
--
-- The script tries three methods in order:
--   1. AppleScript "selection"
--   2. AppleScript "current messages"
--   3. JXA selectedObjects() via osascript (most reliable fallback)
-- ============================================================================

-- Bring Outlook to the front so the selection is registered
tell application "Microsoft Outlook" to activate
delay 0.5

set theSubject to ""
set theMessageId to ""
set theSender to ""
set gotMessage to false

-- ============================================================================
-- Method 1 & 2: Try native AppleScript properties
-- ============================================================================
try
	tell application "Microsoft Outlook"
		-- Try "selection" first
		try
			set sel to selection
			if sel is not {} and sel is not missing value then
				set theMessage to item 1 of sel
				set theSubject to subject of theMessage
				set theMessageId to id of theMessage
				try
					set theSender to (name of sender of theMessage)
				on error
					set theSender to "Unknown"
				end try
				set gotMessage to true
			end if
		end try

		-- Fall back to "current messages"
		if not gotMessage then
			try
				set msgs to current messages
				if msgs is not {} then
					set theMessage to item 1 of msgs
					set theSubject to subject of theMessage
					set theMessageId to id of theMessage
					try
						set theSender to (name of sender of theMessage)
					on error
						set theSender to "Unknown"
					end try
					set gotMessage to true
				end if
			end try
		end if
	end tell
end try

-- ============================================================================
-- Method 3: JXA fallback via osascript
-- ============================================================================
-- When the AppleScript dictionary is broken (common in Outlook 16.75+),
-- JXA's selectedObjects() often still works. We shell out to osascript
-- in JavaScript mode and parse the tab-delimited result.
-- ============================================================================
if not gotMessage then
	try
		set jxaScript to "
var app = Application('Microsoft Outlook');
var sel = app.selectedObjects();
if (sel.length === 0) { 'NO_SELECTION'; }
else {
  var m = sel[0];
  var subj = m.subject();
  var mid = m.id();
  var sender = 'Unknown';
  try { sender = m.sender.name(); } catch(e) {}
  subj + '\\t' + mid + '\\t' + sender;
}
"
		set jxaResult to do shell script "osascript -l JavaScript -e " & quoted form of jxaScript

		if jxaResult is not "NO_SELECTION" then
			set AppleScript's text item delimiters to tab
			set resultParts to text items of jxaResult
			set AppleScript's text item delimiters to ""

			set theSubject to item 1 of resultParts
			set theMessageId to item 2 of resultParts
			set theSender to item 3 of resultParts
			set gotMessage to true
		end if
	end try
end if

-- ============================================================================
-- If all methods failed, notify the user
-- ============================================================================
if not gotMessage then
	display notification "No email selected, or Outlook's AppleScript support is unavailable. If you are on the New Outlook, try reverting to Classic Outlook." with title "Reminder Not Created" sound name "Basso"
	return
end if

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

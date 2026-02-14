# Outlook Email to Apple Reminders

An AppleScript that creates an Apple Reminder from the currently selected Microsoft Outlook email, with a deep link back to the original message.

## What It Does

1. Resolves the currently selected email in Microsoft Outlook (with multiple fallback selection methods)
2. Creates a new reminder in Apple Reminders with:
   - **Title** set to the email subject line
   - **Notes** containing sender details and a clickable Outlook deep link
3. Uses a preferred `message://` link when Outlook provides an RFC Message-ID, with fallback to `outlook://open?itemid=` when needed

## Setup with Keyboard Maestro

1. Create a new macro in Keyboard Maestro
2. Set a trigger (e.g. a hotkey like `Ctrl+Shift+R`)
3. Add an **Execute AppleScript** action
4. Paste the contents of `Outlook Email to Reminder.applescript` into the action
5. Select an email in Outlook and press your hotkey

## Running Standalone

You can also run the script directly from Terminal:

```sh
osascript "Outlook Email to Reminder.applescript"
```

Or open it in Script Editor and click the Run button.

## Requirements

- macOS
- Microsoft Outlook for Mac
- Apple Reminders
- (Optional) Keyboard Maestro for hotkey triggering

## Error Handling

The script handles the following cases with user-facing notifications:

| Scenario | Notification |
|---|---|
| Outlook is not running | "Microsoft Outlook is not running." |
| No email selected | "No email is selected in Outlook." |
| Cannot read the selected email | Shows the underlying error message |
| Cannot generate a deep link | "Could not generate an Outlook link for this email." |
| Cannot create the reminder | Shows the underlying error message |
| Success | Shows the reminder title that was created |

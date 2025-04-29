# Outlook MCP Examples

This document provides examples of how to use the Outlook MCP tools with Claude.

## Calendar Examples

### Listing Calendar Events

```
Can you check my calendar for next week?
```

Claude will use the `listCalendarEvents` tool to fetch and display your calendar events.

### Creating a Calendar Event

```
Schedule a meeting with John tomorrow at 2pm for 1 hour. The subject should be "Project Review" and add john@example.com as an attendee.
```

Claude will use the `createCalendarEvent` tool to create a new meeting on your calendar.

### Updating a Calendar Event

```
Update my meeting with John tomorrow to start at 3pm instead of 2pm, and change the subject to "Project Status Review".
```

Claude will use the `listCalendarEvents` to find the meeting and then `updateCalendarEvent` to modify it.

## Email Examples

### Listing Recent Emails

```
Show me my recent emails from the inbox.
```

Claude will use the `listEmails` tool to fetch and display your recent emails.

```
Show me unread emails from the last 2 days.
```

Claude will use the `listEmails` tool with filtering to show specific emails.

### Reading an Email

```
Can you show me the full content of the email from John about the project update?
```

Claude will use `listEmails` to find the email and then `getEmail` to retrieve the full content.

### Sending an Email

```
Send an email to john@example.com with subject "Project Update" and the following message: "Hi John, Just following up on our discussion yesterday. Can you share the latest project metrics? Thanks!"
```

Claude will use the `sendEmail` tool to send the email.

### Managing Emails

```
Mark the unread email from Jane as read.
```

Claude will use `listEmails` to find the email and then `markEmailAsRead` to mark it as read.

```
Delete the spam email I received yesterday about lottery winnings.
```

Claude will use `listEmails` to find the email and then `deleteEmail` to remove it.

## Advanced Usage

### Combination of Calendar and Email

```
Check my calendar for tomorrow afternoon and then send an email to the meeting attendees reminding them about the meeting.
```

Claude will use both calendar and email tools together to accomplish the task.

### Creating a Draft Email

```
Create a draft email to the marketing team about the upcoming product launch.
```

Claude will use the `createDraft` tool to save a draft email without sending it.

Remember that Claude will handle the technical aspects of calling the MCP tools, parsing responses, and presenting the information in a user-friendly way. You just need to ask for what you want in natural language.
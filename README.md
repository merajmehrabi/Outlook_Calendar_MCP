# Outlook Calendar MCP Tool

A Model Context Protocol (MCP) server that allows Claude to access and manage your local Microsfot Outlook calendar (Windows only).

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Features

- **View Calendar Events**: List events within a date range, view event details, check attendee status
- **Manage Calendar Events**: Create new events and meetings, update existing events
- **Calendar Intelligence**: Find free time slots for scheduling, identify optimal meeting times
- **Multiple Calendar Support**: Access different calendars in your Outlook profile

## Prerequisites

- Windows operating system
- Microsoft Outlook desktop client installed
- Node.js (version 14.x or higher)
- npm (comes with Node.js)

## Installation

1. Clone this repository or download the source code
2. Install dependencies:

```bash
npm install
```

3. Build the project:

```bash
npm run build
```

## MCP Server Configuration

To use this tool with Claude, you need to add it to your MCP settings configuration file.

### For Claude Desktop App

Add the following to your Claude Desktop configuration file (located at `%APPDATA%\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "outlook-calendar": {
      "command": "node",
      "args": ["path/to/outlook-calendar-mcp/src/index.js"],
      "env": {}
    }
  }
}
```

### For Claude VSCode Extension

Add the following to your Claude VSCode extension MCP settings file (located at `%APPDATA%\Code\User\globalStorage\saoudrizwan.claude-dev\settings\cline_mcp_settings.json`):

```json
{
  "mcpServers": {
    "outlook-calendar": {
      "command": "node",
      "args": ["path/to/outlook-calendar-mcp/src/index.js"],
      "env": {}
    }
  }
}
```

Replace `path/to/outlook-calendar-mcp` with the actual path to where you installed this tool.

## Usage

Once configured, Claude will have access to the following tools:

### List Calendar Events

```
list_events
- startDate: Start date in MM/DD/YYYY format
- endDate: End date in MM/DD/YYYY format (optional)
- calendar: Calendar name (optional)
```

Example: "List my calendar events for next week"

### Create Calendar Event

```
create_event
- subject: Event subject/title
- startDate: Start date in MM/DD/YYYY format
- startTime: Start time in HH:MM AM/PM format
- endDate: End date in MM/DD/YYYY format (optional)
- endTime: End time in HH:MM AM/PM format (optional)
- location: Event location (optional)
- body: Event description (optional)
- isMeeting: Whether this is a meeting with attendees (optional)
- attendees: Semicolon-separated list of attendee email addresses (optional)
- calendar: Calendar name (optional)
```

Example: "Add a meeting with John about the project proposal on Friday at 2 PM"

### Find Free Time Slots

```
find_free_slots
- startDate: Start date in MM/DD/YYYY format
- endDate: End date in MM/DD/YYYY format (optional)
- duration: Duration in minutes (optional)
- workDayStart: Work day start hour (0-23) (optional)
- workDayEnd: Work day end hour (0-23) (optional)
- calendar: Calendar name (optional)
```

Example: "When am I free for a 1-hour meeting this week?"

### Get Attendee Status

```
get_attendee_status
- eventId: Event ID
- calendar: Calendar name (optional)
```

Example: "Who hasn't responded to my team meeting invitation?"

### Get Calendars

```
get_calendars
```

Example: "Show me my available calendars"

## Security Notes

- On first use, Outlook may display security prompts to allow script access
- The tool only accesses your local Outlook client and does not send calendar data to external servers
- All calendar operations are performed locally on your computer

## Troubleshooting

- **Outlook Security Prompts**: If you see security prompts from Outlook, you need to allow the script to access your Outlook data
- **Script Execution Policy**: If you encounter script execution errors, you may need to adjust your PowerShell execution policy
- **Path Issues**: Ensure the path in your MCP configuration file points to the correct location of the tool

## Contributing

We welcome contributions to the Outlook Calendar MCP Tool! Please see our [Contributing Guide](CONTRIBUTING.md) for details on how to get started.

By participating in this project, you agree to abide by our [Code of Conduct](CODE_OF_CONDUCT.md).

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

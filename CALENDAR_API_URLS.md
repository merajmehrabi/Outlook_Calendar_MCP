# Outlook Calendar API URLs (Microsoft Graph)

This document lists the Microsoft Graph REST API endpoints that correspond to each calendar operation implemented in this MCP server.

**Base URL:** `https://graph.microsoft.com/v1.0`

---

## 1. List Calendar Events

Retrieves events within a date range.

| | |
|---|---|
| **Method** | `GET` |
| **URL** | `https://graph.microsoft.com/v1.0/me/calendarView` |
| **Query params** | `startDateTime`, `endDateTime` |

```
GET https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=2025-01-01T00:00:00Z&endDateTime=2025-01-31T23:59:59Z
```

For a specific named calendar:
```
GET https://graph.microsoft.com/v1.0/me/calendars/{calendar-id}/calendarView?startDateTime=...&endDateTime=...
```

---

## 2. Create Calendar Event

Creates a new event or meeting in the calendar.

| | |
|---|---|
| **Method** | `POST` |
| **URL** | `https://graph.microsoft.com/v1.0/me/calendar/events` |

```
POST https://graph.microsoft.com/v1.0/me/calendar/events
```

For a specific named calendar:
```
POST https://graph.microsoft.com/v1.0/me/calendars/{calendar-id}/events
```

---

## 3. Get Event / Attendee Status

Retrieves a single event by ID, including attendee response status.

| | |
|---|---|
| **Method** | `GET` |
| **URL** | `https://graph.microsoft.com/v1.0/me/events/{event-id}` |

```
GET https://graph.microsoft.com/v1.0/me/events/{event-id}
```

Attendee info (accept/decline/tentative) is returned in the `attendees` array of the event response.

---

## 4. Update Calendar Event

Updates an existing event by ID.

| | |
|---|---|
| **Method** | `PATCH` |
| **URL** | `https://graph.microsoft.com/v1.0/me/events/{event-id}` |

```
PATCH https://graph.microsoft.com/v1.0/me/events/{event-id}
```

---

## 5. Delete Calendar Event

Deletes an event by ID.

| | |
|---|---|
| **Method** | `DELETE` |
| **URL** | `https://graph.microsoft.com/v1.0/me/events/{event-id}` |

```
DELETE https://graph.microsoft.com/v1.0/me/events/{event-id}
```

---

## 6. Get Calendars

Lists all calendars in the user's mailbox.

| | |
|---|---|
| **Method** | `GET` |
| **URL** | `https://graph.microsoft.com/v1.0/me/calendars` |

```
GET https://graph.microsoft.com/v1.0/me/calendars
```

---

## 7. Find Free / Busy Time Slots

Finds available meeting times across attendees (equivalent to `find_free_slots`).

| | |
|---|---|
| **Method** | `POST` |
| **URL** | `https://graph.microsoft.com/v1.0/me/findMeetingTimes` |

```
POST https://graph.microsoft.com/v1.0/me/findMeetingTimes
```

For raw free/busy information:
```
POST https://graph.microsoft.com/v1.0/me/calendar/getSchedule
```

---

## Quick Reference Table

| MCP Tool | HTTP Method | Graph API URL |
|---|---|---|
| `list_events` | GET | `https://graph.microsoft.com/v1.0/me/calendarView` |
| `create_event` | POST | `https://graph.microsoft.com/v1.0/me/calendar/events` |
| `get_attendee_status` | GET | `https://graph.microsoft.com/v1.0/me/events/{event-id}` |
| `update_event` | PATCH | `https://graph.microsoft.com/v1.0/me/events/{event-id}` |
| `delete_event` | DELETE | `https://graph.microsoft.com/v1.0/me/events/{event-id}` |
| `get_calendars` | GET | `https://graph.microsoft.com/v1.0/me/calendars` |
| `find_free_slots` | POST | `https://graph.microsoft.com/v1.0/me/findMeetingTimes` |

---

## Authentication

All Microsoft Graph requests require a Bearer token in the `Authorization` header:

```
Authorization: Bearer {access_token}
```

Obtain tokens via [Microsoft identity platform (OAuth 2.0)](https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow).

**Required scopes:**
- `Calendars.Read` — list events, get calendars, get attendee status
- `Calendars.ReadWrite` — create, update, delete events
- `Calendars.Read.Shared` — access shared calendars

---

## Reference

- [Microsoft Graph Calendar API overview](https://learn.microsoft.com/en-us/graph/api/resources/calendar?view=graph-rest-1.0)
- [Graph Explorer (test API calls)](https://developer.microsoft.com/en-us/graph/graph-explorer)

/**
 * outlookTools.js - Defines MCP tools for Outlook calendar operations
 */

import {
  listEvents,
  createEvent,
  findFreeSlots,
  getAttendeeStatus,
  getCalendars,
  deleteEvent,
  updateEvent
} from './scriptRunner.js';

/**
 * Defines the MCP tools for Outlook calendar operations
 * @returns {Object} - Object containing tool definitions
 */
export function defineOutlookTools() {
  return {
    // List calendar events
    list_events: {
      name: 'list_events',
      description: 'List calendar events within a specified date range',
      inputSchema: {
        type: 'object',
        properties: {
          startDate: {
            type: 'string',
            description: 'Start date in MM/DD/YYYY format'
          },
          endDate: {
            type: 'string',
            description: 'End date in MM/DD/YYYY format (optional)'
          },
          calendar: {
            type: 'string',
            description: 'Calendar name (optional)'
          }
        },
        required: ['startDate']
      },
      handler: async ({ startDate, endDate, calendar }) => {
        try {
          const events = await listEvents(startDate, endDate, calendar);
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(events, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error listing events: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Create calendar event
    create_event: {
      name: 'create_event',
      description: 'Create a new calendar event or meeting',
      inputSchema: {
        type: 'object',
        properties: {
          subject: {
            type: 'string',
            description: 'Event subject/title'
          },
          startDate: {
            type: 'string',
            description: 'Start date in MM/DD/YYYY format'
          },
          startTime: {
            type: 'string',
            description: 'Start time in HH:MM AM/PM format'
          },
          endDate: {
            type: 'string',
            description: 'End date in MM/DD/YYYY format (optional, defaults to start date)'
          },
          endTime: {
            type: 'string',
            description: 'End time in HH:MM AM/PM format (optional, defaults to 30 minutes after start time)'
          },
          location: {
            type: 'string',
            description: 'Event location (optional)'
          },
          body: {
            type: 'string',
            description: 'Event description/body (optional)'
          },
          isMeeting: {
            type: 'boolean',
            description: 'Whether this is a meeting with attendees (optional, defaults to false)'
          },
          attendees: {
            type: 'string',
            description: 'Semicolon-separated list of attendee email addresses (optional)'
          },
          calendar: {
            type: 'string',
            description: 'Calendar name (optional)'
          }
        },
        required: ['subject', 'startDate', 'startTime']
      },
      handler: async (eventDetails) => {
        try {
          const result = await createEvent(eventDetails);
          return {
            content: [
              {
                type: 'text',
                text: `Event created successfully with ID: ${result.eventId}`
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error creating event: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Find free time slots
    find_free_slots: {
      name: 'find_free_slots',
      description: 'Find available time slots in the calendar',
      inputSchema: {
        type: 'object',
        properties: {
          startDate: {
            type: 'string',
            description: 'Start date in MM/DD/YYYY format'
          },
          endDate: {
            type: 'string',
            description: 'End date in MM/DD/YYYY format (optional, defaults to 7 days from start date)'
          },
          duration: {
            type: 'number',
            description: 'Duration in minutes (optional, defaults to 30)'
          },
          workDayStart: {
            type: 'number',
            description: 'Work day start hour (0-23) (optional, defaults to 9)'
          },
          workDayEnd: {
            type: 'number',
            description: 'Work day end hour (0-23) (optional, defaults to 17)'
          },
          calendar: {
            type: 'string',
            description: 'Calendar name (optional)'
          }
        },
        required: ['startDate']
      },
      handler: async ({ startDate, endDate, duration, workDayStart, workDayEnd, calendar }) => {
        try {
          const freeSlots = await findFreeSlots(startDate, endDate, duration, workDayStart, workDayEnd, calendar);
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(freeSlots, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error finding free slots: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Get attendee status
    get_attendee_status: {
      name: 'get_attendee_status',
      description: 'Check the response status of meeting attendees',
      inputSchema: {
        type: 'object',
        properties: {
          eventId: {
            type: 'string',
            description: 'Event ID'
          },
          calendar: {
            type: 'string',
            description: 'Calendar name (optional)'
          }
        },
        required: ['eventId']
      },
      handler: async ({ eventId, calendar }) => {
        try {
          const attendeeStatus = await getAttendeeStatus(eventId, calendar);
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(attendeeStatus, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error getting attendee status: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Delete calendar event
    delete_event: {
      name: 'delete_event',
      description: 'Delete a calendar event by its ID',
      inputSchema: {
        type: 'object',
        properties: {
          eventId: {
            type: 'string',
            description: 'Event ID to delete'
          },
          calendar: {
            type: 'string',
            description: 'Calendar name (optional)'
          }
        },
        required: ['eventId']
      },
      handler: async ({ eventId, calendar }) => {
        try {
          const result = await deleteEvent(eventId, calendar);
          return {
            content: [
              {
                type: 'text',
                text: result.success 
                  ? `Event deleted successfully` 
                  : `Failed to delete event`
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error deleting event: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Update calendar event
    update_event: {
      name: 'update_event',
      description: 'Update an existing calendar event',
      inputSchema: {
        type: 'object',
        properties: {
          eventId: {
            type: 'string',
            description: 'Event ID to update'
          },
          subject: {
            type: 'string',
            description: 'New event subject/title (optional)'
          },
          startDate: {
            type: 'string',
            description: 'New start date in MM/DD/YYYY format (optional)'
          },
          startTime: {
            type: 'string',
            description: 'New start time in HH:MM AM/PM format (optional)'
          },
          endDate: {
            type: 'string',
            description: 'New end date in MM/DD/YYYY format (optional)'
          },
          endTime: {
            type: 'string',
            description: 'New end time in HH:MM AM/PM format (optional)'
          },
          location: {
            type: 'string',
            description: 'New event location (optional)'
          },
          body: {
            type: 'string',
            description: 'New event description/body (optional)'
          },
          calendar: {
            type: 'string',
            description: 'Calendar name (optional)'
          }
        },
        required: ['eventId']
      },
      handler: async ({ eventId, subject, startDate, startTime, endDate, endTime, location, body, calendar }) => {
        try {
          const result = await updateEvent(eventId, subject, startDate, startTime, endDate, endTime, location, body, calendar);
          return {
            content: [
              {
                type: 'text',
                text: result.success 
                  ? `Event updated successfully` 
                  : `Failed to update event`
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error updating event: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Get calendars
    get_calendars: {
      name: 'get_calendars',
      description: 'List available calendars',
      inputSchema: {
        type: 'object',
        properties: {}
      },
      handler: async () => {
        try {
          const calendars = await getCalendars();
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(calendars, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error getting calendars: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    }
  };
}

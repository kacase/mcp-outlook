#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { 
  CreateEventSchema, 
  SendEmailSchema,
  ListEventsQuerySchema,
  ListEmailsQuerySchema,
  SearchUsersQuerySchema,
  GetScheduleQuerySchema,
  FindMeetingTimesQuerySchema
} from "./types.js";
import { graphClient } from "./graphClient.js";

// Create server instance
const server = new McpServer({
  name: "outlook-mcp",
  version: "1.0.0"
});

// ============= Calendar Tools =============

server.tool(
  "listCalendarEvents",
  "Lists the user's calendar events for a specified time range",
  ListEventsQuerySchema.shape,
  async (params) => {
    try {
      // Fetch events from Graph API
      const events = await graphClient.listEvents({
        startDateTime: params.startDateTime,
        endDateTime: params.endDateTime,
        top: params.top,
        orderBy: params.orderBy
      });

      // Format events for display
      const formattedEvents = events.map(event => {
        const startDate = new Date(event.start.dateTime);
        const endDate = new Date(event.end.dateTime);
        
        return {
          id: event.id,
          subject: event.subject,
          start: startDate.toLocaleString(),
          end: endDate.toLocaleString(),
          timeZone: event.start.timeZone,
          location: event.location?.displayName || 'No location',
          isAllDay: event.isAllDay || false,
          attendees: event.attendees?.map(a => a.emailAddress.address).join(', ') || 'No attendees',
          preview: event.bodyPreview || ''
        };
      });

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(formattedEvents, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error listing calendar events: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "createCalendarEvent",
  "Creates a new calendar event",
  CreateEventSchema.shape,
  async (params) => {
    try {
      // Create event using Graph API
      const createdEvent = await graphClient.createEvent(params);
      
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(createdEvent, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error creating calendar event: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "getCalendarEvent",
  "Gets details of a specific calendar event",
  {
    eventId: z.string().describe("ID of the event to retrieve")
  },
  async (params) => {
    try {
      // Get event using Graph API
      const event = await graphClient.getEvent(params.eventId);
      
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(event, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error getting calendar event: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "updateCalendarEvent",
  "Updates an existing calendar event",
  {
    eventId: z.string().describe("ID of the event to update"),
    ...Object.fromEntries(
      Object.entries(CreateEventSchema.shape).map(([key, schema]) => [
        key, 
        (schema as z.ZodType<any>).optional()
      ])
    )
  },
  async (params) => {
    try {
      const { eventId, ...updateData } = params;
      
      // Update event using Graph API
      const updatedEvent = await graphClient.updateEvent(eventId, updateData);
      
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(updatedEvent, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error updating calendar event: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "deleteCalendarEvent",
  "Deletes a calendar event",
  {
    eventId: z.string().describe("ID of the event to delete")
  },
  async (params) => {
    try {
      // Delete event using Graph API
      await graphClient.deleteEvent(params.eventId);
      
      return {
        content: [
          {
            type: "text",
            text: `Event ${params.eventId} successfully deleted`
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error deleting calendar event: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

// ============= Email Tools =============

server.tool(
  "listEmails",
  "Lists the user's emails from a specified folder",
  ListEmailsQuerySchema.shape,
  async (params) => {
    try {
      // Fetch emails from Graph API
      const emails = await graphClient.listEmails({
        top: params.top,
        folder: params.folder || 'inbox',
        orderBy: params.orderBy,
        filter: params.filter,
        select: params.select
      });

      // Format emails for display
      const formattedEmails = emails.map(email => {
        let receivedDate = email.receivedDateTime ? new Date(email.receivedDateTime).toLocaleString() : 'Unknown';
        
        return {
          id: email.id,
          subject: email.subject || '(No Subject)',
          from: email.from?.emailAddress.address || 'Unknown',
          fromName: email.from?.emailAddress.name,
          received: receivedDate,
          isRead: email.isRead,
          importance: email.importance || 'normal',
          hasAttachments: email.hasAttachments || false,
          preview: email.bodyPreview || ''
        };
      });

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(formattedEmails, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error listing emails: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "getEmail",
  "Gets details of a specific email message",
  {
    messageId: z.string().describe("ID of the email message to retrieve")
  },
  async (params) => {
    try {
      // Get email using Graph API
      const email = await graphClient.getEmail(params.messageId);
      
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(email, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error getting email: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "sendEmail",
  "Sends a new email message",
  SendEmailSchema.shape,
  async (params) => {
    try {
      // Send email using Graph API
      await graphClient.sendEmail(params);
      
      return {
        content: [
          {
            type: "text",
            text: "Email successfully sent"
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error sending email: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "createDraft",
  "Creates a draft email message without sending it",
  SendEmailSchema.shape,
  async (params) => {
    try {
      // Create draft using Graph API
      const draft = await graphClient.createDraft(params);
      
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(draft, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error creating draft email: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "markEmailAsRead",
  "Marks an email message as read",
  {
    messageId: z.string().describe("ID of the email message to mark as read")
  },
  async (params) => {
    try {
      // Mark email as read using Graph API
      await graphClient.markAsRead(params.messageId);
      
      return {
        content: [
          {
            type: "text",
            text: `Email ${params.messageId} marked as read`
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error marking email as read: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "markEmailAsUnread",
  "Marks an email message as unread",
  {
    messageId: z.string().describe("ID of the email message to mark as unread")
  },
  async (params) => {
    try {
      // Mark email as unread using Graph API
      await graphClient.markAsUnread(params.messageId);
      
      return {
        content: [
          {
            type: "text",
            text: `Email ${params.messageId} marked as unread`
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error marking email as unread: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "deleteEmail",
  "Deletes an email message",
  {
    messageId: z.string().describe("ID of the email message to delete")
  },
  async (params) => {
    try {
      // Delete email using Graph API
      await graphClient.deleteEmail(params.messageId);
      
      return {
        content: [
          {
            type: "text",
            text: `Email ${params.messageId} successfully deleted`
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error deleting email: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

// ============= User Tools =============

server.tool(
  "searchUsers",
  "Searches for users by name within the organization",
  SearchUsersQuerySchema.shape,
  async (params) => {
    try {
      // Search users using Graph API
      const users = await graphClient.searchUsers({
        searchTerm: params.searchTerm,
        top: params.top
      });

      // Format users for display
      const formattedUsers = users.map(user => {
        return {
          id: user.id,
          displayName: user.displayName || 'Unknown',
          email: user.mail || user.userPrincipalName,
          jobTitle: user.jobTitle || '',
          department: user.department || ''
        };
      });

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(formattedUsers, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error searching users: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "getUser",
  "Gets details of a specific user by ID",
  {
    userId: z.string().describe("ID of the user to retrieve")
  },
  async (params) => {
    try {
      // Get user using Graph API
      const user = await graphClient.getUser(params.userId);
      
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(user, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error getting user: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

// ============= Schedule Tools =============

server.tool(
  "getSchedule",
  "Gets free/busy schedule information for specified users",
  GetScheduleQuerySchema.shape,
  async (params) => {
    try {
      // Get schedule using Graph API
      const scheduleInfo = await graphClient.getSchedule(params);
      
      // Format schedule for display
      const formattedSchedule = scheduleInfo.map(schedule => {
        const availabilityMap: Record<string, string> = {
          '0': 'free',
          '1': 'tentative',
          '2': 'busy',
          '3': 'out of office',
          '4': 'working elsewhere'
        };
        
        // Parse the availabilityView string which is a sequence of digits
        const parsedAvailability = schedule.availabilityView.split('').map(status => 
          availabilityMap[status] || 'unknown'
        );
        
        return {
          userId: schedule.scheduleId,
          availability: parsedAvailability,
          detailedItems: schedule.scheduleItems || []
        };
      });

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(formattedSchedule, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error getting schedule: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

server.tool(
  "findMeetingTimes",
  "Finds suitable meeting times for a group of attendees",
  FindMeetingTimesQuerySchema.shape,
  async (params) => {
    try {
      // Find meeting times using Graph API
      const suggestions = await graphClient.findMeetingTimes(params);
      
      // Format suggestions for display
      const formattedSuggestions = suggestions.map(suggestion => {
        const startTime = new Date(suggestion.meetingTimeSlot.start.dateTime);
        const endTime = new Date(suggestion.meetingTimeSlot.end.dateTime);
        
        return {
          startTime: startTime.toLocaleString(),
          endTime: endTime.toLocaleString(),
          timeZone: suggestion.meetingTimeSlot.start.timeZone,
          confidence: suggestion.confidence || 0,
          organizerAvailability: suggestion.organizerAvailability || 'unknown',
          attendeeAvailability: suggestion.attendeeAvailability?.map(a => ({
            attendee: a.attendee.emailAddress.address,
            availability: a.availability
          })) || []
        };
      });

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(formattedSuggestions, null, 2)
          }
        ]
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error finding meeting times: ${error instanceof Error ? error.message : String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

// ============= Resources =============

server.resource(
  "calendar",
  "https://graph.microsoft.com/v1.0/me/calendar/events",
  async (uri, extra) => {
    // Just call the graph client with empty parameters
    const events = await graphClient.listEvents({});
    
    // Convert the event data to a string
    const jsonData = JSON.stringify(events);
    
    // Return a properly formatted ReadResourceResult with the required contents property
    return {
      contents: [
        {
          uri: uri.toString(),
          text: jsonData,
          mimeType: "application/json"
        }
      ],
      _meta: {} // Optional metadata
    };
  }
);

server.resource(
  "inbox",
  "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages",
  async (uri, extra) => {
    // Call the graph client with inbox folder as parameter
    const emails = await graphClient.listEmails({ folder: "inbox" });
    
    // Convert the email data to a string
    const jsonData = JSON.stringify(emails);
    
    // Return a properly formatted ReadResourceResult with the required contents property
    return {
      contents: [
        {
          uri: uri.toString(),
          text: jsonData,
          mimeType: "application/json"
        }
      ],
      _meta: {} // Optional metadata
    };
  }
);

// Add an empty prompt to handle prompts/list
server.prompt(
  "outlook-empty-prompt",
  "Empty prompt",
  { param: z.string().optional().describe("Not used") },
  async () => {
    return {
      messages: []
    };
  }
);


// Connect the server to stdio transport
server.connect(new StdioServerTransport());
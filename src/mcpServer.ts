import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { 
  CreateEventSchema, 
  SendEmailSchema,
  ListEventsQuerySchema,
  ListEmailsQuerySchema,
  SearchPeopleQuerySchema,
  GetScheduleQuerySchema,
  FindMeetingTimesQuerySchema,
  AddAttendeesToEventSchema,
  AddAttendeesToEventParams
} from "./types.js";
import { graphClient } from "./graphClient.js";

export function setupMcpServer(): McpServer {
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
        const events = await graphClient.listEvents({
          startDateTime: params.startDateTime,
          endDateTime: params.endDateTime,
          top: params.top,
          orderBy: params.orderBy
        });

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
    "Creates a new calendar event. Ensure availability and user confirmation before sending invites.",
    CreateEventSchema.shape,
    async (params) => {
      try {
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
        const emails = await graphClient.listEmails({
          top: params.top,
          folder: params.folder || 'inbox',
          orderBy: params.orderBy,
          filter: params.filter,
          select: params.select
        });

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
    "sendEmail",
    "Sends a new email message",
    SendEmailSchema.shape,
    async (params) => {
      try {
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
    "getEmail",
    "Gets details of a specific email message",
    {
      messageId: z.string().describe("ID of the email message to retrieve")
    },
    async (params) => {
      try {
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
    "createDraft",
    "Creates a draft email message without sending it",
    SendEmailSchema.shape,
    async (params) => {
      try {
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

  // ============= People Tools =============

  server.tool(
    "searchPeople",
    "Searches for people relevant to the current user (colleagues, contacts, etc.)",
    SearchPeopleQuerySchema.shape,
    async (params) => {
      try {
        const people = await graphClient.searchPeople({
          searchTerm: params.searchTerm,
          filter: params.filter,
          select: params.select,
          top: params.top
        });

        const formattedPeople = people.map(person => {
          const primaryEmail = person.scoredEmailAddresses && person.scoredEmailAddresses.length > 0
            ? person.scoredEmailAddresses[0].address
            : '';

          const personClass = person.personType?.class || 'Unknown';
          const personSubclass = person.personType?.subclass || '';
          
          return {
            id: person.id,
            displayName: person.displayName || 'Unknown',
            email: primaryEmail,
            jobTitle: person.jobTitle || '',
            department: person.department || '',
            type: `${personClass}${personSubclass ? ` (${personSubclass})` : ''}`
          };
        });

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(formattedPeople, null, 2)
            }
          ]
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Error searching people: ${error instanceof Error ? error.message : String(error)}`
            }
          ],
          isError: true
        };
      }
    }
  );

  server.tool(
    "getPerson",
    "Gets details of a specific person by ID",
    {
      personId: z.string().describe("ID of the person to retrieve")
    },
    async (params) => {
      try {
        const person = await graphClient.getPerson(params.personId);
        
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(person, null, 2)
            }
          ]
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Error getting person: ${error instanceof Error ? error.message : String(error)}`
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
        const scheduleInfo = await graphClient.getSchedule(params);
        
        const formattedSchedule = scheduleInfo.map(schedule => {
          const availabilityMap: Record<string, string> = {
            '0': 'free',
            '1': 'tentative',
            '2': 'busy',
            '3': 'out of office',
            '4': 'working elsewhere'
          };
          
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
    "Finds suitable meeting times for a group of attendees. Correct E-Mail addresses are required.",
    FindMeetingTimesQuerySchema.shape,
    async (params) => {
      try {
        const suggestions = await graphClient.findMeetingTimes(params);
        
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

  server.tool(
    "addAttendeesToCalendarEvent",
    "Adds one or more attendees to an existing calendar event. Fetches the event, merges new attendees with existing ones (avoiding duplicates), and updates the event.",
    AddAttendeesToEventSchema.shape,
    async (params: AddAttendeesToEventParams) => {
      try {
        const updatedEvent = await graphClient.addAttendeesToEvent(params.eventId, params.attendees);
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
              text: `Error adding attendees to calendar event: ${error instanceof Error ? error.message : String(error)}`
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
      const events = await graphClient.listEvents({});
      const jsonData = JSON.stringify(events);
      
      return {
        contents: [
          {
            uri: uri.toString(),
            text: jsonData,
            mimeType: "application/json"
          }
        ],
        _meta: {}
      };
    }
  );

  server.resource(
    "inbox",
    "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages",
    async (uri, extra) => {
      const emails = await graphClient.listEmails({ folder: "inbox" });
      const jsonData = JSON.stringify(emails);
      
      return {
        contents: [
          {
            uri: uri.toString(),
            text: jsonData,
            mimeType: "application/json"
          }
        ],
        _meta: {}
      };
    }
  );

  server.prompt(
    "outlook-schedule-meeting-prompt",
    "A prompt to schedule a meeting with multiple attendees.",
    { param: z.string().optional().describe("Not used") },
    async () => {
      return {
        messages: []
      };
    }
  );

  return server;
}
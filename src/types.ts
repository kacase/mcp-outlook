import { z } from "zod";

// Schema for calendar events
export const CalendarEventSchema = z.object({
  id: z.string(),
  subject: z.string(),
  start: z.object({
    dateTime: z.string(),
    timeZone: z.string()
  }),
  end: z.object({
    dateTime: z.string(),
    timeZone: z.string()
  }),
  location: z.object({
    displayName: z.string().optional()
  }).optional(),
  attendees: z.array(
    z.object({
      emailAddress: z.object({
        name: z.string().optional(),
        address: z.string()
      })
    })
  ).optional(),
  bodyPreview: z.string().optional(),
  isAllDay: z.boolean().optional()
});

export type CalendarEvent = z.infer<typeof CalendarEventSchema>;

// Schema for creating events
export const CreateEventSchema = z.object({
  subject: z.string(),
  start: z.object({
    dateTime: z.string(),
    timeZone: z.string()
  }),
  end: z.object({
    dateTime: z.string(),
    timeZone: z.string()
  }),
  location: z.object({
    displayName: z.string()
  }).optional(),
  attendees: z.array(
    z.object({
      emailAddress: z.object({
        address: z.string(),
        name: z.string().optional()
      })
    })
  ).optional(),
  body: z.object({
    contentType: z.enum(["text", "html"]),
    content: z.string()
  }).optional(),
  isAllDay: z.boolean().optional()
});

export type CreateEventParams = z.infer<typeof CreateEventSchema>;

// Schema for Attendee Type )
export const AttendeeTypeSchema = z.enum(["required", "optional", "resource"]);

// Schema for a single Attendee 
export const AttendeeSchema = z.object({
  emailAddress: z.object({
    address: z.string().describe("The email address of the attendee."),
    name: z.string().optional().describe("The display name of the attendee.")
  }),
  type: AttendeeTypeSchema.optional().describe("The type of attendee. Default is 'required'.")
});

export type Attendee = z.infer<typeof AttendeeSchema>;

// Schema for adding attendees to an event
export const AddAttendeesToEventSchema = z.object({
  eventId: z.string().describe("ID of the calendar event to add attendees to."),
  attendees: z.array(AttendeeSchema).min(1).describe("List of attendees to add to the event.")
});

export type AddAttendeesToEventParams = z.infer<typeof AddAttendeesToEventSchema>;

// Schema for listing events query parameters
export const ListEventsQuerySchema = z.object({
  startDateTime: z.string().optional(),
  endDateTime: z.string().optional(),
  top: z.number().int().positive().optional(),
  filter: z.string().optional(),
  orderBy: z.string().optional()
});

export type ListEventsQuery = z.infer<typeof ListEventsQuerySchema>;

// Schema for email messages
export const EmailMessageSchema = z.object({
  id: z.string(),
  subject: z.string().optional(),
  bodyPreview: z.string().optional(),
  body: z.object({
    contentType: z.enum(["text", "html"]),
    content: z.string()
  }).optional(),
  from: z.object({
    emailAddress: z.object({
      name: z.string().optional(),
      address: z.string()
    })
  }).optional(),
  toRecipients: z.array(
    z.object({
      emailAddress: z.object({
        name: z.string().optional(),
        address: z.string()
      })
    })
  ).optional(),
  ccRecipients: z.array(
    z.object({
      emailAddress: z.object({
        name: z.string().optional(),
        address: z.string()
      })
    })
  ).optional(),
  bccRecipients: z.array(
    z.object({
      emailAddress: z.object({
        name: z.string().optional(),
        address: z.string()
      })
    })
  ).optional(),
  receivedDateTime: z.string().optional(),
  hasAttachments: z.boolean().optional(),
  importance: z.enum(["low", "normal", "high"]).optional(),
  isRead: z.boolean().optional()
});

export type EmailMessage = z.infer<typeof EmailMessageSchema>;

// Schema for creating/sending emails
export const SendEmailSchema = z.object({
  subject: z.string(),
  body: z.object({
    contentType: z.enum(["text", "html"]),
    content: z.string()
  }),
  toRecipients: z.array(
    z.object({
      emailAddress: z.object({
        address: z.string(),
        name: z.string().optional()
      })
    })
  ),
  ccRecipients: z.array(
    z.object({
      emailAddress: z.object({
        address: z.string(),
        name: z.string().optional()
      })
    })
  ).optional(),
  bccRecipients: z.array(
    z.object({
      emailAddress: z.object({
        address: z.string(),
        name: z.string().optional()
      })
    })
  ).optional(),
  importance: z.enum(["low", "normal", "high"]).optional(),
  saveToSentItems: z.boolean().optional().default(true)
});

export type SendEmailParams = z.infer<typeof SendEmailSchema>;

// Schema for listing emails query parameters
export const ListEmailsQuerySchema = z.object({
  top: z.number().int().positive().optional(),
  filter: z.string().optional(),
  orderBy: z.string().optional(),
  select: z.string().optional(),
  folder: z.string().optional().default("inbox")
});

export type ListEmailsQuery = z.infer<typeof ListEmailsQuerySchema>;

// Schema for People API response
export const PersonTypeSchema = z.object({
  class: z.enum(["Person", "Group"]),
  subclass: z.enum(["OrganizationUser", "UnifiedGroup"]).optional()
});

export const ScoredEmailAddressSchema = z.object({
  address: z.string(),
  relevanceScore: z.number().optional()
});

export const PhoneSchema = z.object({
  type: z.string().optional(), // business, mobile, home, etc.
  number: z.string()
});

export const PersonSchema = z.object({
  id: z.string(),
  displayName: z.string().optional(),
  givenName: z.string().nullable(),
  surname: z.string().nullable(),
  birthday: z.string().optional(),
  personNotes: z.string().optional(),
  isFavorite: z.boolean().optional(),
  jobTitle: z.string().nullable(),
  companyName: z.string().nullable(),
  department: z.string().nullable(),
  officeLocation: z.string().nullable(),
  profession: z.string().optional(),
  userPrincipalName: z.string().optional(),
  imAddress: z.string().nullable(),
  scoredEmailAddresses: z.array(ScoredEmailAddressSchema).optional(),
  phones: z.array(PhoneSchema).optional(),
  personType: PersonTypeSchema.optional()
});

export type Person = z.infer<typeof PersonSchema>;

// Schema for user search query parameters
export const SearchPeopleQuerySchema = z.object({
  searchTerm: z.string().optional().describe("The term to search for people by name or email (using $search)"),
  filter: z.string().optional().describe("Filter criteria for the people search (using $filter)"),
  top: z.number().int().positive().optional().describe("The maximum number of people to return"),
  select: z.string().optional().describe("Properties to include in the response")
});

export type SearchPeopleQuery = z.infer<typeof SearchPeopleQuerySchema>;

// Schema for schedule information
export const ScheduleInformationSchema = z.object({
  scheduleId: z.string(),
  availabilityView: z.string(),
  scheduleItems: z.array(
    z.object({
      status: z.enum(["free", "tentative", "busy", "oof", "workingElsewhere"]),
      start: z.object({
        dateTime: z.string(),
        timeZone: z.string()
      }),
      end: z.object({
        dateTime: z.string(),
        timeZone: z.string()
      })
    })
  ).optional()
});

export type ScheduleInformation = z.infer<typeof ScheduleInformationSchema>;

// Schema for getting free/busy schedule
export const GetScheduleQuerySchema = z.object({
  schedules: z.array(z.string()).describe("List of user IDs or email addresses"),
  startTime: z.object({
    dateTime: z.string(),
    timeZone: z.string()
  }).describe("Start time for the schedule query"),
  endTime: z.object({
    dateTime: z.string(),
    timeZone: z.string()
  }).describe("End time for the schedule query"),
  availabilityViewInterval: z.number().int().positive().optional().describe("Length of time slots in minutes")
});

export type GetScheduleQuery = z.infer<typeof GetScheduleQuerySchema>;

// Schema for meeting time suggestions
export const MeetingTimeSuggestionSchema = z.object({
  confidence: z.number().optional(),
  organizerAvailability: z.enum(["free", "tentative", "busy", "oof", "workingElsewhere"]).optional(),
  suggestionReason: z.string().optional(),
  meetingTimeSlot: z.object({
    start: z.object({
      dateTime: z.string(),
      timeZone: z.string()
    }),
    end: z.object({
      dateTime: z.string(),
      timeZone: z.string()
    })
  }),
  attendeeAvailability: z.array(
    z.object({
      attendee: z.object({
        emailAddress: z.object({
          address: z.string(),
          name: z.string().optional()
        })
      }),
      availability: z.enum(["free", "tentative", "busy", "oof", "workingElsewhere"])
    })
  ).optional()
});

export type MeetingTimeSuggestion = z.infer<typeof MeetingTimeSuggestionSchema>;

// Schema for finding meeting times
export const FindMeetingTimesQuerySchema = z.object({
  attendees: z.array(
    z.object({
      emailAddress: z.object({
        address: z.string(),
        name: z.string().optional()
      }),
      type: z.enum(["required", "optional"]).optional()
    })
  ).describe("List of attendees for the meeting"),
  timeConstraint: z.object({
    timeslots: z.array(
      z.object({
        start: z.object({
          dateTime: z.string(),
          timeZone: z.string()
        }),
        end: z.object({
          dateTime: z.string(),
          timeZone: z.string()
        })
      })
    )
  }).describe("Time constraints for the meeting"),
  meetingDuration: z.string().optional().describe("Duration of the meeting in ISO8601 format (e.g., 'PT1H' for 1 hour)"),
  maxCandidates: z.number().int().positive().optional().describe("Maximum number of meeting time suggestions"),
  minimumAttendeePercentage: z.number().optional().describe("Minimum percentage of attendees that need to be available")
});

export type FindMeetingTimesQuery = z.infer<typeof FindMeetingTimesQuerySchema>;
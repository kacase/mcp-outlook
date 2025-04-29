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
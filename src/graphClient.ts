import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import { getAccessToken, isAuthenticated, acquireToken } from './auth.js';
import { 
  CalendarEvent, 
  CreateEventParams, 
  ListEventsQuery,
  EmailMessage,
  SendEmailParams,
  ListEmailsQuery,
  Person,
  SearchPeopleQuery,
  GetScheduleQuery,
  ScheduleInformation,
  FindMeetingTimesQuery,
  MeetingTimeSuggestion,
  Attendee
} from './types.js';

// Define a simple function provider
const msalAuthProvider = async (done: (error: any, accessToken: string | null) => void) => {
  try {
    const token = await getAccessToken();
    done(null, token);
  } catch (error) {
    done(error, null);
  }
};

/**
 * Microsoft Graph client wrapper
 */
export class GraphClient {
  private client: Client;

  constructor() {
    this.client = Client.init({
      authProvider: msalAuthProvider,
    });
  }

  /**
   * Ensure the user is authenticated before calling Graph API
   */
  private async ensureAuthenticated(): Promise<void> {
    const authenticated = await isAuthenticated();
    if (!authenticated) {
      await acquireToken();
    }
  }

  // ============= Calendar Methods =============

  /**
   * List calendar events for the current user
   */
  async listEvents(query: ListEventsQuery): Promise<CalendarEvent[]> {
    await this.ensureAuthenticated();

    let endpoint = '/me/calendar/events';
    const queryParams = new URLSearchParams();

    // Add query parameters if provided
    if (query.startDateTime && query.endDateTime) {
      queryParams.append('$filter', `start/dateTime ge '${query.startDateTime}' and end/dateTime le '${query.endDateTime}'`);
    }

    if (query.top) {
      queryParams.append('$top', query.top.toString());
    }

    if (query.orderBy) {
      queryParams.append('$orderby', query.orderBy);
    }

    // Add query string to endpoint if we have parameters
    if (queryParams.toString()) {
      endpoint += `?${queryParams.toString()}`;
    }

    const response = await this.client.api(endpoint).get();
    return response.value;
  }

  /**
   * Create a new calendar event
   */
  async createEvent(eventData: CreateEventParams): Promise<CalendarEvent> {
    await this.ensureAuthenticated();

    const response = await this.client.api('/me/calendar/events').post(eventData);
    return response;
  }

  /**
   * Get a single calendar event by ID
   */
  async getEvent(eventId: string): Promise<CalendarEvent> {
    await this.ensureAuthenticated();

    const response = await this.client.api(`/me/calendar/events/${eventId}`).get();
    return response;
  }

  /**
   * Update an existing calendar event
   */
  async updateEvent(eventId: string, eventData: Partial<CreateEventParams>): Promise<CalendarEvent> {
    await this.ensureAuthenticated();

    const response = await this.client.api(`/me/calendar/events/${eventId}`).patch(eventData);
    return response;
  }

  /**
   * Delete a calendar event
   */
  async deleteEvent(eventId: string): Promise<void> {
    await this.ensureAuthenticated();

    await this.client.api(`/me/calendar/events/${eventId}`).delete();
  }

  /**
   * Add attendees to an existing calendar event.
   * Fetches the event, merges new attendees with existing ones (avoiding duplicates), then updates the event.
   * NOTE: This is a workaround for the Graph API's lack of support for adding attendees to an event, the existing 
   * attendees are passed as is from the event in addition to the new attendees.
   */
  async addAttendeesToEvent(eventId: string, newAttendees: Attendee[]): Promise<CalendarEvent> {
    await this.ensureAuthenticated();

    const event = await this.getEvent(eventId);
    // Existing attendees from CalendarEvent might not have 'type', so we handle them carefully.
    // The structure is { emailAddress: { address: string, name?: string } }
    const existingAttendees: Array<{ emailAddress?: { address?: string | null, name?: string | null } }> = event.attendees || [];

    // Create a Set of existing attendee email addresses for quick lookup
    const existingAttendeeEmails = new Set(
      existingAttendees
        .map((att: { emailAddress?: { address?: string | null } }) => att.emailAddress?.address?.toLowerCase())
        .filter((email: string | undefined): email is string => !!email) // Type guard to ensure email is string
    );

    const attendeesToAdd = newAttendees.filter(newAtt => {
      const newEmail = newAtt.emailAddress?.address?.toLowerCase();
      return newEmail && !existingAttendeeEmails.has(newEmail);
    });

    if (attendeesToAdd.length === 0) {
      return event;
    }

    // Combine existing attendees (exiting attendees from the event) with the genuinely new attendees.
    // The newAttendees conform to the Attendee schema.
    // The existingAttendees are passed as is from the event.
    // When PATCHing, Graph API expects the full desired list of attendees.
    const allAttendees = [
      ...existingAttendees.map(att => ({ // Ensure existing attendees match structure for PATCH if needed, though Graph might be flexible
        emailAddress: att.emailAddress,
        // 'type' would be undefined here, which is fine for existing attendees if Graph preserves it or defaults it.
      })),
      ...attendeesToAdd
    ];

    const response = await this.updateEvent(eventId, { attendees: allAttendees as any }); // Use 'as any' for now due to mixed types, Graph SDK should handle it.
    return response;
  }

  // ============= Email Methods =============

  /**
   * List emails from a specified folder (defaults to inbox)
   */
  async listEmails(query: ListEmailsQuery): Promise<EmailMessage[]> {
    await this.ensureAuthenticated();

    // Default to inbox if no folder specified
    const folder = query.folder || 'inbox';
    let endpoint = `/me/mailFolders/${folder}/messages`;
    const queryParams = new URLSearchParams();

    // Add query parameters if provided
    if (query.filter) {
      queryParams.append('$filter', query.filter);
    }

    if (query.top) {
      queryParams.append('$top', query.top.toString());
    }

    if (query.orderBy) {
      queryParams.append('$orderby', query.orderBy);
    }

    if (query.select) {
      queryParams.append('$select', query.select);
    }

    // Add query string to endpoint if we have parameters
    if (queryParams.toString()) {
      endpoint += `?${queryParams.toString()}`;
    }

    const response = await this.client.api(endpoint).get();
    return response.value;
  }

  /**
   * Get a single email message by ID
   */
  async getEmail(messageId: string): Promise<EmailMessage> {
    await this.ensureAuthenticated();

    const response = await this.client.api(`/me/messages/${messageId}`).get();
    return response;
  }

  /**
   * Send an email message
   */
  async sendEmail(emailData: SendEmailParams): Promise<void> {
    await this.ensureAuthenticated();

    const message = {
      ...emailData,
      message: {
        subject: emailData.subject,
        body: emailData.body,
        toRecipients: emailData.toRecipients,
        ccRecipients: emailData.ccRecipients,
        bccRecipients: emailData.bccRecipients,
        importance: emailData.importance
      },
      saveToSentItems: emailData.saveToSentItems
    };

    await this.client.api('/me/sendMail').post(message);
  }

  /**
   * Draft an email without sending (saves to Drafts folder)
   */
  async createDraft(emailData: SendEmailParams): Promise<EmailMessage> {
    await this.ensureAuthenticated();

    const draftData = {
      subject: emailData.subject,
      body: emailData.body,
      toRecipients: emailData.toRecipients,
      ccRecipients: emailData.ccRecipients,
      bccRecipients: emailData.bccRecipients,
      importance: emailData.importance
    };

    const response = await this.client.api('/me/messages').post(draftData);
    return response;
  }

  /**
   * Mark an email as read
   */
  async markAsRead(messageId: string): Promise<void> {
    await this.ensureAuthenticated();

    await this.client.api(`/me/messages/${messageId}`).patch({
      isRead: true
    });
  }

  /**
   * Mark an email as unread
   */
  async markAsUnread(messageId: string): Promise<void> {
    await this.ensureAuthenticated();

    await this.client.api(`/me/messages/${messageId}`).patch({
      isRead: false
    });
  }

  /**
   * Delete an email message
   */
  async deleteEmail(messageId: string): Promise<void> {
    await this.ensureAuthenticated();

    await this.client.api(`/me/messages/${messageId}`).delete();
  }

  // ============= People Methods =============

  /**
   * Search for people relevant to the current user
   */
  async searchPeople(query: SearchPeopleQuery): Promise<Person[]> {
    await this.ensureAuthenticated();

    let endpoint = '/me/people';
    const queryParams = new URLSearchParams();

    // Add search parameter if provided
    if (query.searchTerm) {
      queryParams.append('$search', `"${query.searchTerm}"`);
    }
    
    // Add filter if provided
    if (query.filter) {
      queryParams.append('$filter', query.filter);
    }

    // Add select parameter if provided
    if (query.select) {
      queryParams.append('$select', query.select);
    } else {
      // Default select to get relevant fields
      queryParams.append('$select', 'id,displayName,givenName,surname,userPrincipalName,scoredEmailAddresses,jobTitle,department,personType');
    }

    // Add top parameter if provided
    if (query.top) {
      queryParams.append('$top', query.top.toString());
    }

    // Add query string to endpoint
    endpoint += `?${queryParams.toString()}`;

    const response = await this.client.api(endpoint).get();
    return response.value;
  }

  /**
   * Get a single person by ID
   */
  async getPerson(personId: string): Promise<Person> {
    await this.ensureAuthenticated();

    const response = await this.client.api(`/me/people/${personId}`).get();
    return response;
  }

  // ============= Schedule Methods =============

  /**
   * Get free/busy schedule for users
   */
  async getSchedule(query: GetScheduleQuery): Promise<ScheduleInformation[]> {
    await this.ensureAuthenticated();

    const requestBody = {
      schedules: query.schedules,
      startTime: query.startTime,
      endTime: query.endTime,
      availabilityViewInterval: query.availabilityViewInterval || 30 // Default to 30-minute intervals
    };

    const response = await this.client.api('/me/calendar/getSchedule').post(requestBody);
    return response.value;
  }

  /**
   * Find meeting times for a group of users
   */
  async findMeetingTimes(query: FindMeetingTimesQuery): Promise<MeetingTimeSuggestion[]> {
    await this.ensureAuthenticated();

    // Use a type with index signature to allow dynamic property assignment
    const requestBody: {
      attendees: typeof query.attendees;
      timeConstraint: typeof query.timeConstraint;
      meetingDuration: string;
      maxCandidates: number;
      [key: string]: any;
    } = {
      attendees: query.attendees,
      timeConstraint: query.timeConstraint,
      meetingDuration: query.meetingDuration || 'PT1H', // Default to 1 hour
      maxCandidates: query.maxCandidates || 10
    };

    if (query.minimumAttendeePercentage) {
      requestBody.minimumAttendeePercentage = query.minimumAttendeePercentage;
    }

    const response = await this.client.api('/me/findMeetingTimes').post(requestBody);
    return response.meetingTimeSuggestions || [];
  }
}

// Export a singleton instance
export const graphClient = new GraphClient();
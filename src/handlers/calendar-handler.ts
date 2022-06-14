import { Client } from "@microsoft/microsoft-graph-client";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export default class CalendarHandler {

    constructor(private readonly client: Client) { }

    public async GetMyEvents(): Promise<[MicrosoftGraph.Event]> {
        try {
            const events = await this.client.api("/me/calendar/events")
                .select('subject,organizer,attendees,start,end,location,onlineMeeting,bodyPreview,webLink,body')
                .get();
            return events;
        } catch (error) {
            throw error;
        }
    }

    public async CreateOutlookCalendarEvent(userEvent: MicrosoftGraph.Event): Promise<[MicrosoftGraph.Event]> {
      //POST /users/{id | userPrincipalName}/calendar/events   <<< Da provare
  
      let res: [MicrosoftGraph.Event] = await this.client.api('/me/events')
        .post(userEvent);
  
      return res;
    }
  
    public async UpdateOutlookCalendarEventAttendees(eventId: string, newAtteendees: string): Promise<MicrosoftGraph.Event> {
      try {
        let res: MicrosoftGraph.Event = await this.client.api(`/me/events/${eventId}`)
          .patch(newAtteendees);
  
        return res;
      }
      catch (error) {
        throw error;
      }
    }
}
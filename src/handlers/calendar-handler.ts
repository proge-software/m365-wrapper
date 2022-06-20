import { Client } from "@microsoft/microsoft-graph-client";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import M365WrapperDataResult from "../models/results/m365-wrapper-data-result";
import ErrorsHandler from "./errors-handler";

export default class CalendarHandler {

  constructor(private readonly client: Client) { }

  public async getMyEvents(): Promise<M365WrapperDataResult<[MicrosoftGraph.Event]>> {
    try {
      let events: [MicrosoftGraph.Event] = await this.client.api("/me/calendar/events")
        .select('subject,organizer,attendees,start,end,location,onlineMeeting,bodyPreview,webLink,body')
        .get();

      return M365WrapperDataResult.createSuccess(events);
    } catch (error) {
      return ErrorsHandler.getErrorDataResult(error);
    }
  }

  public async createEvent(userEvent: MicrosoftGraph.Event): Promise<M365WrapperDataResult<[MicrosoftGraph.Event]>> {
    //POST /users/{id | userPrincipalName}/calendar/events   <<< Da provare

    try {
      let event: [MicrosoftGraph.Event] = await this.client.api('/me/events')
        .post(userEvent);

      return M365WrapperDataResult.createSuccess(event);
    }
    catch (error) {
      return ErrorsHandler.getErrorDataResult(error);
    }
  }

  public async updateEventAttendees(eventId: string, newAtteendees: string): Promise<M365WrapperDataResult<MicrosoftGraph.Event>> {
    try {
      let event: MicrosoftGraph.Event = await this.client.api(`/me/events/${eventId}`)
        .patch(newAtteendees);

      return M365WrapperDataResult.createSuccess(event);
    }
    catch (error) {
      return ErrorsHandler.getErrorDataResult(error);
    }
  }
}
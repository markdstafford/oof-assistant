import { User, Message } from '@microsoft/microsoft-graph-types'
import { getUsersWithExtensions, byNumberOfMeetings, sendReport } from "../graph-helpers";


export async function main (context?, req?) {
    if (context) context.log("Starting Azure function!");

    // GET /beta/users&$expand=extensions
    let users = await getUsersWithExtensions();

    // sort descending order of busiest calendars
    users.sort(byNumberOfMeetings)

    // get top 10 users with busiest calendars
    users = users.slice(0, 10);

    sendReport(users);
};

import { GraphClient } from '../auth-helpers';
import { User } from '@microsoft/microsoft-graph-types'
import { getMeetingCount, getUsersWithExtensions, removeAllExtensionsOnUser, saveMeetingCount } from "../graph-helpers";

// Graph Explorer check: https://graph.microsoft.com/beta/users?$select=displayName&$expand=extensions

export async function main(context?, req?) {
    context.log("Starting count-meetings");

    if (req.query.clear) { context.log("Clearing extensions"); }
    else { context.log("Saving extensions"); }

    // GET /users?$expand=extensions
    let users = await getUsersWithExtensions();

    for (let user of users) {
        if (req.query.clear) {
            // You can clear extensions by adding the "clear" query string parameter to the HTTP request
            removeAllExtensionsOnUser(user);
        }
        else {
            context.log(`Getting meeting count for ${user.userPrincipalName}`);
            // How many meetings are on their calendar next week?
            let meetingCount = await getMeetingCount(user);

            // Save to meeting count as a user extension
            await saveMeetingCount(user, meetingCount);
        }
    }

    let response = {
        status: 200,
        body: "Saved extensions on users!"
    };
    return response;
};

// Can use for local debugging even if you don't have Azure Functions Core Tools
//main();

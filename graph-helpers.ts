import { User, Message } from "@microsoft/microsoft-graph-types/microsoft-graph";
import { GraphClient } from "./auth-helpers";
import { FromEmail, ToEmail } from "./config";

export async function getUsersWithExtensions(): Promise<User[]> {
    const client = await GraphClient();

    return client
        .api("/users")
        .version("beta")
        .expand("extensions")
        .get()
        .then((res) => {
            return res.value;
        })
}

export async function removeAllExtensionsOnUser(user: User) {
    const client = await GraphClient();

    return client
        .api(`/users/${user.id}/extensions`)
        .version(`beta`)
        .get()
        .catch((e) => {
            debugger;
        }).then((res) => {
            let extensionIds = res['value'].map((extension) => extension.id);
            let extensionRemovals = [];
            for (let id of extensionIds) {
                extensionRemovals.push(
                    client
                        .api(`/users/${user.id}/extensions/${id}`)
                        .version(`beta`)
                        .delete()
                        .catch((e) => {
                            debugger;
                        }));
            }
            return Promise.all(extensionRemovals);
        })
}

export async function getMeetingCount(user: User) {
    const client = await GraphClient();

    let today = new Date();
    let inOneMonth = new Date(today.getTime() + 30 * 24 * 60 * 60 * 1000);

    return client
        .api(`/users/${user.mail}/calendarview/`)
        .query({
            startdatetime: today.toISOString(),
            enddatetime: inOneMonth.toISOString()
        })
        .get()
        .then((res) => {
            console.log(res.value.length);
            return res.value.length;
        })
        .catch((e) => {
            debugger;
            console.log(`Failed on user ${user.userPrincipalName}`);
            return 0;
        })
}

export async function saveMeetingCount(user: User, meetingCount: number) {
    const client = await GraphClient();

    let extensionData = {
        extensionName: "meetingCount",
        meetingCount: meetingCount
    };

    let id = "";
    if (user.extensions) {
        for (let extension of user.extensions) {
            if (extension.extensionName == "meetingCount") {
                id = extension.id;
            }
        }
    }

    if (id == "") {
        return client
            .api(`/users/${user.id}/extensions`)
            .version(`beta`)
            .post(extensionData)
            .catch((e) => {
                debugger;
            });
    } else {
        return client
            .api(`/users/${user.id}/extensions/${id}`)
            .version(`beta`)
            .patch(extensionData)
            .catch((e) => {
                debugger;
            })
    }
}

export function byNumberOfMeetings(a, b) {
    if (!a.extensions || !b.extensions) return 0;
    if (a.extensions[0].meetingCount > b.extensions[0].meetingCount)
        return -1;
    if (a.extensions[0].meetingCount < b.extensions[0].meetingCount)
        return 1;
    return 0;
}

export async function sendReport(users) {
    const client = await GraphClient();

    let emailString = ''

    for (let user of users) {
        if (user.extensions)
            emailString += "<tr><td>" + user.displayName + "</td><td>" + user.extensions[0].meetingCount + " events next week</td></tr>"
    }

    let message: Message = {
        subject: "Report on employee calendars",
        toRecipients: [{
            emailAddress: {
                address: ToEmail
            }
        }],
        body: {
            content: `<table>${emailString}</table>`,
            contentType: "html"
        }
    }
    return await client
        .api("/users/" + FromEmail + "/sendMail")
        .post({ message })
        .then((res) => {
            console.log("Mail sent!")
        }).catch((error) => {
            debugger;
        });
}

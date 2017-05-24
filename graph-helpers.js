"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const auth_helpers_1 = require("./auth-helpers");
const config_1 = require("./config");
function getUsersWithExtensions() {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield auth_helpers_1.GraphClient();
        return client
            .api("/users")
            .version("beta")
            .expand("extensions")
            .get()
            .then((res) => {
            return res.value;
        });
    });
}
exports.getUsersWithExtensions = getUsersWithExtensions;
function removeAllExtensionsOnUser(user) {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield auth_helpers_1.GraphClient();
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
                extensionRemovals.push(client
                    .api(`/users/${user.id}/extensions/${id}`)
                    .version(`beta`)
                    .delete()
                    .catch((e) => {
                    debugger;
                }));
            }
            return Promise.all(extensionRemovals);
        });
    });
}
exports.removeAllExtensionsOnUser = removeAllExtensionsOnUser;
function getMeetingCount(user) {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield auth_helpers_1.GraphClient();
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
        });
    });
}
exports.getMeetingCount = getMeetingCount;
function saveMeetingCount(user, meetingCount) {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield auth_helpers_1.GraphClient();
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
        }
        else {
            return client
                .api(`/users/${user.id}/extensions/${id}`)
                .version(`beta`)
                .patch(extensionData)
                .catch((e) => {
                debugger;
            });
        }
    });
}
exports.saveMeetingCount = saveMeetingCount;
function byNumberOfMeetings(a, b) {
    if (!a.extensions || !b.extensions)
        return 0;
    if (a.extensions[0].meetingCount > b.extensions[0].meetingCount)
        return -1;
    if (a.extensions[0].meetingCount < b.extensions[0].meetingCount)
        return 1;
    return 0;
}
exports.byNumberOfMeetings = byNumberOfMeetings;
function sendReport(users) {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield auth_helpers_1.GraphClient();
        let emailString = '';
        for (let user of users) {
            if (user.extensions)
                emailString += "<tr><td>" + user.displayName + "</td><td>" + user.extensions[0].meetingCount + " events next week</td></tr>";
        }
        let message = {
            subject: "Report on employee calendars",
            toRecipients: [{
                    emailAddress: {
                        address: config_1.ToEmail
                    }
                }],
            body: {
                content: `<table>${emailString}</table>`,
                contentType: "html"
            }
        };
        return yield client
            .api("/users/" + config_1.FromEmail + "/sendMail")
            .post({ message })
            .then((res) => {
            console.log("Mail sent!");
        }).catch((error) => {
            debugger;
        });
    });
}
exports.sendReport = sendReport;
//# sourceMappingURL=graph-helpers.js.map
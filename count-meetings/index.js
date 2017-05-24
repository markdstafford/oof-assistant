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
const graph_helpers_1 = require("../graph-helpers");
// Graph Explorer check: https://graph.microsoft.com/beta/users?$select=displayName&$expand=extensions
function main(context, req) {
    return __awaiter(this, void 0, void 0, function* () {
        context.log("Starting count-meetings");
        if (req.query.clear) {
            context.log("Clearing extensions");
        }
        else {
            context.log("Saving extensions");
        }
        // GET /users?$expand=extensions
        let users = yield graph_helpers_1.getUsersWithExtensions();
        for (let user of users) {
            if (req.query.clear) {
                // You can clear extensions by adding the "clear" query string parameter to the HTTP request
                graph_helpers_1.removeAllExtensionsOnUser(user);
            }
            else {
                context.log(`Getting meeting count for ${user.userPrincipalName}`);
                // How many meetings are on their calendar next week?
                let meetingCount = yield graph_helpers_1.getMeetingCount(user);
                // Save to meeting count as a user extension
                yield graph_helpers_1.saveMeetingCount(user, meetingCount);
            }
        }
        let response = {
            status: 200,
            body: "Saved extensions on users!"
        };
        return response;
    });
}
exports.main = main;
;
// Can use for local debugging even if you don't have Azure Functions Core Tools
//main();
//# sourceMappingURL=index.js.map
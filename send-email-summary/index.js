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
function main(context, req) {
    return __awaiter(this, void 0, void 0, function* () {
        if (context)
            context.log("Starting Azure function!");
        // GET /beta/users&$expand=extensions
        let users = yield graph_helpers_1.getUsersWithExtensions();
        // sort descending order of busiest calendars
        users.sort(graph_helpers_1.byNumberOfMeetings);
        // get top 10 users with busiest calendars
        users = users.slice(0, 10);
        graph_helpers_1.sendReport(users);
    });
}
exports.main = main;
;
//# sourceMappingURL=index.js.map
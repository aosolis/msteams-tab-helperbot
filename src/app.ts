// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

let express = require("express");
import { Request, Response } from "express";
let bodyParser = require("body-parser");
let http = require("http");
let path = require("path");
let logger = require("morgan");
let config = require("config");
import * as botbuilder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as winston from "winston";
import { HelperBot } from "./HelperBot";
import * as utils from "./utils";

let app = express();

app.set("port", process.env.PORT || 3978);
app.use(logger("dev"));
app.use(express.static(path.join(__dirname, "../../public")));
app.use(bodyParser.json());

// Create caption bot
let connector = new msteams.TeamsChatConnector({
    appId: config.get("bot.appId"),
    appPassword: config.get("bot.appPassword"),
});
let botSettings = {
    storage: new botbuilder.MemoryBotStorage(),
};
let bot = new HelperBot(connector, botSettings);
bot.on("error", (error: Error) => {
    winston.error(error.message, error);
});
app.post("/api/messages", connector.listen());

// Configure ping route
app.get("/ping", (req, res) => {
    res.status(200).send("OK");
});

// error handlers

// development error handler
// will print stacktrace
if (app.get("env") === "development") {
    app.use(function(err: any, req: Request, res: Response, next: Function): void {
        winston.error("Failed request", err);
        res.status(err.status || 500);
        res.render("error", {
            message: err.message,
            error: err,
        });
    });
}

// production error handler
// no stacktraces leaked to user
app.use(function(err: any, req: Request, res: Response, next: Function): void {
    winston.error("Failed request", err);
    res.status(err.status || 500);
    res.render("error", {
        message: err.message,
        error: {},
    });
});

http.createServer(app).listen(app.get("port"), function (): void {
    winston.verbose("Express server listening on port " + app.get("port"));
});

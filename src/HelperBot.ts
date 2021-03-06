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

import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as consts from "./constants";
import * as utils from "./utils";
import { ITenantStore, TenantData } from "./TenantStore";

// =========================================================
// HelperBot Bot
// =========================================================

export class HelperBot extends builder.UniversalBot {

    private tenantStore: ITenantStore;

    constructor(
        public _connector: builder.IConnector,
        private botSettings: any,
    )
    {
        super(_connector, botSettings);

        this.tenantStore = botSettings.tenantStore as ITenantStore;

        this.dialog(consts.DialogId.Root, this.handleMessage.bind(this));
        this.on("conversationUpdate", this.handleConversationUpdate.bind(this));
    }

    // Handle incoming messages
    private async handleMessage(session: builder.Session) {
        session.send("Hello!");
    }

    // Handle incoming conversation updates
    private async handleConversationUpdate(event: builder.IConversationUpdate) {

        // Save the mapping from tenant id to the service url for that tenant
        // To keep it simple, we do it for all incoming events. You may want to do it less frequently,
        // as the service url is stable for a tenant. 
        const tenantId = utils.getTenantId(event);
        if (tenantId) {
            const address = event.address as builder.IChatConnectorAddress;
            const serviceUrl = address.serviceUrl;
            if (serviceUrl) {
                await this.tenantStore.saveData(tenantId, { serviceUrl: serviceUrl });
            }
        }

    }
}

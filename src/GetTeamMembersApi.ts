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
import * as express from "express";
import { ITenantStore, TenantData } from "./TenantStore";

const knownServiceUrls = [
    "https://smba.trafficmanager.net/amer-client-ss.msg/",          // Americas region
    "https://smba.trafficmanager.net/emea-client-ss.msg/",          // EMEA region
    "https://smba.trafficmanager.net/apac-client-ss.msg/",          // Asia Pacific region
    "https://smba.trafficmanager.net/uklocal-client-ss.msg/",       // UK GoLocal
    "https://smba.trafficmanager.net/inlocal-client-ss.msg/",       // India GoLocal
];

// API handler to get team members
export class GetTeamMembersApi {
    
    constructor(
        private connector: msteams.TeamsChatConnector,
        private tenantStore: ITenantStore,
    ) { }

    public listen(): express.RequestHandler {
        return this.handleRequest.bind(this);
    }

    private async handleRequest(req: express.Request, res: express.Response): Promise<void> {
        const teamId = req.query.teamId;
        const tenantId = req.query.tenantId;
        let teamMembers: msteams.ChannelAccount[] = [];

        // We take the UPN of the user and return the members only of teams that the user is part of.
        // This prevents the service from being used to effectively get permission to read the membership of all groups.
        // IMPORTANT: Don't take this as a query parameter (we do that here only for convenience). In production
        // get this from the user's verified identity, e.g., validated AAD id_token, or your app's login. 
        const upn = req.query.upn;

        // Check that we have all required parameters
        if (!teamId || !tenantId || !upn) {
            res.sendStatus(400);
            return;
        }

        // Get the information we have for that tenant
        const tenantData = await this.tenantStore.getData(tenantId);

        try {
            if (tenantData && tenantData.serviceUrl) {
                // If we have a service url for the tenant use it directly
                teamMembers = await this.fetchTeamMembers(tenantData.serviceUrl, teamId);
            } else {
                // Otherwise, cycle through the known service urls
                for (let i = 0; i < knownServiceUrls.length; i++) {
                    try {
                        teamMembers = await this.fetchTeamMembers(knownServiceUrls[i], teamId);

                        // Success, we found the correct service url! Store it and end the iteration.
                        await this.tenantStore.saveData(tenantId, { serviceUrl: knownServiceUrls[i] });
                        break;
                    } catch (e) {
                        // If we get a 404, we are hitting the wrong region -- try the next one
                        if (e.statusCode && (e.statusCode === 404)) {
                            continue;
                        } else {
                            this.returnErrorResponse(res, e.statusCode, "Failed to get team members.");
                            return;
                        }
                    }
                }

                // We couldn't find the correct service URL
                if (!teamMembers) {
                    this.returnErrorResponse(res, 500, "Failed to find the correct service url.");
                    return;
                }
            }
        } catch (e) {
            if (e.statusCode) {
                this.returnErrorResponse(res, e.statusCode, "Failed to get team members.");
                return;
            } else {
                throw e;
            }
        }

        // Check that the user is a member of the team
        let lowerCaseUpn = upn.toLowerCase();
        if (!teamMembers.find(member => 
            (member.userPrincipalName.toLowerCase() === lowerCaseUpn) || 
            (member.email && member.email.toLowerCase() === lowerCaseUpn))) {
            this.returnErrorResponse(res, 403, "User must be a member of the team.");
            return;
        }

        // Return the UPN and email of team members
        const result = teamMembers.map(member => ({ upn: member.userPrincipalName, email: member.email }));
        res.status(200).json(result);
    }

    private fetchTeamMembers(serviceUrl: string, teamId: string): Promise<msteams.ChannelAccount[]> {
        return new Promise<msteams.ChannelAccount[]>((resolve, reject) => {
            this.connector.fetchMembers(serviceUrl, teamId, (err, members) => {
                if (err) {
                    reject(err);
                } else {
                    resolve(members);
                }
            });
        });
    }

    private returnErrorResponse(res: express.Response, statusCode: number, message: string) {
        const body = {
            statusCode: statusCode,
            message: "Failed to get team members.",
        };
        res.status(500).json(body);
    }

}
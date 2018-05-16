# msteams-tab-helperbot

This sample shows how to use a bot to provide more information to a tab.

## Getting started

The instructions below assume that you are using Visual Studio Code. If you're using a different environment,
in step 4 do the equivalent to set the `MICROSOFT_APP_ID` and `MICROSOFT_APP_PASSWORD` environment variables. 

1. Set up a tunneling service such as [ngrok](https://ngrok.com/).
2. Register a bot in [Microsoft Bot Framework](https://dev.botframework.com/). Remember the app id and password.
3. Set the messaging endpoint to `https://<ngrok_url>/api/messages`.
4. Setup `launch.json`.
    1. Create a `.vscode` folder and copy `launch.json` into that folder.
    2. In `launch.json`, set `MICROSOFT_APP_ID` and `MICROSOFT_APP_PASSWORD` to the bot's app id and password.
5. Run `npm install` then `gulp build`.
6. Launch the project.
7. Change the bot id in `manifest.json` to your bot's app id. Zip up the files in the `manifest` folder into `manifest.zip`.

## Try it out

1. Sideload the app into your team.
    * Your bot will receive a `conversationUpdate` message.
    * The message is handled [here](https://github.com/aosolis/msteams-tab-helperbot/blob/f4516657cd1eafc2a486e3204869bd35e41aadc7/src/HelperBot.ts#L57), where we take the tenant id and service url from the message and remember the mapping.
    * The sample uses an in-memory cache for convenience. You should use a persistent store in production.
2. The bot has an `/api/getTeamMembers?tenantId={tenantId}&teamId={teamId}&upn={upn}` that returns the members of the team.
    * These values are all available from your tab's context. Note that `teamId` is the `19:xxx` id, not the group id (which is a GUID).
    * This endpoint is for demonstration purposes only. We don't recommend making such an endpoint publicly-accessible (certainly not unauthenticated!) as it can be used as a vector to obtain information about the teams that are using your app. Instead, look at the code in the [handler](https://github.com/aosolis/msteams-tab-helperbot/blob/f4516657cd1eafc2a486e3204869bd35e41aadc7/src/GetTeamMembersApi.ts#L49), and integrate it into your backend.
3. Because the bot is marked as `isNotificationOnly = true`, even though it is part of the team, users cannot talk to it. If your app actually has a conversational component, set `isNotificationOnly = false` (or remove the property from the manifest).

Note that you won't actually see anything in Teams itself!

## Other considerations
1. Check that the user issuing the request is part of the team, before returning the member roster. Otherwise, your API can be used indirectly to get the membership of arbitrary teams who have your app installed.
2. Similarly, you should get the UPN of the user from a **verified** source; do not take it directly from the tab context.
    * If your app has its own login, you could take it from that, if the user's UPN is linked to their profile.
    * You can ask the user to login to Azure AD and get the UPN from the `id_token`. See [this blog post](https://techcommunity.microsoft.com/t5/Microsoft-Teams-Blog/Authentication-SSO-and-Microsoft-Graph-in-Microsoft-Teams-Tabs/ba-p/125366) for more information about authenticating the user in a tab.
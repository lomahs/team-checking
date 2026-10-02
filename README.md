# teams-permission-probe

Throwaway spike: finds out which Microsoft Teams operations your work account is allowed to do via Microsoft Graph,
so we can choose how to build the "remind members who didn't react ✅" tool.

No dependencies — just Node 22+.

## 1. Register an app (this is itself a test)

1. Go to https://portal.azure.com → **Microsoft Entra ID** → **App registrations** → **New registration**
   - Name: `teams-permission-probe`
   - Supported account types: **Accounts in this organizational directory only**
   - Redirect URI: leave empty → **Register**
   - ❌ If you can't create one ("you do not have access"), stop here — that's the answer: Graph/bot approaches need IT.
2. In the new app → **Authentication** → **Advanced settings** → **Allow public client flows: Yes** → **Save**
3. In **Overview**, copy **Application (client) ID** and **Directory (tenant) ID**.

You do **not** need to add API permissions in the portal — the script asks for them at sign-in.

## 2. Run

```bash
cp .env.example .env    # then paste the two IDs into .env
npm start
```

The script runs two sign-ins (open the link it prints, enter the code):

| Phase | Permissions | Normally needs admin? |
|---|---|---|
| Basic | User.Read, Team.ReadBasic.All, Channel.ReadBasic.All, Chat.ReadWrite, ChannelMessage.Send | No |
| Channel reading | ChannelMessage.Read.All, TeamMember.Read.All | Yes |

If the browser says **"Need admin approval"**, come back to the terminal and press **Enter** to skip.

Between sign-ins it asks you to pick a team, channel and group chat (Enter skips). Everything is read-only except
three optional tests that each ask `y/N` first:
- send a test message to the group chat you picked
- post a test message to the channel you picked
- open a 1:1 chat with a member (no message sent — they see nothing)

At the end you get a ✅/❌ report and a verdict. Paste it back to Claude.

## Cleanup

Delete the app registration in Entra ID when done.

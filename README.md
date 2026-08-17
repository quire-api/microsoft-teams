# Quire for Microsoft Teams
## Quire NodeJS handler Server for Microsoft Teams
* Handle webhooks from Microsoft Teams

## Issue tracking
Issues for this service are filed in **[zkoss/boeneo](https://github.com/zkoss/boeneo/issues)**, not here — that is where the team's tracker lives and where the related Copilot / MCP work is discussed. This repo has never had an issue opened on it, so anything filed here goes unseen.

Reference the boeneo issue number in commit messages.

## Two identities: manifest app id vs bot id
These are different GUIDs and always will be:

| | |
|---|---|
| Manifest app `id` | `dfad3be2-e62c-49c5-9453-ccd30999f003` |
| Bot registration | `fd335475-2341-4d79-a2ee-ed8439a33e9a` |

`dfad3be2` began as a developer's personal bot app id. The bot was later re-created under `admin@quire.io` as `fd335475`, but the app had already been submitted to the store with `dfad3be2` as its manifest `id`, and Microsoft's review team ruled the app id cannot change after submission.

`index.js` configures the adapter with a single app id:

```js
const adapter = new BotFrameworkAdapter({
  appId: process.env.BotId,          // fd335475-...
  appPassword: process.env.BotPassword
});
```

So any caller that authenticates as the *app* rather than the *bot* is rejected:

```
401 Unauthorized. Invalid AppId passed on token: dfad3be2-e62c-49c5-9453-ccd30999f003
```

This is what Copilot's auto-projected declarative agent (`ProjectedDeclarativeAgent.MessageExtension`) hits. Tracked in [boeneo#25569](https://github.com/zkoss/boeneo/issues/25569).

**Do not try to fix this by editing the manifest `id`** — Microsoft disallows it and it would break the store listing.

## The M365 app manifest lives elsewhere
This repo holds the bot runtime only. The manifest, icons, and release steps are in **[`quire-io/quire-mcp` → `microsoft-m365-package/`](https://github.com/quire-io/quire-mcp/tree/main/microsoft-m365-package)**, because the manifest now also declares the Quire MCP server as an `agentConnectors` entry and its companion tool-description file is generated from that repo's tool definitions.

## Setup
* Update package
```
npm update
```
* Prepare `env` file in project folder
  * It contains Open API id/secret
## Start server
```
node index.js
```

## Deploying
The service runs as the `msteams-msteams-1` container on the `quire-nodejs` EC2 host, from `quire/microsoft-teams` in ECR (`us-west-2`). Image tags are timestamps, e.g. `202608171220`.

Because tags are pinned, pushing a new image is not enough — the running container has to be pointed at the new tag and restarted, and **the logged version string is not a reliable indicator** (it only changes when someone bumps it in the code). To confirm what is actually deployed, grep inside the running container for a string unique to the commit you expect:

```
docker exec $(docker ps -q --filter name=msteams) grep -c "<string from your commit>" /app/bot/botActivityHandler.js
```

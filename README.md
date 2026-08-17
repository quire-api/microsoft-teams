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
Requires **Node 22** (the Dockerfile builds on `node:22-alpine`).

* Install dependencies from the lockfile
```
npm ci
```
* Prepare `env` file in project folder
  * It contains Open API id/secret
## Start server
```
node index.js
```

> **Do not run `npm update`.** The README previously said to, and on 2026-08-17 that is what broke production: it ignores `package-lock.json` and pulled forward a transitive dependency (`@typespec/ts-http-runtime`) that requires Node 19+, while the image was still on Node 16. The container then crash-looped for hours, and because each restart wiped the in-memory OAuth verification codes, Teams sign-in failed persistently.
>
> `package-lock.json` is committed and the Docker build uses `npm ci --omit=dev` precisely so this cannot recur. Use `npm ci` locally too, so what you run matches what ships. If a dependency genuinely needs upgrading, do it deliberately — bump it, commit the lockfile change, and check the Node engine requirements of whatever comes with it.

## Deploying
The service runs as the `msteams-msteams-1` container on the `quire-nodejs` EC2 host, from `quire/microsoft-teams` in ECR (`us-west-2`). Image tags are timestamps, e.g. `202608171220`.

Because tags are pinned, pushing a new image is not enough — the running container has to be pointed at the new tag and restarted, and **the logged version string is not a reliable indicator** (it only changes when someone bumps it in the code). To confirm what is actually deployed, grep inside the running container for a string unique to the commit you expect:

```
docker exec $(docker ps -q --filter name=msteams) grep -c "<string from your commit>" /app/bot/botActivityHandler.js
```

The tag itself comes from SSM parameter `/quire/production/msteams/image`, which **Jenkins updates automatically** on a successful build — every entry in its history is written by `user/Jenkins3-ECR`, so there is no manual step there. `scripts/updatemsteams.sh` on the host reads that parameter, pulls, and recreates the container.

The failure mode to watch for is subtler: a commit landing *after* the last build simply never gets picked up, and the deploy looks entirely successful because it faithfully ships whatever the parameter points at. Bump `build` in `index.js` when you want a deploy to be identifiable from the logs.

### Health endpoint
`GET /heartbeat` returns 200 (blocked externally by nginx; reachable as `localhost:3978/heartbeat` on the host). Intended for a monit watchdog — see [`quire-platform-docs` → `mis/production/nodejs-watchdogs.md`](https://github.com/quire-io/quire-platform-docs/blob/main/mis/production/nodejs-watchdogs.md). Note that a liveness watchdog would not have caught the 2026-08-17 incident, since the container was restarting itself rather than hanging; flapping detection is the part that matters for this service.

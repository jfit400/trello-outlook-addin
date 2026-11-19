# Trello for Outlook

An Outlook add-in that works like the Trello app for Slack: open an email, click **Add to Trello**, and either

- **create a new card** (board/list picker, title from the subject, description with sender, date, a link back to the email and the body, plus labels, members, due date and position), or
- **add the email to an existing card** (search your cards; the email is posted as a comment).

Either way it can attach the email itself as an `.eml` file and upload the email's attachments to the card. Your last board/list is remembered, and the pane can be pinned so it follows you from email to email.

Works in new Outlook for Windows, classic Outlook (Windows/Mac) and Outlook on the web. There's no build step: it's plain HTML/JS in `src/`.

## 1. Get a Trello API key

1. Go to <https://trello.com/power-ups/admin> and create a new Power-Up/integration (any workspace is fine).
2. Open its **API key** tab and generate a key.
3. Under **Allowed origins**, add `https://localhost:3000` (and later your production URL).
4. Paste the key into `src/config.js`.

## 2. Run it locally

```bash
npm install
```

```bash
npm run certs
```

This installs a trusted localhost HTTPS certificate. Windows will ask you to approve it. Then:

```bash
npm start
```

That serves the add-in at `https://localhost:3000`. To check the manifest:

```bash
npm run validate
```

## 3. Sideload into Outlook

**Outlook on the web / new Outlook:** go to <https://aka.ms/olksideload>, then **My add-ins → Custom Addins → Add a custom add-in → Add from File…**, and pick `manifest.xml`.

**Classic Outlook on Windows:** run `npm run sideload` (it starts the server and registers the manifest), or use the same web steps above. Add-ins sync to your mailbox.

Open any email. **Add to Trello** shows up in the ribbon, or under **Apps** / the `…` menu in the new Outlook. The first time, click **Connect Trello** and approve access.

## 4. Hosted version (GitHub Pages)

The `src/` folder is published to <https://jfit400.github.io/trello-outlook-addin/> by `.github/workflows/pages.yml` on every push to `main`. Nothing needs to run locally.

- Install the hosted add-in with **`manifest.prod.xml`**. It has its own add-in ID, so it can sit alongside the localhost dev copy.
- Trello key **Allowed origins** must include `https://jfit400.github.io`.
- If the site address changes, regenerate the manifest: `node scripts/make-manifest.js https://<new-site>/<path>`
- To roll it out to the team: Microsoft 365 admin center → **Settings → Integrated apps → Upload custom apps**, then upload `manifest.prod.xml`.

### Security notes

- The repo is public. The Trello **API key** in `src/config.js` is designed to be public. Never commit the Trello **Secret**.
- Users' Trello tokens live only in their Outlook roaming settings and go only to `api.trello.com`.
- Anyone who can push to this repo can change the code that handles those tokens. Keep 2FA on the GitHub account.
- Email content is only ever rendered as text (no `innerHTML`), so emails can't inject script into the pane.
- Revoke a token at Trello → Settings → Applications. Rotate the API key from the Trello app page if it's abused.

## How it works

| File | Purpose |
| --- | --- |
| `manifest.xml` | Add-in manifest: ribbon button on read messages, pinnable task pane |
| `src/taskpane.*` | The UI |
| `src/trello.js` | Small Trello REST client, called straight from the browser |
| `src/mail.js` | Reads the open email (subject, body, sender, attachments, `.eml`) via Office.js |
| `src/auth-start.html` / `auth-callback.html` | Trello sign-in in an Office dialog. The token comes back in the URL fragment |
| `server.js` | HTTPS static server for development |

- **Auth:** each user's Trello token is stored in Outlook *roaming settings*, so it follows their mailbox and never touches a server of yours. Sign out removes it. To revoke it fully, go to Trello → Settings → Applications.
- **API requirements:** core features need Mailbox 1.5. Attaching files needs 1.8, and attaching the `.eml` needs 1.14. If a client doesn't support these, the options are hidden.
- **Email link:** the "Open email in Outlook" link is a best-effort Outlook on the web deep link. The attached `.eml` is the reliable copy.

## Ideas for next steps

- Quick-create without opening the pane (an `ExecuteFunction` button that files the email into a default list)
- A compose-mode button that adds a Trello card link to an email you're writing
- Showing which cards already reference this email

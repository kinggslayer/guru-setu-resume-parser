# Deploying the resume parser

Two routes. Start with the free one.

---

# Option A — Streamlit Community Cloud (free)

Free, no server, no Docker. The app already reads `st.secrets`, so no code
change is needed.

**The free tier gives you:** ~1 GB memory, the app sleeps after 12 quiet
hours (next visitor sees a "waking up" page for ~30s), one private app, and
a `*.streamlit.app` URL — no custom domain.

**Make the app private.** It holds candidate names, phones, emails and CV
links, and it spends your OpenAI credits. A public app means anyone with the
URL gets all of that. You get exactly one private app on the free tier,
which is all you need.

### 1. Put the code on GitHub

A **private** repo is fine — Community Cloud supports them. Before pushing,
confirm these are NOT in the repo:

```bash
git status --porcelain --ignored | grep -E "\.env|service_account|secrets\.toml$"
```

`.gitignore` already covers `.env`, the `*.json` key, `secrets/` and
`.streamlit/secrets.toml`. Only `secrets.toml.example` should be committed.

### 2. Deploy

1. Sign in at <https://share.streamlit.io> with GitHub
2. **Create app** → pick the repo, branch `main`, main file `app.py`
3. **Advanced settings** → Python 3.12
4. Deploy

### 3. Add the secrets

App → **Settings → Secrets**. Use `.streamlit/secrets.toml.example` as the
shape: your `OPENAI_API_KEY`, and a `[gcp_service_account]` section built
from the Google JSON key file.

The `private_key` field is the usual stumbling block — keep the triple
quotes and leave the `\n` sequences exactly as they appear in the JSON.
Don't reformat it across real lines.

### 4. Make it private

App → **Settings → Sharing** → set to private, then invite the Google
accounts of whoever should have access.

### 5. Share the Drive folder

The service account has its own email (`client_email` in the JSON). Share
your resume Drive folder and the `Resume_Master_DB` sheet with that address,
or the app authenticates fine and then sees nothing.

### Living with the free tier

- **1 GB memory.** Process a few dozen resumes at a time, not several
  hundred. If the app dies mid-run, that's the memory ceiling — the records
  already written to the sheet are safe, so just re-run and the content-hash
  check skips them.
- **Sleeping.** Harmless, just slow on the first visit of the day.
- **`OPENAI_WORKERS = 2`.** Raising it doesn't speed things up much here and
  makes the memory limit easier to hit.

### When to leave

Move to Option B if you need a custom domain, the app to always be awake, or
bigger batches in one go.

---

# Option B — Hostinger VPS (paid)

## Which Hostinger plan

You need **Hostinger VPS**, not shared/web hosting.

Streamlit is a long-lived process holding a WebSocket per browser session,
with session state kept in memory. Shared hosting runs PHP and short-lived
CGI, so there is nowhere for that process to live. A VPS gives you root and
SSH, which is all this needs.

Pick the Ubuntu 24.04 template. The smallest plan (1 vCPU / 4 GB) is enough —
the heavy lifting happens at OpenAI, not on your server.

---

## One-time server setup

SSH in as root, then:

```bash
# Docker + compose plugin
curl -fsSL https://get.docker.com | sh

# 2 GB swap, so a large PDF can't trigger an out-of-memory kill
fallocate -l 2G /swapfile
chmod 600 /swapfile
mkswap /swapfile
swapon /swapfile
echo '/swapfile none swap sw 0 0' >> /etc/fstab

# Only these three ports open
ufw allow 22 && ufw allow 80 && ufw allow 443 && ufw --force enable
```

---

## DNS

In Hostinger's DNS editor for `gurusetuconsultancy.com`, add an **A record**:

| Type | Name     | Points to           | TTL  |
|------|----------|---------------------|------|
| A    | `parser` | *your VPS IPv4*     | 3600 |

That gives you `parser.gurusetuconsultancy.com`. Wait for it to resolve
(`dig +short parser.gurusetuconsultancy.com`) **before** starting Caddy —
Let's Encrypt validates over HTTP and will fail on a domain that doesn't
point anywhere yet, and repeated failures hit rate limits.

---

## Deploy

```bash
mkdir -p /opt/parser && cd /opt/parser
# upload the project here (scp -r ./AI-resume-parser/* root@VPS_IP:/opt/parser/)

mkdir -p secrets
# upload the Google service-account JSON as secrets/service_account.json
chmod 600 secrets/service_account.json
```

Create `/opt/parser/.env`:

```
OPENAI_API_KEY=sk-...
OPENAI_WORKERS=2
```

```bash
chmod 600 .env
```

Set the login password — the app has none of its own:

```bash
docker run --rm -it caddy:2-alpine caddy hash-password
```

Paste the hash into `Caddyfile`, replacing `REPLACE_WITH_BCRYPT_HASH`, and
change the username from `sandeep` if you want something else.

Then:

```bash
docker compose up -d --build
docker compose logs -f
```

Open `https://parser.gurusetuconsultancy.com`. Caddy fetches the TLS
certificate on first request; it takes a few seconds.

---

## Why the login matters

The app has no authentication of its own. Exposed openly, anyone who finds
the URL can read candidate names, phone numbers, emails and CV links, and
spend your OpenAI credits a folder at a time.

Two things in this setup prevent that:

1. `basic_auth` in the Caddyfile.
2. The parser container publishes **no** host ports — only `expose`. So
   `http://VPS_IP:8501` is not reachable from outside, and the auth check
   can't be walked around by hitting the IP directly.

If you ever add `ports: - "8501:8501"` to the parser service for debugging,
you have opened that hole. Use `docker compose exec` instead.

---

## Updating after a code change

```bash
cd /opt/parser
# upload changed files
docker compose up -d --build
```

Only Python changed? `docker compose restart parser` is faster.
`requirements.txt` changed? You need the `--build`.

---

## Secrets never belong in the image

`.dockerignore` excludes `.env`, `secrets/` and `*.json`, so a rebuilt
image contains no credentials. Both are mounted or injected at runtime.
If you push this to GitHub, check `.gitignore` covers them too — it does,
but confirm before the first push.

---

## Troubleshooting

**Certificate won't issue.** DNS isn't pointing at the VPS yet, or port 80
is blocked. Check `dig +short parser.gurusetuconsultancy.com` and `ufw status`.

**Page loads, then "Please wait..." forever.** The WebSocket isn't getting
through. If you put Cloudflare in front, that's the usual cause — either
turn the proxy off (grey cloud) or enable WebSockets in Cloudflare.

**Long runs disconnect.** Raise `read_timeout` in the Caddyfile above 30m.

**Container keeps restarting.** `docker compose logs parser`. Most often the
service-account JSON isn't at `secrets/service_account.json`, so the mount
creates a directory where a file is expected.

---

## Worth considering first

You already run Docker and Caddy on the AWS EC2 box in Mumbai serving
`gurusetuconsultancy.com`. Adding this there as another service behind the
existing Caddy costs nothing extra, keeps one server to patch, and puts the
parser on the same host as the portal it feeds.

The Hostinger VPS makes sense if you'd rather keep the two apps isolated,
or want the parser to stay up independently of the portal.

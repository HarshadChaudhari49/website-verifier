# Website Verifier — System 2

Unattended automation for a Copy & Paste website-qualification workflow. It
logs into the work portal, reads the assigned URL, decides whether the site
qualifies against a rulebook, submits the result, fetches the next URL, and
repeats until stopped.

Verification is done by ChatGPT, driven in a second browser tab: the rulebook
goes in once, then each assigned URL is sent on its own and the verdict is read
back and submitted to the portal.

## Requirements

- Python 3.11+
- [Playwright](https://playwright.dev/python/) with Firefox
  (`pip install playwright pypdf` then `playwright install firefox`)
- Firefox specifically: the rulebook requires a browser that does **not**
  auto-translate non-English pages, which rules out Chrome.

## Setup

Create `.env2` next to the script:

```
PORTAL_USERNAME=your-portal-user
PORTAL_PASSWORD=your-portal-password
```

`.env2` is git-ignored and must stay that way — it holds live credentials.

The rulebook `MASTER_RULES.md` is read from the repository root at runtime.

## Running

```
cd automation2
python website_verifier2.py                  # portal + ChatGPT (default)
python website_verifier2.py --gpt-flow       # the same, stated explicitly
python website_verifier2.py --login-only     # log in and stop
python website_verifier2.py --dump-form      # read-only dump of the portal form
python website_verifier2.py --chatgpt-login  # one-time hand sign-in to ChatGPT
```

`run2.bat` is a double-click launcher; it defaults to the ChatGPT flow and
forwards any flag you give it.

`--gpt-flow` runs until you press Ctrl+C or close the browser window. It stays
out of your way: no window is raised, and if you open the portal's Admin
Console to look at your records it pauses and leaves the page alone until a
record page is showing again.

## How `--gpt-flow` works

One Firefox window, two tabs:

| Tab | Role |
|---|---|
| 1 | The work portal, logged in from `.env2` |
| 2 | `chatgpt.com`, holding the conversation |

Per record:

1. Read the assigned URL from the portal's `#url` field.
2. Send that URL — and nothing else — into the chat. The rulebook already
   states what to do with a URL and exactly how to answer.
3. Wait for a real answer: page furniture, progress text such as "Searching
   the web", and stale answers are all rejected as not-an-answer.
4. Parse the verdict.
5. Confirm the portal form is usable **and** still showing the same record.
6. Submit, then fetch the next URL.

| Verdict | Action |
|---|---|
| `SKIP` | Website Status = *Not Working*, submit |
| The field block (`Email:` … `3+ Product Descriptions: Y`) | Fill all fields, Website Status = *Working*, submit |
| Anything unclear | Nothing submitted; a fresh chat is started and the record retried |

Nothing is ever submitted without a clear verdict, and a qualifying record with
an incomplete field block is left assigned rather than submitted on a guess.

## Recovery

Every failure recovers in place instead of ending the run. The portal keeps
serving the same URL until something is submitted for it, so a retry loses
nothing.

| Problem | Response |
|---|---|
| No answer, or an unreadable one | Fresh chat, rulebook re-loaded, record retried |
| ChatGPT outage or a bot check | Waits, then retries with backoff |
| Portal redirects to its Admin Console after a submit | Navigates back to the record page |
| Someone is *browsing* the Admin Console | Pauses without touching the page |
| Submit fails | Reloads the record page and tries again |
| Queue empty | Waits and looks again |
| Browser window closed | Ends the run with a summary |

A record that fails repeatedly gets a completely fresh chat at attempts 3, 6
and 9, and backs off to a five-minute wait beyond twelve.

## Output

`debug2/` (git-ignored) collects screenshots, portal form dumps, and
`gpt_flow_log.csv` — one row per record: `timestamp,url,verdict,outcome`.
Since this mode writes to live records, that log is the audit trail.

## Notes

- The portal's status dropdown, the Y/N fields and the 0–3 product counts were
  read off the live form rather than guessed; selectors fall back to scanning
  page content when an id changes.
- Country names are typed in the short forms the portal expects — `USA`, `UK`,
  `UAE`, `China (Hong Kong S.A.R.)`.
- Phone numbers are submitted as digits only.

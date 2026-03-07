# Email-to-Spreadsheet Logger

Automatically log Gmail messages to a Google Sheet based on filter rules. No Zapier, no third-party services — just Google Apps Script.

**Demo:** <https://youtu.be/placeholder> (coming soon)  
**Try it:** [Make a copy](https://docs.google.com/spreadsheets/d/1Rwom9A7Vn8Fc4_P9NC-9kBge_qC57tyQD28U20Sw6QY/copy) 

## What it does

- **Define rules** — Create Gmail search queries in a spreadsheet
- **Auto-categorize** — Automatically tag emails by keywords in subject/body
- **Log to sheet** — Timestamp, sender, subject, category, and direct Gmail link
- **Run on schedule** — Time-driven triggers check every 5–60 minutes
- **Idempotent** — Won't log the same email twice (uses Gmail labels for tracking)

## Why use this

- **Free** — No subscription, no limits
- **Private** — Your data stays in your Google account
- **Flexible** — Multiple filter rules, customizable fields
- **Fast** — Processes 100+ emails per run

## Quick Start (3 steps)

### 1. Copy the template

[Make a copy of this spreadsheet](https://docs.google.com/spreadsheets/d/1Rwom9A7Vn8Fc4_P9NC-9kBge_qC57tyQD28U20Sw6QY/copy)

### 2. Set up the script

1. In your copied spreadsheet, click **Extensions → Apps Script**
2. Delete the default `Code.gs` file
3. Click **+** next to **Files** → paste the contents of `Code.gs` from this repo
4. Click **Save** (floppy disk icon)
5. Click **Run** → select `setupTemplate`
6. Authorize the script when prompted (click through permissions)

### 3. Configure and run

1. Switch to the **"Rules"** tab and edit the sample Gmail search queries
2. Switch to **"Categories"** and customize auto-categorization keywords
3. Click **📧 Email Logger → ▶️ Run Logger Now** to test
4. Check the **"Log"** tab — your emails should appear!

The logger will run automatically every 15 minutes (configurable in Settings).

## Sheet Structure

| Sheet | Purpose |
|-------|---------|
| **Settings** | Polling interval, max emails, logging options |
| **Rules** | Gmail search queries to monitor |
| **Categories** | Auto-categorization keywords and priorities |
| **Log** | All logged emails with metadata |

## Sample Rules

| Enabled | Rule Name | Gmail Search Query | Description |
|---------|-----------|-------------------|-------------|
| Yes | Stripe Receipts | `from:stripe.com subject:receipt` | Payment receipts |
| Yes | Client Emails | `from:client@example.com` | Important client comms |
| No | Newsletters | `label:Newsletter` | Weekly digests (disabled) |

## Sample Categories

| Category | Keywords | Priority |
|----------|----------|----------|
| Receipts | receipt, invoice, payment, paid | 1 |
| Client | client, project, deliverable | 2 |
| Support | help, ticket, bug, issue | 3 |

Categories are checked in priority order. First match wins.

## Features

- ✅ Multiple filter rules with Gmail search syntax
- ✅ Auto-categorization by subject/body keywords
- ✅ Deduplication via Gmail labels
- ✅ Time-driven triggers (5–60 minutes)
- ✅ Direct Gmail links for each logged email
- ✅ Configurable body logging (with character limits)
- ✅ One-click template setup
- ✅ Status dashboard

## Gmail Search Tips

Gmail's search syntax is powerful. Some examples:

| Query | Finds emails... |
|-------|-----------------|
| `from:stripe.com` | From Stripe |
| `subject:invoice` | With "invoice" in subject |
| `from:client.com OR from:vendor.com` | From either domain |
| `label:receipts -label:processed` | In Receipts label, not processed |
| `newer_than:1d` | Received in last 24 hours |
| `has:attachment filename:pdf` | With PDF attachments |

## FAQ

**Q: Will this delete my emails?**  
A: No. It only reads emails and applies a "Processed" label. Original emails stay untouched.

**Q: How do I stop it?**  
A: Click **📧 Email Logger → ⏸️ Stop Auto-Logging**.

**Q: Can I change the polling interval?**  
A: Yes — edit the "Polling Interval (minutes)" value in the Settings tab, then restart the trigger.

**Q: What if I want to re-log an email?**  
A: Remove the "EmailLogger/Processed" label from the email in Gmail, then run the logger again.

**Q: Is there a limit?**  
A: Google Apps Script has daily quotas (20,000 emails/day for read operations). The default max is 100 emails per run to stay within limits.

## Troubleshooting

**"Authorization required" keeps appearing**  
- This is normal for the first run. Click through all authorization prompts.

**No emails appearing in Log**  
- Check that your Rules have "Yes" in the Enabled column
- Verify your Gmail search query works in Gmail first
- Check **View → Executions** in Apps Script for error messages

**"Exceeded maximum execution time"**  
- Reduce "Max Emails Per Run" in Settings (try 50)
- Add more specific date filters to your rules (e.g., `newer_than:7d`)

## Roadmap

- [ ] Attachment logging (link to Drive)
- [ ] Email body full-text search
- [ ] Export to CSV/Excel
- [ ] Slack/Discord notifications
- [ ] Multiple spreadsheet outputs

## License

MIT — See [LICENSE](./LICENSE)

## Credits

Built by [PressedCoffee](https://github.com/PressedCoffee) as part of the Automation Tool Loop.

---

**Found this useful?** Star ⭐ the repo and [share what you're building](mailto:Shaddock.Mercer@gmail.com).

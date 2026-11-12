# Course Reminders

Pick UBC Extended Learning courses and reminder emails go out automatically from
`chhavi.nayyar@ubc.ca`.

Live at **https://chhavi-reminders.fly.dev**

## Signing in

Anyone with a UBC email (`@ubc.ca`, including subdomains such as `@student.ubc.ca`; change with the
`ALLOWED_EMAIL_DOMAINS` env var) signs in with a one-time code or link emailed from the sender address.
Each person sees and manages only their own courses, and their reminders default to their own email.
Admins (`ADMIN_EMAILS`, default `chhavi.nayyar@ubc.ca`) also see everyone's courses and the Settings, and
can sign in with the `APP_PASSWORD` secret under "Administrator sign-in" (needed before email works).

## WordPress (UBC CMS)

The site can be embedded in `*.ubc.ca` pages (`FRAME_ANCESTORS` env var). Paste
`wordpress-embed.html` into a Custom HTML block. If UBC CMS strips the iframe, the snippet's button
still links to the site.

## Using it

Everything is on one page:

- **Your courses:** the courses you get reminders for, each upcoming session, and when the next email
  goes out. **Edit** changes who gets the emails, when, and the program name. **Remove** stops
  reminders for the course; **skip** stops them for one session.
- **Add courses:** search the Extended Learning catalogue, tick courses, choose who to send to and when
  (default 14, 7, 2 and 1 days before the start date), and click **Add**. Reminders are on immediately.
  New sessions of your courses are picked up automatically, and if Extended Learning changes a session's
  dates the reminders move with it. The course list is refreshed daily.
- **Settings:** default recipients and reminder times, the time of day emails go out (9:00 AM), and the
  UBC email password. Saving the password sends you a test email.

Reminder times are written like `14d, 7d, 2d, 1d` (d = days, w = weeks, h = hours). Add `after` for a
follow-up after the course ends, e.g. `14d after`.

Each email looks like this:

```
Subject: Course Reminder: Health Data and Visualization

Dear Team,

This is a reminder that Health Data and Visualization is starting in 2 days.

Course Details:
- Course: Health Data and Visualization
- Program: Health Data
- Start Date: 2025-10-07
- End Date: 2025-11-10
```

## Running it

Hosted on Fly.io (Toronto, always on, data on a persistent volume at `/data`).

```bash
export PATH="$HOME/.fly/bin:$PATH"
./scripts/deploy_fly.sh                                                  # test and deploy
fly secrets set APP_PASSWORD='new-site-password' --app chhavi-reminders  # website login
fly logs --app chhavi-reminders                                          # what it's doing
```

Keep it at one machine (`fly scale count 1`), since the email sender runs inside the web process.

Local development:

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python app.py            # http://127.0.0.1:5055
.venv/bin/python -m pytest
```

## Files

| File | What it does |
|------|--------------|
| `app.py` | The website |
| `catalog.py` | Reads courses, sessions and programs from extendedlearning.ubc.ca |
| `scheduler.py` | Schedules reminders and sends the ones that are due |
| `reminder_logic.py` | Reminder times and the email text |
| `mailer.py` | Sends through UBC SMTP (`smtp.mail.ubc.ca:587`) |
| `schedule_parser.py` | Date/time parsing used when reading course pages |
| `db.py` | SQLite storage |
| `legacy/` | The original Excel-driven scripts |

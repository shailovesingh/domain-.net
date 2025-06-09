import smtplib
import time
import email.utils
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import random

# TESTING_MODE=True → 10 s delays; False → real delays (60 s between sends, 1 day follow-ups)
TESTING_MODE    = False
INITIAL_GAP     = 10 if TESTING_MODE else 60      # seconds between initial emails
FOLLOWUP_DELAY  = 10 if TESTING_MODE else 86400   # seconds (1 day) before each follow-up

# Single sender configuration
SENDER = {
    "sender_email":    "neal@filldesigngroup.net",
    "sender_password": "Fdg@9874#",
    "smtp_server":     "smtp.office365.com",
    "smtp_port":       587
}

# Bounded thread-pool for scheduling follow-ups
executor = ThreadPoolExecutor(max_workers=5)

def spin_email_template(person_name, company, is_followup=False, followup_number=None):
    # -- placeholder lists: fill in your 3+ variants later --
    greeting_options = [
        f"Hi {person_name},",
        f"Hello {person_name},",
        f"Dear {person_name},"
        # Add more greeting variants here
    ]
    sentence1_options = [
        "I noticed you just booked – exciting move! Most new business owners stall at this point, but launching your website fast can help you start getting leads right away.",
        "I noticed you just locked in a slot—awesome step! Most startups hesitate now, but a swift website launch helps you generate leads from day one.",
        "I see you’ve just made your booking—fantastic step! Many entrepreneurs hit pause here, but a fast website launch means you’ll start pulling in leads immediately."
    ]
    sentence2_options = [
        "We help startups like yours go live with a professional, conversion-ready website in just 7 days – at startup-friendly pricing.",
        "We enable startups like yours to launch a polished, lead-converting website in only 7 days—without breaking the bank.",
        "In just one week, we’ll deliver a crisp, conversion-optimized site for your new business—priced perfectly for early-stage budgets.",
    ]
    sentence3_options = [
        "I recorded this short 1-minute video for you:",
        "I’ve put together a brief one-minute clip just for you:",
        "Check out this short, 1-minute walkthrough I recorded for you:"
    ]
    bullet_intro_options = [
        "It walks you through:",
        "This guide takes you step-by-step through:",
        "You’ll discover in this overview:"
    ]
    closing_options = [
        "Want to see a few recent sites we built for new businesses? Just reply \"yes\" and I’ll send them over.",
        "Interested in examples of our recent startup sites? Just type \"yes\" and I’ll forward them.",
        "Want to check out a handful of our newest business sites? Say \"yes\" and I’ll send the links."
    ]

    # pick one of each
    greeting = random.choice(greeting_options)
    s1 = random.choice(sentence1_options)
    s2 = random.choice(sentence2_options)
    s3 = random.choice(sentence3_options)
    bullet_intro = random.choice(bullet_intro_options)
    closing = random.choice(closing_options)

    loom_link = "https://www.loom.com/share/1915f664b7f145f193d7b0fd6873ecb1"
    loom_gif  = "https://cdn.loom.com/sessions/thumbnails/1915f664b7f145f193d7b0fd6873ecb1-12ee91ac978e3ba5-full-play.gif"

    text_body = f"""{greeting}

{s1}

{s2}

{s3}
🎥 Watch here: {loom_link}

{bullet_intro}
• How we help new business owners like you launch a website fast
• The top 3 things to avoid when starting your online presence
• What you can expect (cost + timeline)

Even if you’re not ready yet, it’ll give you a solid starting point.

{closing}

Best regards,
Neal
https://filldesigngroup.com/
"""

    html_body = f"""\
<html>
  <body>
    <p>{greeting}</p>
    <p>{s1}</p>
    <p>{s2}</p>
    <p>{s3}</p>
    <p>🎥 <a href=\"{loom_link}\">Watch here</a></p>
    <div>
      <a href=\"{loom_link}\">  
        <img style=\"max-width:300px;\" src=\"{loom_gif}\" alt=\"Watch Video\">  
      </a>
    </div>
    <p>{bullet_intro}</p>
    <ul>
      <li>How we help new business owners like you launch a website fast</li>
      <li>The top 3 things to avoid when starting your online presence</li>
      <li>What you can expect (cost + timeline)</li>
    </ul>
    <p>Even if you’re not ready yet, it’ll give you a solid starting point.</p>
    <p>{closing}</p>
    <p>Best regards,<br>
       Neal<br>
       <a href=\"https://filldesigngroup.com/\">Fill Design Group</a>
    </p>
  </body>
</html>
"""
    return text_body, html_body

# (rest of the script remains unchanged)

def choose_subject(company):
    templates = [
        "Question for {Company}",
        "See this for {Company}",
        "Quick Question for {Company}"
    ]
    return random.choice(templates).format(Company=company)

def check_reply(recipient_email):
    # TODO: integrate IMAP/CRM reply-checking
    return False

def send_email(to_addr, name, company, sender,
               is_followup=False, followup_number=None,
               orig_msg_id=None, orig_subject=None):
    text, html = spin_email_template(name, company, is_followup, followup_number)
    msg = MIMEMultipart('alternative')
    msg['From'] = sender['sender_email']
    msg['To']   = to_addr

    if is_followup:
        msg['Subject']     = "Re: " + orig_subject
        msg['In-Reply-To'] = orig_msg_id
        msg['References']  = orig_msg_id
    else:
        msg['Subject'] = choose_subject(company)

    msg_id = email.utils.make_msgid()
    if not is_followup:
        msg['Message-ID'] = msg_id

    msg.attach(MIMEText(text, 'plain'))
    msg.attach(MIMEText(html, 'html'))

    try:
        with smtplib.SMTP(sender['smtp_server'], sender['smtp_port'], timeout=10) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(sender['sender_email'], sender['sender_password'])
            server.sendmail(sender['sender_email'], to_addr, msg.as_string())
            tag = "Follow-up" if is_followup else "Initial"
            num = f" #{followup_number}" if is_followup else ""
            print(f"{tag}{num} email sent to {to_addr} from {sender['sender_email']}")
    except Exception as e:
        kind = "follow-up" if is_followup else "initial"
        print(f"Error sending {kind} email to {to_addr}: {e}")

    return msg_id, msg['Subject']

def followup_scheduler(to_addr, name, company, sender, orig_msg_id, orig_subject):
    # Follow-up #1
    time.sleep(FOLLOWUP_DELAY)
    if not check_reply(to_addr):
        send_email(to_addr, name, company, sender, True, 1, orig_msg_id, orig_subject)
    else:
        return
    # Follow-up #2
    time.sleep(FOLLOWUP_DELAY)
    if not check_reply(to_addr):
        send_email(to_addr, name, company, sender, True, 2, orig_msg_id, orig_subject)

def send_emails(xlsx_path):
    df = pd.read_excel(xlsx_path, engine='openpyxl')

    for _, row in df.iterrows():
        company = row['company']
        name    = row['name']
        email   = row['email']
        print(f"Processing: {company} | {name} | {email}")

        # Always use the single SENDER
        msg_id, subject = send_email(email, name, company, SENDER)

        # Schedule follow-ups via the thread-pool
        executor.submit(
            followup_scheduler,
            email, name, company, SENDER, msg_id, subject
        )

        # 60 s gap between initial emails
        time.sleep(INITIAL_GAP)

if __name__ == "__main__":
    send_emails("test Email.xlsx")

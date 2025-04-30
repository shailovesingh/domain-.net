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
    "sender_email":    "neal@filldesignprojects.website",
    "sender_password": "Fdg@9874#",
    "smtp_server":     "smtp.office365.com",
    "smtp_port":       587
}

# Bounded thread-pool for scheduling follow-ups
executor = ThreadPoolExecutor(max_workers=5)

def spin_email_template(person_name, company, is_followup=False, followup_number=None):
    greetings = [
        f"Hi {person_name},",
        f"Hello {person_name},",
        f"Dear {person_name},"
    ]
    sentence1_options = [
        "I see you booked your new domain, marking an important step toward establishing a strong online presence.",
        "I noticed you secured your new domain—an essential move toward building a reliable online identity.",
        "I noticed you secured your domain. This marks the beginning of your online journey."
    ]
    sentence2_options = [
        "In the past six months, we’ve worked with several businesses to build websites, improve their search performance, and refine their social media presence. Consider how a well-designed digital platform can support your goals.",
        "Over the past six months, we’ve assisted a number of companies with website design, search optimization, and social media strategy. Think about how a customized digital solution could benefit your business.",
        "Recently, we’ve helped several businesses develop websites, enhance their search performance, and improve their social media efforts. Imagine a digital solution that aligns with your business needs."
    ]
    sentence3_options = [
        "I’m contacting you personally to share how our services may be of benefit. Please take a moment to watch the brief video I recorded, which explains our approach.",
        "I’m contacting you directly to share more about our services. I’ve prepared a brief video introduction outlining our process.",
        "I’m reaching out personally to share how our services may help. I’ve recorded a short video to introduce myself and explain our approach."
    ]

    greeting = random.choice(greetings)
    sentence1 = random.choice(sentence1_options)
    sentence2 = random.choice(sentence2_options)
    sentence3 = random.choice(sentence3_options)
    extra = (f"\nThis is follow-up #{followup_number}. Just checking in regarding my previous email."
             if is_followup and followup_number else "")

    loom_link = "https://www.loom.com/share/35049856e0e447e8ada77a44a1297342"

    text_body = f"""{greeting}

{sentence1}

{sentence2}

{sentence3}

{extra}

{loom_link}

Looking forward to hearing from you.

Best regards,
Neal
https://filldesigngroup.com/
"""
    html_body = f"""\
<html>
  <body>
    <p>{greeting}</p>
    <p>{sentence1}</p>
    <p>{sentence2}</p>
    <p>{sentence3}</p>
    {f"<p>{extra}</p>" if extra else ""}
    <div>
      <a href="{loom_link}">
        <img style="max-width:300px;"
             src="https://cdn.loom.com/sessions/thumbnails/35049856e0e447e8ada77a44a1297342-b9abd9c74a5b4e39-full-play.gif"
             alt="Watch Video">
      </a>
    </div>
    <p>Looking forward to hearing from you.<br>
       Best regards,<br>
       Neal<br>
       <a href="https://filldesigngroup.com/">Fill Design Group</a>
    </p>
  </body>
</html>
"""
    return text_body, html_body

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

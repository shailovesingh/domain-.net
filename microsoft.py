import smtplib
import time
import random
import threading
import email.utils
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Set TESTING_MODE = True for testing (short delays) and False for production (1 and 2 days)
TESTING_MODE = False
followup_delay = 10 if TESTING_MODE else 86400  # 86400 seconds = 1 day

# List of sender accounts with their credentials and SMTP details
SENDER_ACCOUNTS = [
    {"email": "neal@filldesigngroup.net",      "password": "Fdg@9874#", "smtp_server": "smtp.office365.com", "smtp_port": 587},
    {"email": "neal@filldesignprojects.com","password": "Fdg@9874#", "smtp_server": "smtp.office365.com", "smtp_port": 587},
    {"email": "neal@filldesignprojects.website",     "password": "Fdg@9874#", "smtp_server": "smtp.office365.com", "smtp_port": 587},
]

def get_random_sender():
    return random.choice(SENDER_ACCOUNTS)

def spin_email_template(person_name, company, is_followup=False, followup_number=None):
    greetings = [f"Hi {person_name},", f"Hello {person_name},", f"Dear {person_name},"]
    sentence1 = random.choice([
        "I see you booked your new domain, marking an important step toward establishing a strong online presence.",
        "I noticed you secured your new domain—an essential move toward building a reliable online identity.",
        "I noticed you secured your domain. This marks the beginning of your online journey."
    ])
    sentence2 = random.choice([
        "Over the past few years, we’ve worked with 500+ businesses — most of them saw a 40–60% boost in leads and sales within 3 to 6 months.",
        "We’ve supported 500+ companies in the past few years — many saw a 40–60% increase in leads and sales in just 3 to 6 months",
        "Over the last few years, we've partnered with 500+ brands — most saw a 40–60% jump in leads and revenue within 3–6 months."
    ])
    sentence4 = random.choice([
        "Not sure if you're doing any marketing at the moment, but if growth is on your mind, I’d be happy to share a few ideas that might help.",
        "Not sure if you’re focused on marketing right now, but if growth is a goal, I’d love to share a few helpful ideas.",
        "If you’re thinking about growing, I’d be happy to pass along a few ideas that could help."
    ])
    sentence3 = random.choice([
        "I’m contacting you personally to share how our services may be of benefit. Please take a moment to watch the brief video I recorded, which explains our approach.",
        "I’m contacting you directly to share more about our services. I’ve prepared a brief video introduction outlining our process.",
        "I’m reaching out personally to share how our services may help. I’ve recorded a short video to introduce myself and explain our approach."
    ])
    extra = f"\nThis is follow-up #{followup_number}. Just checking in regarding my previous email." if is_followup and followup_number else ""
    loom_link = "https://www.loom.com/share/1915f664b7f145f193d7b0fd6873ecb1"

    text = f"""{random.choice(greetings)}

{sentence1}

{sentence2}

{sentence4}

{sentence3}

{extra}

{loom_link}

Looking forward to hearing from you.

Best regards,
Neal
https://filldesigngroup.com/
"""
    html = f"""\
<html><body>
  <p>{random.choice(greetings)}</p>
  <p>{sentence1}</p>
  <p>{sentence2}</p>
  <p>{sentence4}</p>
  <p>{sentence3}</p>
  {f"<p>{extra}</p>" if extra else ""}
  <div>
    <a href="{loom_link}">
      <img style="max-width:300px;" src="https://cdn.loom.com/sessions/thumbnails/1915f664b7f145f193d7b0fd6873ecb1-12ee91ac978e3ba5-full-play.gif" alt="Watch Video">
    </a>
  </div>
  <p>Looking forward to hearing from you.<br>Best regards,<br>Neal<br>
     <a href="https://filldesigngroup.com/">Fill Design Group</a></p>
</body></html>
"""
    return text, html

def choose_subject(company):
    return random.choice([
        "Question for {Company}",
        "See this for {Company}",
        "Quick Question for {Company}"
    ]).format(Company=company)

def check_reply(email_address):
    return False  # placeholder

def send_initial_email(row):
    company = row['company']
    name    = row['name']
    to_addr = row['email']
    subject = choose_subject(company)
    text, html = spin_email_template(name, company)

    sender = get_random_sender()
    msg = MIMEMultipart('alternative')
    msg['From'] = sender['email']
    msg['To']   = to_addr
    msg['Subject'] = subject
    msg_id = email.utils.make_msgid()
    msg['Message-ID'] = msg_id
    msg.attach(MIMEText(text, 'plain'))
    msg.attach(MIMEText(html, 'html'))

    try:
        with smtplib.SMTP(sender['smtp_server'], sender['smtp_port'], timeout=10) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(sender['email'], sender['password'])
            server.sendmail(sender['email'], to_addr, msg.as_string())
            print(f"Initial email sent to {to_addr} from {sender['email']}")
    except Exception as e:
        print(f"Error sending to {to_addr}: {e}")
        return None, None, None

    return msg_id, subject, sender

def send_followup(to_addr, msg_id, name, company, num, sender, orig_subj):
    text, html = spin_email_template(name, company, True, num)
    msg = MIMEMultipart('alternative')
    msg['From'] = sender['email']
    msg['To']   = to_addr
    msg['Subject'] = "Re: " + orig_subj
    msg['In-Reply-To'] = msg_id
    msg['References']  = msg_id
    msg.attach(MIMEText(text, 'plain'))
    msg.attach(MIMEText(html, 'html'))

    try:
        with smtplib.SMTP(sender['smtp_server'], sender['smtp_port'], timeout=10) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(sender['email'], sender['password'])
            server.sendmail(sender['email'], to_addr, msg.as_string())
            print(f"Follow-up #{num} sent to {to_addr}")
    except Exception as e:
        print(f"Error sending follow-up #{num} to {to_addr}: {e}")

def followup_scheduler(to_addr, msg_id, name, company, sender, subj):
    time.sleep(followup_delay)
    if not check_reply(to_addr):
        send_followup(to_addr, msg_id, name, company, 1, sender, subj)
    else:
        return
    time.sleep(followup_delay)
    if not check_reply(to_addr):
        send_followup(to_addr, msg_id, name, company, 2, sender, subj)

def send_emails(xlsx_path):
    # Read Excel into DataFrame
    df = pd.read_excel(xlsx_path, engine='openpyxl')
    # Expect columns: company, name, email
    for _, row in df.iterrows():
        company = row['company']
        name    = row['name']
        email   = row['email']
        print(f"Processing: {company} | {name} | {email}")

        msg_id, subj, sender = send_initial_email(row)
        if not msg_id:
            continue

        threading.Thread(
            target=followup_scheduler,
            args=(email, msg_id, name, company, sender, subj)
        ).start()

        time.sleep(60)  # pause between sends

if __name__ == "__main__":
    send_emails("test Email.xlsx")

import time
import codecs
import smtplib
import datetime
import email
import optparse
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.encoders import encode_base64
from email.mime.multipart import MIMEMultipart
from email.utils import COMMASPACE, formatdate
from getpass import getpass


# smtp settings
SERVER = ''
PORT = 0
USER_EMAIL = ""
USER_PASS = ""

# email settings
EMAIL_SUBJECT = ""

# event settings
EVENT_DESCRIPTION = ""
EVENT_SUMMARY = ""

ORGANIZER_NAME = ""
ORGANIZER_EMAIL = ""
ATTENDEES = []

# template settings
EVENT_TEXT = ""
EVENT_URL = ""

def setup_args():
    parser = optparse.OptionParser()

    parser.add_option("--organizer-name", dest="org_name", type='string', help='Set the organizer name')
    parser.add_option("--organizer-email", dest="org_email", type='string', help='Set the organizer email')
    parser.add_option("--email-subject", dest="email_subject", type="string", help='Email Subject')
    parser.add_option("--event-description", dest="event_desc", type="string", help="Event description")
    parser.add_option("--event-summary", dest="event_summary", type="string", help="Event Summary")
    parser.add_option("--event-text", dest="event_text", type="string", help="Event text")
    parser.add_option("--phishing-url", dest="ph_url", type="string", help="Set the phishng url in \"Join to the meeting\". Format: https://<url>")
    parser.add_option("--from", dest="em_from", type="string", help="From email")
    parser.add_option("--to", dest="to", type="string", help="To email")
    parser.add_option("--attendees", dest="att", type="string", help="Attendees comma sepparated")
    parser.add_option("--output", dest="output", type="string", help="Output eml file")
    parser.add_option("--send-email", dest="send", action="store_false", help="Send the email, credentials are asked while using this option")

    return parser.parse_args()

def load_template():
    template = ""
    with codecs.open("email_template.html", 'r', 'utf-8') as f:
        template = f.read()
    return template


def prepare_template():
    email_template = load_template()
    email_template = email_template.format(EVENT_TEXT=EVENT_TEXT, EVENT_URL=EVENT_URL)
    return email_template


def load_ics():
    ics = ""
    with codecs.open("iCalendar_template.ics", 'r', 'utf-8') as f:
        ics = f.read()
    return ics


def prepare_ics(dtstamp, dtstart, dtend):
    ics_template = load_ics()
    ics_template = ics_template.format(DTSTAMP=dtstamp, DTSTART=dtstart, DTEND=dtend, ORGANIZER_NAME=ORGANIZER_NAME, ORGANIZER_EMAIL=ORGANIZER_EMAIL, DESCRIPTION=EVENT_DESCRIPTION, SUMMARY=EVENT_SUMMARY, ATTENDEES=generate_attendees())
    return ics_template


def generate_attendees():
    attendees = []
    for attendee in ATTENDEES:
        attendees.append("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=ACCEPTED;RSVP=FALSE\r\n ;CN={attendee};X-NUM-GUESTS=0:\r\n mailto:{attendee}".format(attendee=attendee))
    return "\r\n".join(attendees)


def create_email(to, opts):
    print("Creating the template using this parameters: {}".format(ots))
    

    # in .ics file timezone is set to be utc
    utc_offset = time.localtime().tm_gmtoff / 60
    ddtstart = datetime.datetime.now()
    dtoff = datetime.timedelta(minutes=utc_offset + 5) # meeting has started 5 minutes ago
    duration = datetime.timedelta(hours = 1)  # meeting duration
    ddtstart = ddtstart - dtoff
    dtend = ddtstart + duration
    dtstamp = datetime.datetime.now().strftime("%Y%m%dT%H%M%SZ")
    dtstart = ddtstart.strftime("%Y%m%dT%H%M%SZ")
    dtend = dtend.strftime("%Y%m%dT%H%M%SZ")

    ics = prepare_ics(dtstamp, dtstart, dtend)

    email_body = prepare_template()

    msg = MIMEMultipart('mixed')
    msg['Reply-To']=USER_EMAIL
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = EMAIL_SUBJECT
    msg['From'] = USER_EMAIL
    msg['To'] = to + ", ".join(ATTENDEES)

    part_email = MIMEText(email_body,"html")
    part_cal = MIMEText(ics,'calendar;method=REQUEST')

    msgAlternative = MIMEMultipart('alternative')
    msg.attach(msgAlternative)

    ics_atch = MIMEBase('application/ics',' ;name="%s"' % ("invite.ics"))
    ics_atch.set_payload(ics)
    encode_base64(ics_atch)
    ics_atch.add_header('Content-Disposition', 'attachment; filename="%s"' % ("invite.ics"))

    eml_atch = MIMEBase('text/plain','')
    eml_atch.set_payload("")
    encode_base64(eml_atch)
    eml_atch.add_header('Content-Transfer-Encoding', "")

    msgAlternative.attach(part_email)
    msgAlternative.attach(part_cal)

    return msg

def send_email(msg):
    mailServer = smtplib.SMTP(SERVER, PORT)
    mailServer.ehlo()
    mailServer.starttls()
    mailServer.ehlo()
    mailServer.login(USER_EMAIL, USER_PASS)
    mailServer.sendmail(USER_EMAIL, to, msg.as_string())
    mailServer.close()

def main():
    (opts, args) = setup_args()
    

    # email settings
    EMAIL_SUBJECT = opts.email_subject
    EMAIL_FROM = opts.em_from
    EMAIL_TO = opts.to

    # event settings
    EVENT_DESCRIPTION = opts.event_desc
    EVENT_SUMMARY = opts.event_summary

    ORGANIZER_NAME = opts.org_name
    ORGANIZER_EMAIL = opts.org_email
    ATTENDEES = opts.att.replace(' ','').split(',')

    # template settings
    EVENT_TEXT = opts.event_text
    EVENT_URL = opts.ph_url
    
    print("This tool is created for internal audits or pentesters, do not use it for harm or impersonating people to take profit")
    print("This tool is created by ExAndroidDev and edited by Davidc96, all credits to him")
    msg = create_email(EMAIL_TO, opts)
    
    if opts.output:
        with open(opts.output, 'w') as f:
            gen = email.generator.Generator(f)
            gen.flatten(msg)
    
    if opts.send:
        SERVER = input("SMTP Server: ")
        PORT = input("SMTP Port: ")
        USER_EMAIL = input("Email Username: ")
        USER_PASS = getpass("Email password: ")

        send_email(msg)


main()

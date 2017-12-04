#!/usr/bin/env python3

import os
import argparse
import yaml
import base64

from random import shuffle
from datetime import datetime
from googleapiclient import discovery, errors
from email.mime.text import MIMEText
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/gmail-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/gmail.send https://www.googleapis.com/auth/gmail.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Secret Santa Randomizer'
SUBJECT = "Your {:%Y} Secret Santa Drawing!"
BOILERPLATE = """<p>Hi, {fr:s}!</p>

<p>This email is to tell you the results of the automated Secret Santa drawing for
this year.  This is an automated message:  the drawing was completely
randomized, and the results are unknown to me.</p>

<p>{fr:s}, you are Secret Santa for:</p>

<p style="font-weight:bold;font-size:125%">{t:s}</p>

<p>Have a great holiday season!</p>
"""


def get_credentials(args):
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'secret-santa.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if args:
            credentials = tools.run_flow(flow, store, args)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def parse_yaml(arg):
    """Opens a given file name and parses the content for YAML.

    Returns an object parsed from the YAML file (typically a map)
    """
    try:
        with open(arg, 'r') as f:
            return yaml.load(f)
    except OSError as err:
        raise argparse.ArgumentTypeError("Unable to read file '{:s}': {}".format(arg, err))
    except yaml.YAMLError as err:
        exc = "Error parsing file `{:s}`"
        if hasattr(err, 'problem_mark'):
            exc += ": line {:d}, column {:d}".format(err.problem_mark.line+1, err.problem_mark.column+1)
        raise argparse.ArgumentTypeError(exc)

def main():
    """Shows basic usage of the Gmail API.

    Creates a Gmail API service object and outputs a list of label names
    of the user's Gmail account.
    """

    parser = argparse.ArgumentParser(description="Randomize Secret Santa draws using defined rules", parents=[tools.argparser])

    parser.add_argument("config", help="Configuration YAML file", type=parse_yaml)
    parser.add_argument("--dry", "-d", help="Dry run: do not send emails, just report intended actions and exit", action="store_true")
    parser.add_argument("--tries", "-N", help="Number of tries to shuffle the email permutations before giving up", type=int, default=100)
    parser.add_argument("--allowself", help="Allow self-gifting (typically not allowed in Secret Santa processes)", action="store_true")

    args = parser.parse_args()
    dry_run = args.dry
    tries = args.tries
    config = args.config
    allow_self = args.allowself

    credentials = get_credentials(args)
    service = discovery.build('gmail', 'v1', credentials=credentials)

    emails = config['emails']
    users = emails.keys()
    forbidden = config['forbidden']
    if not allow_self:
        for u in users:
            forbidden[u] = forbidden.get(u, []) + [u]

    users_from = list(users)
    users_to = list(users)

    for i in range(tries):
        shuffle(users_to)
        bad = False
        for uf, ut in zip(users_from, users_to):
            if uf in forbidden and ut in forbidden[uf]:
                bad = True
                break
        if bad:
            continue
        print("Good shuffle found after {:d} iterations".format(i+1))
        break

    if i == tries - 1 and bad:
        raise Exception("Unable to generate a good shuffle after {:d} attempts: consider increasing 'tries'".format(tries))

    if dry_run:
        print([(f, t) for f, t in zip(users_from, users_to)])
        from_email = "fake@fake.fake"
    else:
        resp = service.users().getProfile(userId="me").execute()
        from_email = resp['emailAddress']

    for uf, ut in zip(users_from, users_to):
        message = MIMEText(BOILERPLATE.format(fr=uf, t=ut), 'html')
        message['to'] = '"{:s}" <{:s}>'.format(uf, emails[uf])
        message['from'] = from_email
        message['subject'] = SUBJECT.format(datetime.now())

        if dry_run:
            print(message)
            continue

        message_obj = {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')}
        resp = service.users().messages().send(userId='me', body=message_obj).execute()
        print("Message ID sent: {:s}".format(resp['id']))

    print("Done")


if __name__ == '__main__':
    main()

"""
This Python script pulls the first 20 unread emails, matches a set of rules, then forwards.

You can then run this sample with a JSON configuration file:
    python handle_email.py parameters.json
"""

import sys
import json
import logging
import requests
import msal
import time

def printResponseError(response):
    print(response.status_code)
    print(response.reason)
    print(response.text)

def sendDraft(config, draft_id, access_token):
    print("Sending draft")
    auth_header = {'Authorization': 'Bearer ' + access_token}
    url = config["postSendEndpoint"].format(draft_id)

    response = requests.post(url, headers=auth_header)
    if response.ok:
        print('sent!')
    else:
        printResponseError(response)

def updateDraft(config, draft_id, person, access_token):
    print("Updating draft to send to person {}".format(person['nick_name']))
    auth_header = {'Authorization': 'Bearer ' + access_token, "Content-Type": "application/json"}
    url = config["patchMessageEndpoint"].format(draft_id)
    payload = config['patchMessagePayload']
    payload['toRecipients'][0]['emailAddress']['address'] = person['email']
    # WARNING - UPDATING THE BODY REPLACES THE EXISTING FOWARDED TEXT
    payload['body']['content'] = 'Hi {0}, look at this!'.format(person['nick_name'])

    response = requests.patch(url, json.dumps(payload), headers=auth_header)
    if response.ok:
        if config["sendEmailAfterDrafting"]:
            print('sending email...')
            sendDraft(config, draft_id, access_token)
        else:
            print('NOT SENDING - sendEmailAfterDrafting is False, look in your drafts folder')

    else:
        printResponseError(response)

def forwardEmail(config, email_id, access_token):
    print("Forwarding an email - creating email in drafts folder")
    url = config["postForwardEndpoint"].format(email_id)
    auth_header = {'Authorization': 'Bearer ' + access_token}
    response = requests.post(url, headers=auth_header)
    if response.ok:
        emailForwardJsonObject = json.loads(response.text)
        body = emailForwardJsonObject['body']['content']
        email_id = emailForwardJsonObject['id']

        for person in config['people']:
            # use address book in the config, if name is in the body then forward to that person
            if person['name'] in body:
                updateDraft(config, email_id, person, access_token)
    else:
        printResponseError(response)

# load the config file
config = json.load(open(sys.argv[1]))

# configure the application to connect to Azure Active Directory to sign in
app = msal.PublicClientApplication(config["client_id"], authority=config["authority"])

# start the sign in flow
flow = app.initiate_device_flow(scopes=config["scope"])
if "user_code" not in flow:
    raise ValueError(
        "Fail to create device flow. Err: %s" % json.dumps(flow, indent=4))

# print instructions to the console telling the user how to sign in
print(flow["message"])
sys.stdout.flush()

# wait until user signs in using the browser
result = app.acquire_token_by_device_flow(flow)

while True:
    # verify that sign in was successful
    if "access_token" in result:
        access_token = result['access_token']
        
        # Call Microsoft Graph to GET email using the access_token
        url = config["endpoint"]
        auth_header = {'Authorization': 'Bearer ' + access_token}
        response = requests.get(url, headers=auth_header,)
        if response.ok:
            emailResponseJsonObject = json.loads(response.text)

            print("found {} unread emails".format(len(emailResponseJsonObject['value'])))
            # enumerate all the email, make a decision if we need to take action on this email
            for email in emailResponseJsonObject['value']:
                fromEmail = email['from']['emailAddress']['address']
                body = email['body']['content']
                subject = email['subject']

                # uncomment these if you need to email debug matching logic
                # print(fromEmail)
                # print(bodyText)
                # print(subject)

                if ("o365mc@microsoft.com" in fromEmail and "Yammer Communities" in body):
                    print("found email from o365mc@microsoft.com with Yammer Communities")
                    forwardEmail(config, email['id'], access_token)
        else:
            printResponseError(response)
    else:
        print(result.get("error"))
        print(result.get("error_description"))
        print(result.get("correlation_id"))
        exit()
    print("Done processing, waiting for 1 min")
    time.sleep( 1 * 60 )
    print("Awake, attempting to review email : %s" % time.ctime())
    accounts = app.get_accounts()
    # started with empty in memory account cache, so only one account signed in
    print("Getting access_token for account {}".format(accounts[0]["username"]))
    result = app.acquire_token_silent(config["scope"], account=accounts[0])
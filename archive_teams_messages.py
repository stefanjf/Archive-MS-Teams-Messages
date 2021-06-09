
import requests, json, datetime, re, dateutil.parser, sys
from yaspin import yaspin

import os.path
from os import path

this = sys.modules[__name__]

this.token = ''

if (path.isfile('token')):
    print('Found token file')
    with open('token') as f:
        this.token = f.read().strip()
else:
    print('No token file found')

def setToken(mytoken):
    this.token = mytoken

def listMyTeams():
    response = requests.get('https://graph.microsoft.com/beta/me/joinedTeams', headers={'Authorization': this.token})
    if (response.status_code == 200):
        teams = json.loads(response.content)
        for team in teams['value']:
            print(team['id'] + '   ' + team['displayName'])
            responseChannels = requests.get('https://graph.microsoft.com/beta/teams/' + team['id'] + '/channels', headers={'Authorization': this.token})
            if (responseChannels.status_code == 200):
                channels = json.loads(responseChannels.content)
                for channel in channels['value']:
                    print('\t' + channel['id'] + '   ' + channel['displayName'])
    else:
        print('Error: ' + str(response))
        

def pullAllChannelMessagesInGroup(groupID):
    response = requests.get('https://graph.microsoft.com/beta/teams/' + groupID + '/channels', headers={'Authorization': this.token})
    if (response.status_code == 200):
        channels = json.loads(response.content)
        for channel in channels['value']:
            print(channel['id'] + '   ' + channel['displayName'])

            # Get JSON and save to file
            chat = pullMessagesIntoJSON(groupID, channel['id'])
            f = open(channel['displayName'] + "_output.json", "w")
            f.write(json.dumps(chat))
            f.close()
            print('Saved raw JSON to ' + str(channel['displayName']) + "_output.json")

            # Parse into HTML and save to file
            chatHTML = parseJSONintoHTML(chat)
            f = open(channel['displayName'] + "_output.html", "w")
            f.write(json.dumps(chatHTML))
            f.close()
            print('Saved HTML to ' + str(channel['displayName']) + "_output.html")

        print('All done with ' + str(channel['displayName']))
    else:
        print('Error: ' + str(response))
  
def pullSingleChannelMessagesInGroup(groupID, channelID):
        print('Collecting messages from ' + channelID)

        # Get JSON and save to file
        chat = pullMessagesIntoJSON(groupID, channelID)
        f = open("myChannel_output.json", "w")
        f.write(json.dumps(chat))
        f.close()
        print("Saved raw JSON to myChannel_output.json")

        # Parse into HTML and save to file
        chatHTML = parseJSONintoHTML(chat)
        f = open("myChannel_output.html", "w")
        f.write(json.dumps(chatHTML))
        f.close()
        print("Saved HTML to myChannel_output.html")


def pullfromAPI(url):
    # print(url)
    response = requests.get(url, headers={'Authorization': this.token})
    messages = json.loads(response.content)
    # print("Pulled " +  " items.")
    return messages

@yaspin(text="Downloading messages...")
def pullMessagesIntoJSON(_groupID, _channelID):
    # Gather all messages
    allMessagesRaw = []
    linkToNextBatch = ""

    # Get list of channels
    messages = pullfromAPI('https://graph.microsoft.com/beta/teams/' + _groupID + '/channels/' + _channelID + '/messages?$top=100')
    for item in messages["value"]: allMessagesRaw.append(item)

    # If there's another batch
    if "@odata.nextLink" in messages.keys():
        linkToNextBatch = messages["@odata.nextLink"]
        # print('Another batch available')

        while True:
            messages = pullfromAPI(linkToNextBatch)
            for item in messages["value"]: allMessagesRaw.append(item)

            if "@odata.nextLink" in messages.keys():
                linkToNextBatch = messages["@odata.nextLink"]
                # print('Another batch available')
            else:
                break

    # print('Done with pulling messages! Now adding in replies...')

    # For each message, pull replies and add to dict
    for msg in allMessagesRaw: 
        replies = pullfromAPI('https://graph.microsoft.com/beta/teams/' + _groupID + '/channels/' + _channelID + '/messages/' + msg["id"] + '/replies')
        msg["replies"] = []

        if replies['@odata.count'] > 0:
            for reply in replies["value"]: msg["replies"].append(reply)

            # Check if more
            if "@odata.nextLink" in replies.keys():
                while True:
                    linkToNextBatch = replies["@odata.nextLink"]
                    # print('Another batch available')
                    replies = pullfromAPI(linkToNextBatch)
                    for reply in replies["value"]: msg["replies"].append(reply)

                    # If more replies, repeat, otherwise break
                    if "@odata.nextLink" in replies.keys():
                        linkToNextBatch = replies["@odata.nextLink"]
                        # print('Another batch available')
                    else:
                        break
    print('Collected ' + str(len(allMessagesRaw)) + ' messages!')
    return allMessagesRaw

def parseJSONintoHTML(jsonChatMessages):
    finalHTMLOutput = ""

    jsonChatMessages.sort(key=lambda x: dateutil.parser.isoparse(x['createdDateTime']))

    for msg in jsonChatMessages: 
        b = msg['body']['content']
        if b is not None:
            b = re.sub('\n+', '', b)
            b = re.sub('\t+', '', b)

            finalHTMLOutput += '<hr><hr><h3>' + msg['from']['user']['displayName'] + ':</h3><h5>Created: ' + msg["createdDateTime"] + '</h5>' + b + '<blockquote>'

        msg['replies'].sort(key=lambda x: dateutil.parser.isoparse(x['createdDateTime']))
        for reply in msg['replies']:
            try:
                user = str(reply['from']['user']['displayName'])
            except TypeError:
                user = 'unknown'
                pass

            try:
                replyContent = reply['body']['content']
                replyContent = re.sub('\n+', '', replyContent)
                replyContent = re.sub('\t+', '', replyContent)
            except TypeError:
                replyContent = 'unknown'
                pass

            finalHTMLOutput += '<h3>Reply From: ' + user + '</h3>' + '<h5>Created: ' + reply["createdDateTime"] + '</h5>' + replyContent

        finalHTMLOutput += '</blockquote>'
    return finalHTMLOutput


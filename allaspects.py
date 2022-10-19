'''

Author       : Anuraj Pilanku

Code Utility :Monitor folder in  Outlook and update worklog in HPSM Incident

'''

import os
import json
import requests
import time
from exchangelib.protocol import BaseProtocol
from exchangelib import Credentials, Account, DELEGATE, Configuration, FileAttachment
import requests
from bs4 import BeautifulSoup, BeautifulStoneSoup
from robot.api import logger

jsonpath=r"C:\Users\USSACDev\Desktop\anuraj\monitorMailbox\itsmworkloginputs.json"
passwordUpdate='''"C:\Program Files (x86)\CyberArk\ApplicationPasswordSdk\CLIPasswordSDK.exe>" GetPassword /p AppDescs.AppID=APP_ADOE-CAC /p Query="Safe=WW-TTS-EES-AUTOCNTR-AD;Folder=Root;Object=mmm.com_{userid}" /o Password'''
readjson = open(jsonpath)
jsoninputs = json.load(readjson)


# ITSMworklogupdate-foldername


class ProxyAdapter(requests.adapters.HTTPAdapter):
    def send(self, *args, **kwargs):
        kwargs["proxies"] = ProxyAdapter.proxies
        return super(ProxyAdapter, self).send(*args, **kwargs)


class monitorMail():
    def _connect(self):  # (self, aeuser, aepassword, aeci, aeparameters):
        '''
        Method for creating connection
        and configuration to monitor mailbox
        '''
        o365mailid="ussacdev@mmm.com"
        mailusername=jsoninputs["user1"].strip()
        proxy = {}
        proxy.update({})  # (aeci.get("proxy", {}))
        ProxyAdapter.proxies = proxy
        BaseProtocol.HTTP_ADAPTER_CLS = ProxyAdapter
        smtp_address = "ussacdev@mmm.com"  # "mailserv.mmm.com"#aeci.get("primarysmtpaddress", aeuser)
        if not smtp_address:
            smtp_address = "ussacdev@mmm.com"  # aeuser
        server = "Outlook.Office365.com"  # "mailserv.mmm.com"# aeci.get("server", "Outlook.Office365.com")
        if not server:
            server = "Outlook.Office365.com"
        credentials = Credentials(username=o365mailid,password=os.popen(passwordUpdate.format(userid=mailusername)).read().strip())#",=I}-)7(}4Uf_%</Mrl/")  # password="vLcRU0%@ZnUd+Xy(.a5w")
        config = Configuration(
            server=server,
            credentials=credentials)
        account = Account(
            primary_smtp_address=smtp_address,
            credentials=credentials,
            autodiscover=False,
            config=config,
            access_type=DELEGATE)
        return account


mailfolder = monitorMail()._connect().root
ourfolder = monitorMail()._connect()

folder_name = "ITSMworklogupdate"  # "ApplensSMO"#aeparameters.get("folder_name")
if folder_name:
    my_folder = ourfolder.root / "Top of Information Store"
    for folder in folder_name.split("/"):
        my_folder = my_folder / folder
    filter_criteria = 'isRead:False'  # AND ({0})'.format(createactivitykeywords)
    items = my_folder.filter(filter_criteria)
else:
    filter_criteria = 'isRead:False'  # AND ({0})'.format(createactivitykeywords)
    items = ourfolder.inbox.filter(filter_criteria)
result = {}
result["items"] = items

filter_criteria = 'isRead:False'
# items = mailfolder.filter(filter_criteria)
# result=dict()
# result["items"] = items
print(items, result)
# count = 0
payload = {}
mailCollections = {}
count = 0
for item in items.__iter__():
    soup = BeautifulSoup(item.text_body, 'html.parser')
    body = soup.text.replace('\n', "").replace('\r', "").strip()
    subject = item.subject.replace('\n', "").replace('\r', "").strip()  # ID:IM43545678 CODE:CACHPSMO365AUTO
    mailaddress = item.author.email_address.replace('\n', "").replace('\r', "").strip()
    count += 1
    mailCollections["mail_{0}".format(str(count))] = {"subject": subject.strip(), "mailaddress": mailaddress.strip(),
                                                      "body": body.strip()}
    # to make mail as read!!
    item.is_read = True
    item.save()
    # Carboncopy=item.cc_recipients
    # payload = {}
    '''payload["isreply"] = False
    if item.in_reply_to:
        payload["isreply"] = True
    payload["sender"] = str(item.author.email_address)
    if(item.cc_recipients):
        cc_len = item.cc_recipients
        i = 0
        while i < len(cc_len):
            CC = str(item.cc_recipients[i].email_address)
            i +=1
            payload.setdefault("cc",[]).append(CC)
    else:
        payload["cc"]=""
    str1=","
    payload["cc"]=str1.join(payload["cc"])
    #payload["notify_on_success"] = aeparameters.get("notify_on_success", "true")
    #payload["notify_on_failure"] = aeparameters.get("notify_on_failure", "true")'''
    # payload["subject"] = item.subject
    # print(body,subject,mailaddress,C)#(payload)
# print(mailCollections)


username =jsoninputs["user2"].strip()#"USSACITSMDev"
password =os.popen(passwordUpdate.format(userid=username)).read().strip() #"Ts)Q5Ll3N-ICF_pa9hcl"  # "wmkJjlrO0Ew>cObVpP_<"
host = "http://itsmwsqa.mmm.com/"  # "http://itsmws.mmm.com/"
table = "AC_IncidentManagement"


def postconnect():
    success_counter = 0
    # print("url===", url)
    http_proxy = os.getenv('http_proxy', None)
    https_proxy = os.getenv('https_proxy', None)
    no_proxy = os.getenv('no_proxy', None)
    proxies = {"http": http_proxy, "https": https_proxy, "no": no_proxy}
    for maildetails in range(0, len(mailCollections)):
        mailsubject = mailCollections[list(mailCollections.keys())[maildetails]]["subject"].strip()
        incidentID = mailsubject[mailsubject.index("ID:") + len("ID:"):mailsubject.index("CODE:")].strip()
        mailbody = mailCollections[list(mailCollections.keys())[maildetails]]["body"].strip()
        #remove unnecessary data from mailbody
        if "This e-mail and any files transmitted" in mailbody:
            mailbody=mailbody[:mailbody.index("This e-mail and any files transmitted")]
        if "From:" and "Sent:" in mailbody:
            mailbody=mailbody[:mailbody.index("From:")]
        print(mailbody)
        statusnum=mailsubject[mailsubject.index("O365AUTO(")+len("O365AUTO("):mailsubject.index("O365AUTO(")+len("O365AUTO(")+2].strip().replace(")","")
        ticketstatus=jsoninputs["ticketStatus"][mailsubject[mailsubject.index("O365AUTO(")+len("O365AUTO("):mailsubject.index("O365AUTO(")+len("O365AUTO(")+2].strip().replace(")","")]
        print(ticketstatus)
        ID = incidentID
        if int(statusnum) in [1,2,3,4] and "Assignee_ID:" in mailbody:#change status
            Assignee=mailbody[mailbody.index("Assignee_ID:")+len("Assignee_ID"):mailbody.index(")]}")]
            url_statusUpdate = os.path.join(str(host), 'SM/9/rest/' + str(table) + '/' + str(ID))
            data_statusUpdate = {'Incident': {'Status': ticketstatus,'Assignee': Assignee}}  # 'Pending Emergency Change'}}#ticketstatus#'a9ht9zz'
            body_status = json.dumps(data_statusUpdate)
            response_status = requests.post(url_statusUpdate, proxies=proxies, auth=(username, password), data=body_status, headers={"Connection": "close"})
            print("StatusChange_status_code===", response_status.status_code)
        elif int(statusnum) in [0]:
            url = os.path.join(str(host), 'SM/9/rest/' + str(table) + '/' + str(ID) + '?view=expand')
            data = {'Incident': {"JournalUpdates": mailbody, "Type": "Communication with customer"}}
            body_worklog = json.dumps(data)
            response_worklog = requests.post(url, proxies=proxies, auth=(username, password), data=body_worklog,headers={"Connection": "close"})
            print("worklogUpdate_status_code===",response_worklog.status_code)
        elif int(statusnum) in [5] and "ClosureCode" in mailbody:
            #logger.debug("!!!!!!!!!!!!!!! "+str(statusnum)+"mailbody:"+str(mailbody))
            print("!!!!!!!!!!!!!!! "+str(statusnum)+"mailbody:"+str(mailbody))
            ClosureCode=mailbody[mailbody.index("ClosureCode:")+len("ClosureCode:"):mailbody.index("Solution:")].strip()#{[(Assignee_ID:ac5qzz ClosureCode:let it be Solution:give solution)]}
            Solution=mailbody[mailbody.index("Solution:")+len("Solution:"):mailbody.index(")]}")].strip()
            url = os.path.join(str(host), 'SM/9/rest/' + str(table) + '/' + str(ID) + '?view=expand')
            data_closeIncident = {'Incident': {'ClosureCode': ClosureCode, 'Solution': Solution}}
            body_closeIncident = json.dumps(data_closeIncident)
            #response_closeIncident = requests.post(url, proxies=proxies, auth=(username, password),data=body_closeIncident, headers={"Connection": "close"})
            response_closeIncident = requests.put(url, proxies=proxies, auth=(username, password),data=body_closeIncident, headers={"Connection": "close"})#,"Content-Type": "application/json"})
            print("CloseIncident_status_code===", response_closeIncident.status_code)
            print(str(response_closeIncident.json()['Messages']))

postconnect()
#print(mailCollections)



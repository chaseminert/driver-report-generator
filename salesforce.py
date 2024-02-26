import requests.exceptions
from simple_salesforce import Salesforce
from simple_salesforce import SalesforceLogin
import time

DOT_NAME = "USDOT__c"
POC_EMAIL_NAME = "POC_Email__c"
CONTACT_EMAIL_NAME = "Email"
POC_ID = "0125Y000001THZvQAO"
OWNER_OPERATOR_ID = "0125Y000001THa0QAG"
OWNER_ID = "0125Y000001TJipQAG"
RECORD_TYPE_ID = "RecordTypeId"
NO_EMAIL_IDENTIFIER = "NO EMAIL"

RECORD_TYPE_PREFERENCE = [POC_ID, OWNER_OPERATOR_ID, OWNER_ID]  # the order to get the emails


# search for POC contact, then the owner operator, than owner

class SalesForce:
    def __init__(self, username: str, password: str):
        self._sf = None
        self._username = username
        self._password = password

    def login(self):
        session_id, instance = SalesforceLogin(username=self._username, password=self._password)
        while True:
            try:
                self._sf = Salesforce(instance=instance, session_id=session_id)
                return
            except requests.exceptions.ConnectionError:
                print("Retrying salesforce login...")
                time.sleep(5)
                continue

    '''
    This method searches for the POC email and then if it can't find it it queries the contacts and looks for the POC
    Email there. If it can't find it, it looks for the owner operator contact and tries to grab their email. If it
    still can't find it, it searches the owner contact
    
    '''

    def get_email(self, dot_number):
        query1 = self._sf.query(f"SELECT {POC_EMAIL_NAME} FROM Account WHERE {DOT_NAME} = '{dot_number}'")["records"]
        if query1:
            query1 = query1[0]
            email1 = query1[POC_EMAIL_NAME]
            if email1 is not None:
                return email1
        email_dict = dict()  # maps the contact record id to an email address
        result2_query = (
            f"SELECT {CONTACT_EMAIL_NAME}, {RECORD_TYPE_ID} FROM Contact WHERE Account.{DOT_NAME} = '{dot_number}' "
            f"AND {RECORD_TYPE_ID} IN ('{RECORD_TYPE_PREFERENCE[0]}', '{RECORD_TYPE_PREFERENCE[1]}', "
            f"'{RECORD_TYPE_PREFERENCE[2]}')")

        query2 = self._sf.query(result2_query)["records"]
        for result in query2:
            email = result[CONTACT_EMAIL_NAME]
            record_id = result[RECORD_TYPE_ID]
            if email is not None:
                email_dict[record_id] = email

        for record_type in RECORD_TYPE_PREFERENCE:
            if record_type in email_dict:
                email = email_dict[record_type]
                if email is not None:
                    return email

        return NO_EMAIL_IDENTIFIER

from docassemble.base.core import DAObject, DAList, DADict
from docassemble.base.util import Individual, Person, Address
from docassemble.base.config import daconfig

import requests
import json

class MSGraphConnectionObject(DAObject):
    """Creates a connection object that can be used to access resources with the Microsoft Graph API with application-level credentials.
    Only a few limited API calls are implemented. Use the Docassemble config options microsoft graph: tenant id, client id, and client secret
    or specify tenant_id, client_id, and client_secret as arguments to the class constructor."""
    def init(self, *pargs, **kwargs):
        super(MSGraphConnectionObject, self).init(*pargs, **kwargs)

        # Default to using Docassemble configuration to retrieve credentials to connect to Microsoft Graph
        # 
        if not hasattr(self,'tenant_id'):
            tenant_id = daconfig.get('microsoft graph',{}).get('tenant id')
        else:
            tenant_id = self.tenant_id
        if not hasattr(self,'client_id'):
            client_id = daconfig.get('microsoft graph',{}).get('client id')
        else:
            client_id = self.client_id
        if not hasattr(self,'client_secret'):
            client_secret = daconfig.get('microsoft graph',{}).get('client secret')
        else:
            client_secret = self.client_secret

        token_url = "https://login.microsoftonline.com/" + tenant_id + "/oauth2/v2.0/token"

        token_data = {
            "client_id" : client_id,
            "client_secret" : client_secret,
            "scope" : "https://graph.microsoft.com/.default",
            "grant_type" :"client_credentials"
        }

        r = requests.post(token_url,data=token_data)
        self.token = r.json()['access_token']

        self.authorization_header = {
           "Authorization": "Bearer " + self.token
        }

    def get_request(self,url, top=100):
        """ Return JSON parsed data for the given URL. Handles using authorization header automatically. Doesn't handle pagination yet but you can specify results limit by changing the value of top.
        Defaults to returning 100 results. Up to 1000 could be returned without pagination."""
        # params = {}
        params = {
                '$top': top
        }
        return (requests.get(url, headers=self.authorization_header, params=params)).json()

    def get_user(self,upn, who = None):
        """Will replace attributes from the given Individual object with user information from Microsoft Graph request. Returns raw JSON results
        if no Individual object is passed in the keyword argument 'who' """
        user_url="https://graph.microsoft.com/v1.0/users/" + upn
        # drq = requests.get(user_url, headers=self.authorization_header)
        # res = drq.json()

        # Unimplemented: get a photo for the user
        # "https://graph.microsoft.com/v1.0/users/qsteenhuis@gbls.org/photo" # ("@odata.mediaContentType": "image/jpeg")
        # "https://graph.microsoft.com/v1.0/users/qsteenhuis@gbls.org/photo/$value"
        # open('headshot.jpg', 'wb').write(r.content)

        res = self.get_request(user_url)

        if who is None:
            return res
        else:
            who.name.first = res.get('givenName')
            who.name.last = res.get('surname')
            who.email = res.get('mail')
            #if isinstance(res.get('businessPhones', None), list):
            who.phone = next(iter(res.get('businessPhones',[])),None)
            who.jobTitle = res.get('jobTitle') 

    def get_contacts(self, upn, default_address = 'home'):
        """ Return a list of contacts from the given user's Universal Principal Name (i.e., username@domain.org for most organizations, or username@org.onmicrosoft.com).
        Does not paginate--will return the first 100 contacts only for now. You can choose whether to default to 'home' or 'business' address and phone."""
        contacts_url = "https://graph.microsoft.com/v1.0/users/" + upn + "/contacts"

        res = self.get_request(contacts_url)

        people = DAList(object_type = Individual, auto_gather=False, gathered=True)

        for p_res in res.get('value',[]):

            person = people.appendObject()
            person.name.first = p_res.get('givenName','')
            person.name.last = p_res.get('surname','')
            if p_res.get('middleName'):
                person.name.middle = p_res.get('middleName')
            person.jobTitle = p_res.get('jobTitle')
            person.title = p_res.get('title')

            person.business_phones = p_res.get('businessPhones',[])
            person.home_phones = p_res.get('homePhones',[])
            person.mobile_number = p_res.get('mobilePhone')

            person.initializeAttribute('home_address',Address)
            person.home_address.address = p_res.get('homeAddress',{}).get('street')
            person.home_address.city = p_res.get('homeAddress',{}).get('city')
            person.home_address.state = p_res.get('homeAddress',{}).get('state')
            person.home_address.zip = p_res.get('homeAddress',{}).get('postalCode')
            #if not p_res.get('homeAddress',{}).get('countryOrRegion') is None:
            #    person.home_address.country = p_res.get('homeAddress',{}).get('countryOrRegion')
            

            person.initializeAttribute('business_address',Address)
            person.business_address.address = p_res.get('businessAddress',{}).get('street')
            person.business_address.city = p_res.get('businessAddress',{}).get('city')
            person.business_address.state = p_res.get('businessAddress',{}).get('state')
            person.business_address.zip = p_res.get('businessAddress',{}).get('postalCode')
            #if not p_res.get('businessAddress',{}).get('countryOrRegion') is None:
            #    person.business_address.country = p_res.get('businessAddress',{}).get('countryOrRegion')

            person.emails = p_res.get('emailAddresses',[])

            if p_res.get('emailAddresses'):
                person.email = next(iter(person.emails),[]).get('address',None) # take the first email

            # Try to respect the address kind the user wants, but if not present, use the address we have (which might be null)
            # Information is often put in the business phone/address fields if the contact only has one address/phone
            if default_address == 'home':
                if p_res.get('homeAddress'):
                    person.address = person.home_address
                else:
                    person.address = person.business_address
            else:
                if p_res.get('businessAddress'):
                    person.address = person.business_address
                else:
                    person.address = person.home_address

            if default_address == 'home':
                if p_res.get('homePhones'):
                    person.phone_number = next(iter(person.home_phones),'') # just take the first home phone in the list
                else:
                    person.phone_number = next(iter(person.business_phones),'')
            else:
                if p_res.get('businessPhones'):
                    person.phone_number = next(iter(person.business_phones),'') # just take the first business phone
                else:
                    person.phone_number = next(iter(person.home_phones),'')                                    

        return people

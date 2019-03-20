# Microsoft Graph API (Office 365) for Docassemble

Provides Graph API access to Docassemble, returning Docassemble objects. E.g.,
Individual, Address, etc.

Works with application level permissions, not user permissions.

## Administrator Setup

Create an application in Azure Portal following instructions [here](https://docs.microsoft.com/en-us/graph/auth-v2-service?view=graph-rest-1.0)
and provide it with a set of credentials.

Add a new section to your Docassemble configuration that looks like this, replacing
with the details from your new application:

```
microsoft graph:
  tenant id: xxxxxxx
  client id: xxxxxxxx
  client secret: xxxxxxxxx
```

## Implemented APIs
* Get user information (with user principal name, typically email address)
* Get user contacts (with upn)

## Usage
See example interview, `msgraph_example.yml`
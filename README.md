# microsoft-python
Microsoft graph API wrapper for Microsoft Graph written in Python.

## Before start
To use Microsoft Graph to read and write resources on behalf of a user, your app must get an access token from
the Microsoft identity platform and attach the token to requests that it sends to Microsoft Graph. The exact
authentication flow that you will use to get access tokens will depend on the kind of app you are developing and
whether you want to use OpenID Connect to sign the user in to your app. One common flow used by native and mobile
apps and also by some Web apps is the OAuth 2.0 authorization code grant flow.

See https://docs.microsoft.com/en-us/graph/auth-v2-user

## Breaking changes if you're upgrading prior 1.0.0
- Added API structure to library for e.g. `client.get_me()` => `client.users.get_me()`.
- Renamed several methods to match API documentation for e.g. `client.get_me_events()` => `client.calendar.list_events()`.
- Result from calling a method is not longer a dictionary but a Response object. To access the dict response as before then call `.data` property for e.g `r = client.users.get_me()` then `r.data`.
- Previous API calls made through beta endpoints are now pointing to v1.0 by default. This can be changed to beta if needed with the parameter `api_version` in the client instantiation.
## New in 1.0.0
- You can access to [Requests library's Response Object](https://docs.python-requests.org/en/latest/user/advanced/#request-and-response-objects) for e.g. `r = client.users.get_me()` then `r.original` or the response handled by the library `r.data`.
- New Response properties `r.status_code` and `r.throttling`.
- You can pass [Requests library's Event Hooks](https://docs.python-requests.org/en/latest/user/advanced/#event-hooks) with the parameter `requests_hooks` in the client instantiation. If you are using Django and want to log in database every request made through this library, see [django-requests-logger](https://github.com/GearPlug/django-requests-logger).
- Better method docstrings and type hinting.
- Better library structure.
## Installing
```
pip install microsoftgraph-python
```
## Usage
### Instantiation
```
from microsoftgraph.client import Client
client = Client('CLIENT_ID', 'CLIENT_SECRET', account_type='common') # by default common, thus account_type is optional parameter.
```

### OAuth 2.0
#### Get authorization url
```
url = client.authorization_url(redirect_uri, scope, state=None)
```

#### Exchange the code for an access token
```
response = client.exchange_code(redirect_uri, code)
```

#### Refresh token
```
response = client.refresh_token(redirect_uri, refresh_token)
```

#### Set token
```
client.set_token(token)
```

### Users
#### Get me
```
response = client.users.get_me()
```

### Mail
#### Get message
```
response = client.mail.get_message(message_id)
```

#### Send mail
```
response = client.mail.send_mail(subject, content, to_recipients)
```

### Notes
#### List notebooks
```
response = client.notes.list_notebooks()
```

#### Get notebook
```
response = client.notes.get_notebook(notebook_id)
```

#### Get notebook sections
```
response = client.notes.list_sections(notebook_id)
```

#### List pages
```
response = client.notes.list_pages()
```

#### Create page
```
response = client.notes.create_page(section_id, files)
```

### Calendar
#### Get events
```
response = client.calendar.list_events()
```

#### Create calendar event
```
response = client.calendar.create_event(subject, content, start_datetime,start_timezone, end_datetime, end_timezone, location, calendar, content_type)
```

#### Get calendars
```
response = client.calendar.list_calendars()
```

#### Create calendar
```
response = client.calendar.create_calendar(name)
```

### Contacts
#### Get a contact
```
response = client.contacts.get_contact(contact_id)
```

#### Get contacts
```
response = client.contacts.list_contacts()
```

#### Create contact
```
response = client.contacts.create_contact()
```

#### Create contact in specific folder
```
response = client.contacts.create_contact_in_folder(folder_id)
```

#### Get contact folders
```
response = client.contacts.list_contact_folders()
```

#### Create contact folders
```
response = client.contacts.create_contact_folder()
```

### Files
#### Get root items
```
response = client.files.drive_root_items()
```

#### Get root children items
```
response = client.files.drive_root_children_items()
```

#### Get specific folder items
```
response = client.files.drive_specific_folder(folder_id)
```

#### Get item
```
response = client.files.drive_get_item(item_id)
```

#### Download the contents of a specific item
```
response = client.files.drive_download_contents(item_id)
```

### Workbooks
#### Create session for specific item
```
response = client.workbooks.create_session(workbook_id)
```

#### Refresh session for specific item
```
response = client.workbooks.refresh_session(workbook_id)
```

#### Close session for specific item
```
response = client.workbooks.close_session(workbook_id)
```

#### Get worksheets
```
response = client.workbooks.list_worksheets(workbook_id)
```

#### Get specific worksheet
```
response = client.workbooks.get_worksheet(workbook_id, worksheet_id)
```

#### Add worksheets
```
response = client.workbooks.add_worksheet(workbook_id)
```

#### Update worksheet
```
response = client.workbooks.update_worksheet(workbook_id, worksheet_id)
```

#### Get charts
```
response = client.workbooks.list_charts(workbook_id, worksheet_id)
```

#### Add chart
```
response = client.workbooks.add_chart(workbook_id, worksheet_id)
```

#### Get tables
```
response = client.workbooks.list_tables(workbook_id)
```

#### Add table
```
response = client.workbooks.add_table(workbook_id)
```

#### Add column to table
```
response = client.workbooks.create_column(workbook_id, worksheet_id, table_id)
```

#### Add row to table
```
response = client.workbooks.create_row(workbook_id, worksheet_id, table_id)
```

#### Get table rows
```
response = client.workbooks.list_rows(workbook_id, table_id)
```

#### Get range
```
response = client.workbooks.get_range(workbook_id, worksheet_id)
```

#### Update range
```
response = client.workbooks.update_range(workbook_id, worksheet_id)
```

### Webhooks
#### Create subscription
```
response = client.webhooks.create_subscription(change_type, notification_url, resource, expiration_datetime, client_state=None)
```

#### Renew subscription
```
response = client.webhooks.renew_subscription(subscription_id, expiration_datetime)
```

#### Delete subscription
```
response = client.webhooks.delete_subscription(subscription_id)
```


## Requirements
- requests

## Tests
```
test/test.py
```

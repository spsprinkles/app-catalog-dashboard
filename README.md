[![NodeJS with Webpack](https://github.com/spsprinkles/app-catalog-dashboard/actions/workflows/webpack.yml/badge.svg)](https://github.com/spsprinkles/app-catalog-dashboard/actions/workflows/webpack.yml)

# App Catalog Management Tool

Helps manage solutions that are deployed to site collection and/or tenant app catalogs.

![Demo](https://github.com/spsprinkles/app-catalog-dashboard/raw/main/solution.png)

## Configuration

The [configuration file](https://github.com/spsprinkles/app-catalog-dashboard/raw/main/assets/config.json) contains the settings for the application.

### URLs

| Name | Description |
| --- | --- |
| appCatalogUrl | The url to the test site collection to be used for testing the application. This site must be have an app catalog enabled on it. |
| dashboardUrl | The url to the dashboard page. This will be used by the `PageUrl` notification content replacement. |
| helpPageUrl | The url to the help page button located in the navigation bar. If not provided, the help button will not be rendered. |
| templatesLibraryUrl | The url to the folder containing the templates. If not provided, the templates dropdown will not be rendered. |
| tenantAppCatalogUrl | The url to the tenant app catalog. |

### Other

| Name | Description |
| --- | --- |
| appCatalogRequests | An array of user types to send notification requests to for app catalog creation. |
| appDetails | The app details information to display. |
| appFlows | The flow ids to trigger based on actions of the app. |
| appNotifications | The ability to send notifications based on app actions. |
| flowEndpoint | The flow authorization endpoint to use when triggering a flow. This is will default to commercial unless specified. Possible values: Flow, FlowChina, FlowDoD, FlowGov, FlowHigh, FlowUSNat, FlowUSSec or the url to the flow authorization endpoint. |
| errorEmails | An array of email addresses to send to when a user has an error when using the application. This will help with troubleshooting issues in production. |
| helpdeskGroupName | The site group name to use to send notifications to the `HelpdeskGroup` users. |
| paging | The default # of apps to display on the dashboard. Valid values are `10, 25, 50, 100`. |
| userAgreement | The user agreement displayed to the developers adding packages. |

### App Details

The `app metadata` to display on the app details page. If the columns are not defined, then the default fields will be displayed.

| Name | Description |
| --- | --- |
| left | An array of internal field names to display in the left column. |
| right | An array of internal field names to display in the right column. |

### App Flows

The `app actions` outside of the workflow that can be triggered.

| Name | Description |
| --- | --- |
| deployToSiteCollection | Triggers a flow when the app is deployed to a site collection. |
| deployToTeams | Triggers a flow when the app is deployed to teams. |
| deployToTenant | Triggers a flow when the app is deployed to the tenant. |
| newApp | Triggers a flow when a new app is initially added. |
| upgradeApp | Triggers a flow when the app is upgraded. |

### App Notifications

During the approval process, the app will go through different actions. This configuration will allow notifications to be sent during these events.

| Name | Description |
| --- | --- |
| deploy | An array of notifications to be sent when an app is deployed to the site collection or tenant app catalog. |
| test | An array of notifications to be sent when the app is upgraded in the test site. |
| upgrade | An array of notifications to be sent when a new app is uploaded with a new version. |

### Status

The `status` configuration value is an object consisting of the following properties.

| Name | Description |
| --- | --- |
| actions | An array of buttons to be displayed on the app dashboard. |
| alert | The alert content that is displayed at the top of the app dashboard for the status. |
| approval | An array of user type values that can approve the item. |
| checklists | Displayed on the approval/submission form. |
| createTestSite | If set to true, will create the test site for the application. |
| flowId | The flow id to trigger for the list item, on approval/submission of the item. |
| notification | The email information to notify users on submission/approval of the item. |
| requiresTechReview | If set to true, will validate the technical review form as part of the approval process. |
| requiresTestCases | If set to true, will validate the test cases form as part of the approval process. |

#### Actions

The available values for the `actions` status property.

| Name | Description |
| --- | --- |
| ApproveReject | Displays the approve and reject button if the user is listed in the `approval` property. |
| AuditLog | Displays the audit log history for an app. |
| DeploySiteCatalog | Displays the deploy to site collection button, allowing the user to deploy the application to a site of their choice. |
| Delete | Displays the delete button to remove the app and test site. |
| DeleteTestSite | Displays the delete test site button if the site exists. |
| DeployTenantCatalog | Displays the option to deploy to the tenant app catalog. |
| Edit | Displays the edit form to modify the app metadata. |
| EditTechReview | Displays the edit form to modify the technical review assessment. |
| EditTestCases | Displays the edit form to modify the test cases assessment. |
| GetHelp | Displays a `help` button to send a notification to the app developers and app sponsor. |
| Notification | Displays a button to notify user types. |
| Submit | Displays the submission form to request approval. |
| TestSite | Displays a `create` or `view` button to the test site collection. |
| Upgrade | Displays an upgrade button if site collection url(s) exist. |
| View | Displays the app metadata view form. |
| ViewTechReview | Displays the last technical review for the application. |
| ViewTestCases | Displays the last test cases for the application. |

#### Alert

The alert property will have the following options.

| Name | Description |
| --- | --- |
| header | The alert header. |
| content | The alert content. |

#### Notification

The notification configuration value is an array of objects with the following properties.

| Name | Description |
| --- | --- |
| approval | If set to true, will send the notification when the approval button is clicked. |
| submission | If set to true, will send the notification when the submit button is clicked. |
| to | An array of user type values to notify. |
| cc | An array of user type values to notify. |
| subject | The email subject. |
| content | The email body. |

##### Notification Content

The notification email content can be dynamically set by using the `[Internal Field Name]` template as a placeholder within it. Complex field types, like a user field, can be used to access their properties. For example, to set the sponsor's name the value `[AppSponsor.Title]` would be used as the placeholder. Other wildcard values that can be used are listed below.

| Name | Description |
| --- | --- |
| DeploySiteUrl | The html link to the site the app was deployed to, for the `appNotifications` deploy configuration. |
| DeploySiteUrlText | The text link to the site the app was deployed to, for the `appNotifications` deploy configuration. |
| Internal Field Name | The item's metadata. |
| PageUrl | The html link to display the item. |
| PageUrlText | The text link to display the item. |
| TestSiteUrl | The html link to the app's test site. |
| TestSiteUrlText | The text link to the app's test site. |

#### User Types

The available values for the user type property values.

| Name | Description |
| --- | --- |
| ApproversGroup | A reference to the approver's security group. |
| CurrentUser | A reference to the current user. |
| Developers | The developers of the application. _Linked to the AppOwners metadata field._ |
| DevelopersGroup | A reference to the developer's security group. |
| FinalApproversGroup | A reference to the final approver's security group. |
| HelpdeskGroup | A reference to the configuration's helpdesk group. _This will require the helpdeskGroupName configuration setting to be set._ |
| Sponsor | The sponsor of the application. _Linked to the AppSponsor metadata field._ |
| SponsorsGroup | A reference to the sponsor's security group. |

### Validation

| Name | Description |
| --- | --- |
| techReview | The valid responses for the technical review form. |
| testCases | The valid responses for the test cases form. |

Each of the validation properties will be an object containing the form fields and valid responses.

{
    "InternalFieldName": ["Valid Responses"]
}
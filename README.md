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
| userAgreement | The user agreement displayed to the developers adding packages. |

### Status

The `status` configuration value is an object consisting of the following properties.

| Name | Description |
| --- | --- |
| requiresTechReview | If set to true, will validate the technical review form as part of the approval process. |
| requiresTestCases | If set to true, will validate the test cases form as part of the approval process. |
| actions | An array of buttons to be displayed on the app dashboard. |
| approval | An array of user type values that can approve the item. |
| checklists | Displayed on the approval/submission form. |
| createTestSite | If set to true, will create the test site for the application. |
| notification | The email information to notify users on submission/approval of the item. |

#### Actions

The available values for the `actions` status property.

| Name | Description |
| --- | --- |
| AppData | Displays the edit form to modify the metadata related to the app catalog. |
| ApproveReject | Displays the approve and reject button if the user is listed in the `approval` property. |
| DeploySiteCatalog | Displays the deploy to site collection button, allowing the user to deploy the application to a site of their choice. |
| DeleteTestSite | Displays the delete test site button if the site exists. |
| DeployTenantCatalog | Displays the option to deploy to the tenant app catalog. |
| Edit | Displays the edit form to modify the app metadata. |
| EditTechReview | Displays the edit form to modify the technical review assessment. |
| EditTestCases | Displays the edit form to modify the test cases assessment. |
| Submit | Displays the submission form to request approval. |
| TestSite | Displays a `create` or `view` button to the test site collection. |
| View | Displays the app metadata view form. |
| ViewTechReview | Displays the last technical review for the application. |
| ViewTestCases | Displays the last test cases for the application. |

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

The notification email content can be dynamically set by using the `[[Internal Field Name]]` template as a placeholder within it. Complex field types, like a user field, can be used to access their properties. For example, to set the sponsor's name the value `[[AppSponsor.Title]]` would be used as the placeholder. Other wildcard values that can be used are listed below.

| Name | Description |
| --- | --- |
| Internal Field Name | The item's metadata. |
| PageUrl | The html link to display the item. |

#### User Types

The available values for the user type property values.

| Name | Description |
| --- | --- |
| ApproversGroup | A reference to the approver's security group. |
| Developers | The developers of the application. _Linked to the AppOwners metadata field._ |
| DevelopersGroup | A reference to the developer's security group. |
| Sponsor | The sponsor of the application. _Linked to the AppSponsor metadata field._ |

### Validation

| Name | Description |
| --- | --- |
| techReview | The valid responses for the technical review form. |
| testCases | The valid responses for the test cases form. |

Each of the validation properties will be an object containing the form fields and valid responses.

{
    "InternalFieldName": ["Valid Responses"]
}
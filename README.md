[![NodeJS with Webpack](https://github.com/spsprinkles/app-catalog-dashboard/actions/workflows/webpack.yml/badge.svg)](https://github.com/spsprinkles/app-catalog-dashboard/actions/workflows/webpack.yml)

# App Catalog Management Tool

Helps manage solutions that are deployed to site collection and/or tenant app catalogs.

![Demo](https://github.com/spsprinkles/app-catalog-dashboard/raw/main/solution.png)

## Audit Log

The tool will track a history of actions taken on an app. This audit log is visible to admins and owners. Below is a list of states that are tracked.

| Name | Description |
| --- | --- |
| App Added | Logged when an app is added to the tool. |
| App Approved | Logged when an app is approved. |
| App Catalog Request | Logged when an app catalog is being requested for a site collection. |
| App Deployed | Logged when an app is deployed to a site collection app catalog. |
| App Metadata Updated | Logged when the metadata is updated past the technical review state. |
| App Rejected | Logged when an app is rejected. |
| App Retracted | Logged when an app is retracted. |
| App Resubmitted | Logged when an app is resubmitted from a rejected state. |
| App Submitted | Logged when an app is submitted for approval. |
| App Tenant Deployed | Logged when an app is deployed to the tenant app catalog. |
| App Updated | Logged when an app is updated in the tool. |
| App Upgraded | Logged when an app is upgraded in the tool. |
| Create Test Site | Logged when the test site is created for the app. |
| Delete App | Logged when an app is deleted from the tool |
| Delete Test Site | Logged when the test site is deleted for the app. |

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
| cdnImage | The absolute url to the library used to store the app icons. |
| cdnTest | The absolute url to the library used to store the client side assets for testing. |
| cdnProd | The absolute url to the library used to store the client side assets for production. |
| cloudEnv | The default cloud environment the app is in. This will default to commerical if no value is specified. |
| dateFormat | The date format to use. _(Examples: YYYY-MM-DD or YYYY-MM-DD HH:mm:ss)_ |
| flowEndpoint | The flow authorization endpoint to use when triggering a flow. This is will default to commercial unless specified. Possible values: Flow, FlowChina, FlowDoD, FlowGov, FlowHigh, FlowUSNat, FlowUSSec or the url to the flow authorization endpoint. |
| paging | The default # of apps to display on the dashboard. Valid values are `10, 25, 50, 100`. |
| supportUrl | The default url to use for support. |
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
| CDN | Displays the action button for uploading the CDN files. |
| DeploySiteCatalog | Displays the deploy to site collection button, allowing the user to deploy the application to a site of their choice. |
| Delete | Displays the delete button to remove the app and test site. |
| DeleteTestSite | Displays the delete test site button if the site exists. |
| DeployTenantCatalog | Displays the option to deploy to the tenant app catalog. |
| Edit | Displays the edit form to modify the app metadata. |
| EditTechReview | Displays the edit form to modify the technical review assessment. |
| EditTestCases | Displays the edit form to modify the test cases assessment. |
| Metadata| Displays a button to update the metadata of an approved app. This will revert the item to "Pending Approval". |
| Submit | Displays the submission form to request approval. |
| TestSite | Displays a `create` or `view` button to the test site collection. |
| Upgrade | Displays an upgrade button if site collection url(s) exist. |
| UpgradeInfo | Displays a dialog for entering the app's upgrade information. |
| UpgradeTenant | Displays an upgrade button if tenant version needs to be upgraded. |
| View | Displays the app metadata view form. |
| ViewSourceControl | Displays the app's source control in a new tab. |
| ViewTechReview | Displays the last technical review for the application. |
| ViewTestCases | Displays the last test cases for the application. |

#### Alert

The alert property will have the following options.

| Name | Description |
| --- | --- |
| header | The alert header. |
| content | The alert content. |

#### User Types

The available values for the user type property values.

| Name | Description |
| --- | --- |
| ApproversGroup | A reference to the approver's security group. |
| CurrentUser | A reference to the current user. |
| Developers | The developers of the application. _Linked to the AppOwners metadata field._ |
| DevelopersGroup | A reference to the developer's security group. |
| FinalApproversGroup | A reference to the final approver's security group. |
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
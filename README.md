# App Catalog Management Tool

Helps manage solutions that are deployed to site collection and/or tenant app catalogs.

![Demo](https://github.com/spsprinkles/app-catalog-dashboard/raw/main/solution.png)

## Configuration

The [configuration file](https://github.com/spsprinkles/app-catalog-dashboard/raw/main/assets/config.json) contains the settings for the application.

### URLs

| Name | Description |
| --- | --- |
| appCatalogUrl | The url to the test site collection to be used for testing the application. This site must be have an app catalog enabled on it. |
| helpPageUrl | The url to the help page button located in the navigation bar. If not provided, the help button will not be rendered. |
| templatesLibraryUrl | The url to the folder containing the templates. If not provided, the templates dropdown will not be rendered. |
| tenantAppCatalogUrl | The url to the tenant app catalog. |

### Status

| Name | Description |
| --- | --- |
| requiresTechReview | If set to true, will validate the technical review form as part of the approval process. |
| requiresTestCases | If set to true, will validate the test cases form as part of the approval process. |
| actions | An array of buttons to be displayed on the app dashboard. |
| approval | The type of user that can approve the item. |
| checklists | Displayed on the approval/submission form. |
| notification | The email information to notify users on submission/approval of the item. |

The notification property will be an array of objects with the following properties.

| Name | Description |
| --- | --- |
| approval | If set to true, will send the notification when the approval button is clicked. |
| submission | If set to true, will send the notification when the submit button is clicked. |
| to | An array of user types to notify. |
| cc | An array of user types to notify. |
| subject | The email subject. |
| content | The email body. Utilize [[Internal Field Name]] to set dynamic values from the associated item. If it's a user field, add the additional property to display: _[[AppSponser.Title]]_ |

### Validation

| Name | Description |
| --- | --- |
| techReview | The valid responses for the technical review form. |
| testCases | The valid responses for the test cases form. |

Each of the validation properties will be an object containing the form fields and valid responses.

{
    "InternalFieldName": ["Valid Responses"]
}
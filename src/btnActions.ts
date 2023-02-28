import { Components, ContextInfo, Web } from "gd-sprest-bs";
import { appIndicator } from "gd-sprest-bs/build/icons/svgs/appIndicator";
import { arrowClockwise } from "gd-sprest-bs/build/icons/svgs/arrowClockwise";
import { boxArrowDownRight } from "gd-sprest-bs/build/icons/svgs/boxArrowDownRight";
import { boxArrowInLeft } from "gd-sprest-bs/build/icons/svgs/boxArrowInLeft";
import { boxArrowRight } from "gd-sprest-bs/build/icons/svgs/boxArrowRight";
import { cardChecklist } from "gd-sprest-bs/build/icons/svgs/cardChecklist";
import { cardText } from "gd-sprest-bs/build/icons/svgs/cardText";
import { clockHistory } from "gd-sprest-bs/build/icons/svgs/clockHistory";
import { envelope } from "gd-sprest-bs/build/icons/svgs/envelope";
import { git } from "gd-sprest-bs/build/icons/svgs/git";
import { handThumbsDown } from "gd-sprest-bs/build/icons/svgs/handThumbsDown";
import { handThumbsUp } from "gd-sprest-bs/build/icons/svgs/handThumbsUp";
import { inputCursorText } from "gd-sprest-bs/build/icons/svgs/inputCursorText";
import { microsoftTeams } from "gd-sprest-bs/build/icons/svgs/microsoftTeams";
import { nodePlus } from "gd-sprest-bs/build/icons/svgs/nodePlus";
import { questionSquare } from "gd-sprest-bs/build/icons/svgs/questionSquare";
import { send } from "gd-sprest-bs/build/icons/svgs/send";
import { trash3 } from "gd-sprest-bs/build/icons/svgs/trash3";
import { window_ } from "gd-sprest-bs/build/icons/svgs/window_";
import { windowStack } from "gd-sprest-bs/build/icons/svgs/windowStack";
import * as moment from "moment";
import { AppConfig, IStatus, UserTypes } from "./appCfg";
import { AppForms } from "./appForms";
import { AppSecurity } from "./appSecurity";
import * as Common from "./common";
import { DataSource, IAppItem } from "./ds";
import Strings from "./strings";

/**
 * Actions
 * Renders the actions available, based on the status.
 */
export class ButtonActions {
  private _el: HTMLElement = null;
  private _forms: AppForms = null;
  private _item: IAppItem = null;
  private _onUpdate: () => void = null;

  // Constructor
  constructor(el: HTMLElement, item: IAppItem, onUpdate: () => void) {
    // Initialize the component
    this._el = el;
    this._forms = new AppForms();
    this._item = item;
    this._onUpdate = onUpdate;

    // Render the actions
    this.render();
  }

  // Status
  private get Status(): IStatus {
    return AppConfig.Status[this._item.AppStatus] || ({} as any);
  }

  // Determines if the user is an approver
  private isApprover() {
    // See if there is an approver setting for this status
    let approvers = this.Status.approval || [];
    if (approvers.length == 0) {
      return false;
    }

    // Parse the approver
    let isApprover = false;
    for (let i = 0; i < approvers.length; i++) {
      // See which approver type it is
      switch (approvers[i]) {
        // Approver's Group
        case UserTypes.ApproversGroup:
          // See if this user is an approver
          isApprover = isApprover || AppSecurity.IsApprover;
          break;

        // Developers of the application
        case UserTypes.Developers:
          // See if this user is an owner of the app'
          let developerIds = this._item.AppDevelopersId
            ? this._item.AppDevelopersId.results
            : [];
          for (let j = 0; j < developerIds.length; i++) {
            // See if the user is a developer
            if (developerIds[j] == ContextInfo.userId) {
              // Set the flag
              isApprover = true;
              break;
            }
          }
          break;

        // Developer's Group
        case UserTypes.DevelopersGroup:
          // See if this user is a developer
          isApprover = isApprover || AppSecurity.IsDeveloper;
          break;

        // Final Approver's Group
        case UserTypes.FinalApproversGroup:
          // See if this user is a final approver
          isApprover = isApprover || AppSecurity.IsFinalApprover;
          break;

        // Sponsor
        case UserTypes.Sponsor:
          // See if this user is a sponsor
          isApprover =
            isApprover || this._item.AppSponsorId == ContextInfo.userId;
          break;

        // Sponsor's Group
        case UserTypes.SponsorsGroup:
          // See if this user is a sponsor
          isApprover = isApprover || AppSecurity.IsSponsor;
          break;
      }
    }

    // Return the flag
    return isApprover;
  }

  // Determines if the user is an app developer
  private isDeveloper() {
    let isDeveloper = false;

    // See if there is an approver setting for this status
    let developers =
      (this._item.AppDevelopersId
        ? this._item.AppDevelopersId.results
        : null) || [];

    // Parse the developers
    for (let i = 0; i < developers.length; i++) {
      // See if the current user is the app developer
      if (developers[i] == ContextInfo.userId) {
        // Set the flag and break from the loop
        isDeveloper = true;
        break;
      }
    }

    // Return the flag
    return isDeveloper;
  }

  // Renders the actions
  render() {
    // Ensure actions exist
    let btnActions = this.Status.actions;
    if (btnActions == null || btnActions.length == 0) {
      return;
    }

    // Determine if the user can edit items
    let canEdit = Common.canEdit(this._item);

    // Render the tooltip group
    let tooltips = Components.TooltipGroup({
      el: this._el,
      className: "w-50",
      isVertical: true,
      tooltipPlacement: Components.TooltipPlacements.Right,
    });

    // Parse the button actions
    for (let i = 0; i < btnActions.length; i++) {
      // Render the action button
      switch (btnActions[i]) {
        // Approve/Reject
        case "ApproveReject":
          // See if the app is rejected
          if (this._item.AppIsRejected) {
            // See if this is an approver or developer
            if (this.isApprover() || this.isDeveloper()) {
              // Render the resubmit button
              tooltips.add({
                content: "Resubmits the app for approval.",
                btnProps: {
                  text: "Resubmit",
                  iconClassName: "me-1",
                  iconSize: 20,
                  iconType: arrowClockwise,
                  isSmall: true,
                  type: Components.ButtonTypes.OutlinePrimary,
                  onClick: () => {
                    // Display the approval form
                    this._forms.submit(this._item, () => {
                      // Call the update event
                      this._onUpdate();
                    });
                  },
                },
              });
            }
          }
          // Else, ensure this user can approve this item
          else if (this.isApprover()) {
            // Render the approval button
            tooltips.add({
              content: "Approves the application.",
              btnProps: {
                text: "Approve",
                iconClassName: "me-1",
                iconSize: 20,
                iconType: handThumbsUp,
                isSmall: true,
                type: Components.ButtonTypes.OutlineSuccess,
                onClick: () => {
                  // Display the approval form
                  this._forms.approve(this._item, () => {
                    // Call the update event
                    this._onUpdate();
                  });
                },
              },
            });

            // Render the reject button
            tooltips.add({
              content: "Sends the request back to the developer(s).",
              btnProps: {
                text: "Reject",
                iconClassName: "me-1",
                iconSize: 20,
                iconType: handThumbsDown,
                isSmall: true,
                type: Components.ButtonTypes.OutlineDanger,
                onClick: () => {
                  // Display the reject form
                  this._forms.reject(this._item, () => {
                    // Call the update event
                    this._onUpdate();
                  });
                },
              },
            });
          }
          break;

        // Audit Log
        case "AuditLog":
          // Ensure the user is an admin or owner
          if (AppSecurity.IsAdmin || AppSecurity.isOwner) {
            // Render the resubmit button
            tooltips.add({
              content: "Displays the audit log information.",
              btnProps: {
                text: "Audit History",
                iconClassName: "me-1",
                iconSize: 20,
                iconType: clockHistory,
                isSmall: true,
                type: Components.ButtonTypes.OutlineSecondary,
                onClick: () => {
                  // Display the audit log information
                  DataSource.AuditLog.viewLog({
                    id: this._item.AppProductID,
                    listName: Strings.Lists.Apps,
                    onTableCellRendering: (el, col, item) => {
                      // See if this is the date/time column
                      if (col.name == "Created") {
                        // Format the value
                        el.innerHTML = moment(item.Created).format(AppConfig.Configuration.dateFormat);
                      }
                    }
                  });
                },
              },
            });
          }
          break;

        // Delete
        case "Delete":
          // Ensure this is an approver
          if (AppSecurity.IsApprover) {
            // Render the button
            tooltips.add({
              content: "Deletes the app.",
              btnProps: {
                text: "Delete App/Solution",
                iconClassName: "me-1",
                iconSize: 20,
                iconType: trash3,
                isSmall: true,
                type: Components.ButtonTypes.OutlineDanger,
                onClick: () => {
                  // Display the delete form
                  this._forms.delete(this._item, () => {
                    // Redirect to the dashboard
                    window.open(AppConfig.Configuration.dashboardUrl, "_self");
                  });
                },
              },
            });
          }
          break;

        // Delete Test Site
        case "DeleteTestSite":
          // See if this is an approver
          if (AppSecurity.IsSiteAppCatalogOwner) {
            // See if a test site exists
            DataSource.loadTestSite(this._item).then(
              // Test site exists
              () => {
                // Render the button
                tooltips.add({
                  content: "Deletes the test site.",
                  btnProps: {
                    text: "Delete Test Site",
                    iconClassName: "me-1",
                    iconSize: 20,
                    iconType: trash3,
                    isSmall: true,
                    type: Components.ButtonTypes.OutlineDanger,
                    onClick: () => {
                      // Display the delete site form
                      this._forms.deleteSite(this._item, () => {
                        // Redirect to the dashboard
                        window.open(
                          AppConfig.Configuration.dashboardUrl,
                          "_self"
                        );
                      });
                    },
                  },
                });
              },

              // Test site doesn't exist
              () => { }
            );
          }
          break;

        // Deploy Site Catalog
        case "DeploySiteCatalog":
          // Render the retract button
          tooltips.add({
            content:
              "Deploys the application to a site collection app catalog.",
            btnProps: {
              text: "Deploy to Site Collection",
              iconClassName: "me-1",
              iconSize: 20,
              iconType: boxArrowDownRight,
              isSmall: true,
              type: Components.ButtonTypes.OutlineSuccess,
              onClick: () => {
                // Display the delete site form
                this._forms.deployToSite(this._item, () => {
                  // Call the update event
                  this._onUpdate();
                });
              },
            },
          });
          break;

        // Deploy Tenant Catalog
        case "DeployTenantCatalog":
          // Ensure this is a tenant app catalog owner
          if (AppSecurity.IsTenantAppCatalogOwner) {
            let showDeploy = false;
            let showRemove = false;
            let showRetract = false;

            // See if the app is deployed
            if (DataSource.AppCatalogTenantItem) {
              // See if it's deployed
              if (DataSource.AppCatalogTenantItem.Deployed) {
                // Set the flags
                showRemove = true;
                showRetract = true;

                // Load the context of the app catalog
                ContextInfo.getWeb(AppConfig.Configuration.tenantAppCatalogUrl).execute((context) => {
                  let requestDigest = context.GetContextWebInformation.FormDigestValue;
                  let web = Web(AppConfig.Configuration.tenantAppCatalogUrl, { requestDigest });

                  // Ensure this app can be deployed to the tenant
                  web.TenantAppCatalog().solutionContainsTeamsComponent(DataSource.AppCatalogTenantItem.ID).execute((resp: any) => {
                    // See if we can deploy this app to teams
                    if (resp.SolutionContainsTeamsComponent) {
                      // Render the deploy to teams button
                      tooltips.add({
                        content: "Deploys the solution to Teams.",
                        btnProps: {
                          text: "Deploy to Teams",
                          iconClassName: "me-1",
                          iconSize: 20,
                          iconType: microsoftTeams,
                          isDisabled: AppSecurity.IsTenantAppCatalogOwner,
                          isSmall: true,
                          type: Components.ButtonTypes.OutlineSuccess,
                          onClick: () => {
                            // Deploy the app
                            this._forms.deployToTeams(this._item, () => {
                              // Call the update event
                              this._onUpdate();
                            });
                          },
                        },
                      });
                    }
                  });
                });
              } else {
                // Set the flags
                showDeploy = true;
                showRemove = true;
              }
            } else {
              // Set the flag
              showDeploy = true;
            }

            // See if we are showing the deploy button
            if (showDeploy) {
              // Render the deploy button
              tooltips.add({
                content: "Deploys the app to the tenant app catalog.",
                btnProps: {
                  text: "Deploy to Tenant",
                  iconClassName: "me-1",
                  iconSize: 20,
                  iconType: boxArrowRight,
                  isSmall: true,
                  type: Components.ButtonTypes.OutlineSuccess,
                  onClick: () => {
                    // Deploy the app
                    this._forms.deploy(this._item, true, () => {
                      // Call the update event
                      this._onUpdate();
                    });
                  },
                },
              });
            }

            // See if we are showing the remove button
            if (showRemove) {
              // Render the retract button
              tooltips.add({
                content: "Removes the app from the tenant app catalog.",
                btnProps: {
                  text: "Remove from Tenant",
                  iconClassName: "me-1",
                  iconSize: 20,
                  iconType: boxArrowInLeft,
                  isSmall: true,
                  type: Components.ButtonTypes.OutlineDanger,
                  onClick: () => {
                    // Remove the app
                    this._forms.removeFromTenant(this._item, () => {
                      // Call the update event
                      this._onUpdate();
                    });
                  },
                },
              });
            }

            // See if we are showing the retract button
            if (showRetract) {
              // Render the retract button
              tooltips.add({
                content: "Retracts the app from the tenant app catalog.",
                btnProps: {
                  text: "Retract from Tenant",
                  iconClassName: "me-1",
                  iconSize: 20,
                  iconType: boxArrowInLeft,
                  isSmall: true,
                  type: Components.ButtonTypes.OutlineDanger,
                  onClick: () => {
                    // Retract the app
                    this._forms.retractFromTenant(this._item, () => {
                      // Call the update event
                      this._onUpdate();
                    });
                  },
                },
              });
            }
          }
          break;

        // Edit
        case "Edit":
          // Render the button
          tooltips.add({
            content: "Displays the edit form to update the app properties.",
            btnProps: {
              text: "Edit Properties",
              iconClassName: "me-1",
              iconSize: 20,
              iconType: inputCursorText,
              isDisabled: !canEdit,
              isSmall: true,
              type: Components.ButtonTypes.OutlinePrimary,
              onClick: () => {
                // Display the edit form
                this._forms.edit(this._item.Id, () => {
                  // Call the update event
                  this._onUpdate();
                });
              },
            },
          });
          break;

        // Edit Tech Review
        case "EditTechReview":
          // Render the button
          tooltips.add({
            content: "Edits the technical review for this app.",
            btnProps: {
              text: "Edit Technical Review",
              iconClassName: "me-1",
              iconSize: 20,
              iconType: cardChecklist,
              isSmall: true,
              type: Components.ButtonTypes.OutlinePrimary,
              onClick: () => {
                // Display the test cases
                this._forms.editTechReview(this._item, () => {
                  // Call the update event
                  this._onUpdate();
                });
              },
            },
          });
          break;

        // Edit Test Cases
        case "EditTestCases":
          // Render the button
          tooltips.add({
            content: "Edit the test cases for this app.",
            btnProps: {
              text: "Edit Test Cases",
              iconClassName: "me-1",
              iconSize: 20,
              iconType: cardChecklist,
              isDisabled: !canEdit,
              isSmall: true,
              type: Components.ButtonTypes.OutlinePrimary,
              onClick: () => {
                // Display the review form
                this._forms.editTestCases(this._item, () => {
                  // Call the update event
                  this._onUpdate();
                });
              },
            },
          });
          break;

        // Get Help
        case "GetHelp":
          // Render the button
          tooltips.add({
            content: "Sends a notification to the app owners.",
            btnProps: {
              text: "Get Help",
              iconClassName: "me-1",
              iconSize: 20,
              iconType: questionSquare,
              isSmall: true,
              type: Components.ButtonTypes.OutlinePrimary,
              onClick: () => {
                this._forms.sendNotification(
                  this._item,
                  ["Developers", "Sponsor"],
                  "Request for Help",
                  `App Developers,

We are requesting help for the app ${this._item.Title}.

(Enter Details Here)

r/,
${ContextInfo.userDisplayName}`.trim()
                );
              },
            },
          });
          break;

        // Notification
        case "Notification":
          // Render the button
          tooltips.add({
            content: "Sends a notification to user(s).",
            btnProps: {
              text: "Notification",
              iconClassName: "me-1",
              iconSize: 20,
              iconType: envelope,
              isSmall: true,
              type: Components.ButtonTypes.OutlinePrimary,
              onClick: () => {
                // Display the send notification form
                this._forms.sendNotification(this._item);
              },
            },
          });
          break;

        // Submit
        case "Submit":
          // Render the button
          tooltips.add({
            content: "Submits the app for approval/review",
            btnProps: {
              text: "Submit",
              iconClassName: "me-1",
              iconSize: 20,
              iconType: send,
              isDisabled: !canEdit || this.Status.lastStep,
              isSmall: true,
              type: Components.ButtonTypes.OutlinePrimary,
              onClick: () => {
                // Display the submit form
                this._forms.submit(this._item, () => {
                  // Call the update event
                  this._onUpdate();
                });
              },
            },
          });
          break;

        // Test Site
        case "TestSite":
          // See if a test site exists
          DataSource.loadTestSite(this._item).then(
            // Test site exists
            (web) => {
              // Render the view button
              tooltips.add({
                content: "Opens the test site in a new tab.",
                btnProps: {
                  text: "View Test Site",
                  iconClassName: "me-1",
                  iconSize: 20,
                  iconType: windowStack,
                  isSmall: true,
                  type: Components.ButtonTypes.OutlineSecondary,
                  onClick: () => {
                    // Open the test site in a new tab
                    window.open(web.Url, "_blank");
                  },
                },
              });

              // See if the current version is not deployed
              if (
                DataSource.AppCatalogSiteItem &&
                this._item.AppVersion !=
                DataSource.AppCatalogSiteItem.InstalledVersion &&
                this._item.AppVersion !=
                DataSource.AppCatalogSiteItem.AppCatalogVersion
              ) {
                // Render the update button
                tooltips.add({
                  content:
                    "Versions do not match. Click to update the test site.",
                  btnProps: {
                    text: "Update Test Site",
                    iconClassName: "me-1",
                    iconSize: 20,
                    iconType: appIndicator,
                    isDisabled: !AppSecurity.IsSiteAppCatalogOwner,
                    isSmall: true,
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => {
                      // Show the update form
                      this._forms.updateApp(
                        this._item,
                        web.Url,
                        () => {
                          // Call the update event
                          this._onUpdate();
                        }
                      );
                    },
                  },
                });
              }
              // Else, see if the app doesn't exist
              else if (DataSource.AppCatalogSiteItem == null) {
                // Render the update button
                tooltips.add({
                  content:
                    "App is not installed on the test site. Click to update the test site.",
                  btnProps: {
                    text: "Update Test Site",
                    iconClassName: "me-1",
                    iconSize: 20,
                    iconType: appIndicator,
                    isDisabled: !AppSecurity.IsSiteAppCatalogOwner,
                    isSmall: true,
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => {
                      // Show the update form
                      this._forms.updateApp(
                        this._item,
                        web.Url,
                        () => {
                          // Call the update event
                          this._onUpdate();
                        }
                      );
                    },
                  },
                });
              }
            },

            // Test site doesn't exist
            () => {
              // Ensure the app item exists
              if (DataSource.AppCatalogSiteItem) {
                // Render the create button
                tooltips.add({
                  content: "Creates the test site for the app.",
                  btnProps: {
                    text: "Create Test Site",
                    iconClassName: "me-1",
                    iconSize: 20,
                    iconType: nodePlus,
                    isDisabled:
                      !AppSecurity.IsSiteAppCatalogOwner ||
                      !DataSource.AppCatalogSiteItem.IsValidAppPackage,
                    isSmall: true,
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => {
                      // Display the create test site form
                      this._forms.createTestSite(this._item, () => {
                        // Call the update event
                        this._onUpdate();
                      });
                    },
                  },
                });
              }
            }
          );
          break;

        // Upgrade
        case "Upgrade":
          // See if site collection urls exist
          let urls = (this._item.AppSiteDeployments || "").trim();
          if (urls.length > 0) {
            // Render the button
            tooltips.add({
              content:
                "Upgrades the apps in the site collections currently deployed to.",
              btnProps: {
                text: "Upgrade",
                iconClassName: "me-1",
                iconSize: 20,
                iconType: appIndicator,
                isSmall: true,
                type: Components.ButtonTypes.OutlinePrimary,
                onClick: () => {
                  // Display the upgrade form
                  this._forms.upgrade(this._item);
                },
              },
            });
          }
          break;

        // View
        case "View":
          // Render the button
          tooltips.add({
            content: "Displays the the app properties.",
            btnProps: {
              text: "View Properties",
              iconClassName: "me-1",
              iconSize: 20,
              iconType: window_,
              isSmall: true,
              type: Components.ButtonTypes.OutlineSecondary,
              onClick: () => {
                // Display the edit form
                this._forms.display(this._item.Id);
              },
            },
          });
          break;

        // View Source Control
        case "ViewSourceControl":
          // Ensure the source control url exists for this app
          if (this._item.AppSourceControl) {
            // Render the button
            tooltips.add({
              content: "Views the source code for the app.",
              btnProps: {
                text: "View Source Control",
                iconClassName: "me-1",
                iconSize: 20,
                iconType: git,
                isSmall: true,
                type: Components.ButtonTypes.OutlineSecondary,
                onClick: () => {
                  // Display the url in a new tab
                  window.open(this._item.AppSourceControl.Url, "_blank");
                },
              },
            });
          }
          break;

        // View Tech Review
        case "ViewTechReview":
          // Render the button
          tooltips.add({
            content: "Views the technical review for this app.",
            btnProps: {
              text: "View Technical Review",
              iconClassName: "me-1",
              iconSize: 20,
              iconType: cardText,
              isSmall: true,
              type: Components.ButtonTypes.OutlineSecondary,
              onClick: () => {
                // Display the test cases
                this._forms.displayTechReview(this._item);
              },
            },
          });
          break;

        // View Test Cases
        case "ViewTestCases":
          // Render the button
          tooltips.add({
            content: "View the test cases for this app.",
            btnProps: {
              text: "View Test Cases",
              iconClassName: "me-1",
              iconSize: 20,
              iconType: cardText,
              isSmall: true,
              type: Components.ButtonTypes.OutlineSecondary,
              onClick: () => {
                // Display the review form
                this._forms.displayTestCases(this._item);
              },
            },
          });
          break;
      }
    }
    // Fix weird alignment of the first button element
    (
      tooltips.el.querySelector("button:first-child") as HTMLButtonElement
    ).style.marginLeft = "-1px";
  }
}

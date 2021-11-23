(function() {
	function readOnlyField(ctx) {
		return RenderFieldValueDefault(ctx); //text only
	}
	
	SPClientTemplates.TemplateManager.RegisterTemplateOverrides({
		Templates: {
			Fields: {
				"Title": {
					//Options: View & DisplayForm
					EditForm: readOnlyField
				}
			}
		},
		OnPreRender: function (ctx) {
			//BaseViewID: 1 or ControlMode: 4 (a list view, not a form)
			//ControlMode: 1 displayform 2:edit 3: new 4: list view
			/*view: "{7C951752-CEAB-46B7-A06B-084C75D11AF0}"
			viewTitle: "All Documents"
			wpq: "WPQ1"
			*/
			
			//ctx.view: "{F62BFB09-B6AD-42A0-BC78-112EF321A7F3}"
			//div webpartid="f62bfb09-b6ad-42a0-bc78-112ef321a7f3"
			if (ctx.BaseViewID == 1 && !ctx.CurrentUserIsSiteAdmin) {
				var style = document.createElement("style");
				style.innerHTML = "table.ms-listviewtable { display:none; }";
				document.getElementsByTagName("head")[0].append(style);
				"stop" in window ? window.stop() : document.execCommand("Stop");
				location.assign(_spPageContextInfo.webAbsoluteUrl + "/SitePages/AppDashboard.aspx");
			}
			
			if (ctx.BaseViewID == "DisplayForm" && !_spPageContextInfo.isSiteAdmin && ctx.ListSchema.Field[0].InternalName == "FileLeafRef") { //or Created
				//.Owners: "14;#John Doe"  or"8;#Michael Vasiloff;#16;#Jane Doe"
				//.Author: "14;#John Doe,#i:0#.f|membership|john.doe@domain.onmicrosoft.com,#john.doe@domain.onmicrosoft.com,#john.doe@domain.onmicrosoft.com,#John Doe"
				var userFound = false;
				ctx.ListData.Items[0].Owners.split(";#").forEach(function(item) {
				    if (parseInt(item) === _spPageContextInfo.userId)
				    	userFound = true;
				});
				if (ctx.ListData.Items[0].Author.split(";#")[0] == _spPageContextInfo.userId)
				    userFound = true;
				//Add style(s)
				var style = document.createElement("style");
				style.innerHTML = (!userFound ? "#Ribbon\\.ListForm\\.Display\\.Manage\\.EditItem-Large, " : "") + "#Ribbon\\.ListForm\\.Display\\.Manage\\.CheckOut-Large { display:none; }";
				document.getElementsByTagName("head")[0].append(style);
			}
			
			if (ctx.BaseViewID == "EditForm" && !_spPageContextInfo.isSiteAdmin && ctx.ListSchema.Field[0].InternalName == "FileLeafRef") { //or Created
				//Owners is an array in EditForm: [].EntityData.SPUserID
				//Author is same format as Display
				var userFound = false;
				ctx.ListData.Items[0].Owners.forEach(function(item) {
				    if (item.EntityData.SPUserID === _spPageContextInfo.userId)
				    	userFound = true;
				});
				if (ctx.ListData.Items[0].Author.split(";#")[0] == _spPageContextInfo.userId)
				    userFound = true;
				if (!userFound) {
					var style = document.createElement("style");
					style.innerHTML = "#Ribbon\\.DocLibListForm\\.Edit\\.Commit\\.Publish-Large, table.ms-formtable { display:none; }";
					document.getElementsByTagName("head")[0].append(style);
					"stop" in window ? window.stop() : document.execCommand("Stop");
					location.assign(_spPageContextInfo.webAbsoluteUrl + "/SitePages/AppDashboard.aspx");
				}
			}
			
			if (ctx.ControlMode < 4 && ctx.ListSchema.Field[0].InternalName == "FileLeafRef") {
				var style = document.createElement("style");
				//Add styling to table
				style.innerHTML = "table.ms-formtable td.ms-formlabel { padding-right:25px; } " +
					"table.ms-formtable > tbody > tr > td { font-size:1.1em; border-bottom:1px solid rgb(234,234,234); } " +
					"table.ms-formtable input[type=text], table.ms-formtable div.sp-peoplepicker-topLevel, table.ms-formtable select { min-height:25px; } " +
					"table.ms-formtable input[type=checkbox] { width:20px; height:20px; } " +
					//Enlarge fields
					".ms-long { width:550px; } " +
					".sp-peoplepicker-topLevel { width:535px; } " +
					".sp-peoplepicker-autoFillContainer { max-width:450px; } ";
					//Shorten date fields
					//input[id$='_$DateTimeFieldDate'] { width:112px; }
				document.getElementsByTagName("head")[0].append(style);
			}
		},
		OnPostRender: function (ctx) {
			if (ctx.BaseViewID == "EditForm") {
				switch (ctx.ListSchema.Field[0].InternalName) {
					case "AppThumbnailURL":
					case "AppSupportURL":
					case "AppImageURL1":
					case "AppImageURL2":
					case "AppImageURL3":
					case "AppImageURL4":
					case "AppImageURL5":
					case "AppVideoURL":
						if (ctx.ListSchema.Field[0].FieldType == "URL") {
							var elem = $get(ctx.ListSchema.Field[0].InternalName + "_" + ctx.ListSchema.Field[0].Id + "_$UrlFieldDescription");
							if (elem) {
								elem.value = ""; //clear the URL *description* field
								elem.style.display = "none";
								elem.previousElementSibling.style.display = "none";
								elem.previousElementSibling.previousElementSibling.style.display = "none";
							}
						}
						break;
				}
				
				if (ctx.ListSchema.Field[0].InternalName == "DevAppStatus") {
					var elem = $get("DevAppStatus_" + ctx.ListSchema.Field[0].Id + "_$DropDownChoice");
					//.FieldType: "Choice"
					//.Type: "Choice"
					elem.classList.add("appStatus");
					elem.parentElement.parentElement.parentElement.style.display = "none";
					
					window.PreSaveAction = function() {
						var ddl = document.querySelector(".appStatus");
						if (ddl && ddl.value == "Missing Metadata") //Deleting...
							ddl.value = "In Testing";
						return true;
					}
				}
				
				//Runs for each field, so only execute once at the last field (once all other fields have rendered)
				//if (ctx.ListSchema.Field[0].InternalName == "Attachments") {	
				//}
			}
		}
	});
})();
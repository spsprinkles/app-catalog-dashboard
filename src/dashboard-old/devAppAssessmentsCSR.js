SPClientTemplates.TemplateManager.RegisterTemplateOverrides({
	Templates: {
	},
	OnPreRender: function (ctx) {
		if (ctx.ControlMode < 4 && ctx.ListSchema.Field[0].InternalName == "RelatedApp") {
			var style = document.createElement("style");
			//Add styling to table
			style.innerHTML = "table.ms-formtable td.ms-formlabel { padding-right:25px; } " +
				"table.ms-formtable > tbody > tr > td { font-size:1.1em; border-bottom:1px solid rgb(234,234,234); } " +
				"table.ms-formtable input[type=text], table.ms-formtable div.sp-peoplepicker-topLevel, table.ms-formtable select { min-height:25px; } " +
				"table.ms-formtable input[type=checkbox] { width:20px; height:20px; } " +
				//Enlarge fields
				//".ms-long { width:550px; } " +
				//".sp-peoplepicker-topLevel { width:535px; } " +
				//".sp-peoplepicker-autoFillContainer { max-width:450px; } ";
				//Shorten date fields
				//input[id$='_$DateTimeFieldDate'] { width:112px; }
			document.getElementsByTagName("head")[0].append(style);
		}
	},
	OnPostRender: function (ctx) {
		if (ctx.BaseViewID == "NewForm" && GetUrlKeyValue("relatedId") == "" && ctx.ListSchema.Field[0].InternalName == "RelatedApp") {
			var style = document.createElement("style");
			style.innerHTML = "table.ms-formtable { display:none; }";
			document.getElementsByTagName("head")[0].append(style);
			"stop" in window ? window.stop() : document.execCommand("Stop");
			alert("You must use the App Dashboard to submit app assessments/reviews");
			location.assign(_spPageContextInfo.webAbsoluteUrl + "/SitePages/AppDashboard.aspx");
		}
		
		//Runs for each field, so only execute once at the last field (once all other fields have rendered)
		if (ctx.ListSchema.Field[0].InternalName == "Attachments") {
			var titleField = document.querySelector("input[title='Title Required Field']");
			if (titleField) {
				titleField.value = "Assessment";
				titleField.parentElement.parentElement.parentElement.style.display = "none";
			}
			
			var relatedAppField = document.querySelector("select[title='Related App Required Field']");
			if (relatedAppField && GetUrlKeyValue("relatedId") != "") {
				relatedAppField.value = GetUrlKeyValue("relatedId");
				var appTitle = relatedAppField.querySelector("option[value='" + GetUrlKeyValue("relatedId") + "']").innerText;
				relatedAppField.style.display = "none";
				var span = document.createElement("span");
				span.innerText = appTitle;
				relatedAppField.parentElement.insertBefore(span, relatedAppField);
			}
		}
	}
});
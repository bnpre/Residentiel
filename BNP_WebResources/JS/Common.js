
//Get the attribute value for specified attribute from form.
function getAttributeValue(attributeName) {
	var attribute = Xrm.Page.getAttribute(attributeName);
	var attributeValue = null;
	if (attribute != null) {
		attributeValue = attribute.getValue();
	}
	return attributeValue;
}

//Set the attribute value for specified attribute.
function setAttributeValue(attributeName, attributeValue) {
	var attribute = Xrm.Page.getAttribute(attributeName);
	if (attribute != null) {
		attribute.setValue(attributeValue);
	}
}


//To remove the curly braces from the GUID.
function removeCurlyBraceFromGuid(guid) {
	var returnGuid = guid;
	try {
		if (guid != null &&
        guid.toString().length == 38) {
			returnGuid = returnGuid.replace("{", "");
			returnGuid = returnGuid.replace("}", "");
		}
	}
	catch (error)
	{ }  //Supress error.
	return returnGuid;
}

//To call onDemand Workflow
function startWorkflow(entityRecordId, workflowProcessId) {
	entityRecordId = removeCurlyBraceFromGuid(entityRecordId);
	workflowProcessId = removeCurlyBraceFromGuid(workflowProcessId);
	var executeWorkflowRequest = new Sdk.ExecuteWorkflowRequest(entityRecordId, workflowProcessId);
	var result = Sdk.Sync.execute(executeWorkflowRequest);
}

//To convert the string to Title case
function toTitleCase(str) {
	return str.replace(/\w\S*/g, function (txt) { return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase(); });
}

//To convert the string to Upper case
function toUpperCase(str) {
	return str.toUpperCase();
}





// Programme codé en JavaScript pour gérer l'entité de contact

if (typeof (PackageContact) == "undefined")
{ PackageContact = {}; }
PackageContact.Contact = {

	// C'est la fonction de changer le cas des champs Contact
	onContactSave: function () {//PackageContact.Contact.onContactSave
		var firstName = getAttributeValue("firstname");//Prénom
		var lastName = getAttributeValue("lastname");//Nom
		var lieuDeNaissance = getAttributeValue("resi_lieudenaissance");//Lieu de naissance
		if (firstName != null) {
			firstName = toTitleCase(firstName);
			setAttributeValue("firstname", firstName);
		}
		if (lastName != null) {
			lastName = toUpperCase(lastName);
			setAttributeValue("lastname", lastName);
		}
		if (lieuDeNaissance != null) {
			lieuDeNaissance = toTitleCase(lieuDeNaissance);
			setAttributeValue("resi_lieudenaissance", lieuDeNaissance);
		}
	},

	// C'est la fonction à appeler sur le chargement du formulaire de contact
	onContactLoad: function () {//PackageContact.Contact.onContactLoad
		//Part 1 -->build fetchxml
		var viewId = "{99D5E768-7E2A-E711-8101-5065F38BD4E0}"; //Any Guid is fine.
		var entityName = "account";// Entity to be filtered
		var viewDisplayName = "RESI - AgenceType" 

		var fetchXml = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
			"<entity name='account'>" +
				"<attribute name='name' />" +
				"<order attribute='name' descending='false' /> " +
				"<filter type='and'>" +
					"<condition attribute='resi_typeagencecode' operator='not-null' />" +
				"</filter>" +
		 "</entity>" +
		"</fetch>";

		// Part 2 --> Build Grid Layout. In other words, building a custom view for the Lookup
		//building grid layout
		var layoutXml = "<grid name='resultset' " +
										"object='1' " +
										"jump='name' " +
										"select='1' " +
										"icon='1' " +
										"preview='1'>" +
										"<row name='result' " +
										"id='accountid'>" +
										"<cell name='name' " +
										"width='300' />" +
										"<cell name='resi_typeagencecode' " +
										"width='150' />" +
										"disableSorting='1' />" +
										"</row>" +
										"</grid>";

		setTimeout(function () {
			Xrm.Page.getControl("parentcustomerid").addCustomView(viewId, entityName, viewDisplayName, fetchXml, layoutXml, true);
		}, 1000);
	},
}


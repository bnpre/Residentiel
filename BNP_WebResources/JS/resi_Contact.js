// Programme cod� en JavaScript pour g�rer l'entit� de contact

if (typeof (PackageContact) == "undefined")
{ PackageContact = {}; }
PackageContact.Contact = {

	// C'est la fonction de changer le cas des champs Contact
	onContactSave: function () {//PackageContact.Contact.onContactSave
		var firstName = getAttributeValue("firstname");//Pr�nom
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
}


// Programme codé en JavaScript permettant de gérer l'entité Programme

if (typeof (PackageProgramme) == "undefined")
{ PackageProgramme = {}; }
PackageProgramme.Programme = {

	// C'est la fonction d'ouvrir le formulaire "Sales Literature"
	onClickDocumentDirectoryButton: function () {//PackageProgramme.Programme.onClickDocumentDirectoryButton
		var recordId = Xrm.Page.data.entity.getId();
		var parameters = {};
		parameters["resi_associatedprogrammeid"] = recordId;
		var options = {
			openInNewWindow: true
		};
		Xrm.Utility.openEntityForm("salesliterature", null, parameters, options);
	},
}
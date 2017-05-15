// Programme codé en JavaScript permettant de gérer l'entité Contrat

if (typeof (PackageContrat) == "undefined")
{ PackageContrat = {}; }
PackageContrat.Contrat = {

  // Il s'agit de la fonction pour résoudre le Contrat 
  onClickCloseButton: function () {//PackageContrat.Contrat.onClickCloseButton
    var actionName = "resi_CancelContrat";
    var serverURL = Xrm.Page.context.getClientUrl();
    var recordEntityName = Xrm.Page.data.entity.getEntityName();
    var recordId = Xrm.Page.data.entity.getId();
    Process.callAction(actionName, [{
      key: "Target",
      type: Process.Type.EntityReference,
      value: { id: recordId, entityType: recordEntityName }
    }],
							function (response) {
							  alert("Contrat annulé avec succès");//Contrat Cancelled successfully.
							  //Reload the form.
							  Xrm.Utility.openEntityForm(recordEntityName, recordId);
							},
							function (error) {
							  alert("Error executing action : " + error);
							},
							serverURL
						);
  },

  //Il s'agit d'une fonction pour créer un contrat xml
  onClickCreateXMLButton: function () {//PackageContrat.Contrat.onClickCreateXMLButton
    var actionName = "resi_GenerateContractXML";
    var recordId = Xrm.Page.data.entity.getId();
    recordId = removeCurlyBraceFromGuid(recordId);
    var serverURL = Xrm.Page.context.getClientUrl();

    var xmlHttpRequest = new XMLHttpRequest();
    xmlHttpRequest.open("POST", serverURL + "/api/data/v8.2/resi_contrats(" + recordId + ")/Microsoft.Dynamics.CRM." + actionName, false);
    xmlHttpRequest.setRequestHeader("Accept", "application/json");
    xmlHttpRequest.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    xmlHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
    xmlHttpRequest.setRequestHeader("OData-Version", "4.0");
    xmlHttpRequest.onreadystatechange = function () {
      if (this.readyState == 4 /* complete */) {
        xmlHttpRequest.onreadystatechange = null;
        if (this.status == 200 ||
						this.status == 204) {
          alert("Contrat XML created successfully.");
          //Reload the form.
          var recordEntityName = Xrm.Page.data.entity.getEntityName();
          Xrm.Utility.openEntityForm(recordEntityName, recordId);
        }
        else {
          var error = JSON.parse(this.response).error;
          alert(error.message);
        }
      }
    };
    xmlHttpRequest.send();
  },

  //Il s'agit d'une fonction appelant le workflow pour générer le document
  onClickGenerateDocumentButton: function () {//PackageContrat.Contrat.onClickGenerateDocumentButton
    var actionName = "resi_SetContratDocumentTemplate";
    var serverURL = Xrm.Page.context.getClientUrl();
    var recordEntityName = Xrm.Page.data.entity.getEntityName();
    var recordId = Xrm.Page.data.entity.getId();
    Process.callAction(actionName, [{
      key: "Target",
      type: Process.Type.EntityReference,
      value: { id: recordId, entityType: recordEntityName }
    }],
							function (response) {
							  //Reload the form.
							  Xrm.Utility.openEntityForm(recordEntityName, recordId);
							},
							function (error) {
							  alert("Error executing action : " + error);
							},
							serverURL
						);
  },

  //Il s'agit d'une fonction js appelant le workflow pour générer le mot document au niveau Global
  onClickGenerateDocumentForContract: function () {//PackageContrat.Contrat.onClickGenerateDocumentForContract
    var workflowId = "99f3506b-6634-4f3c-802f-389e3a91217c";//Générer le contrat (Global)
    var serverURL = Xrm.Page.context.getClientUrl();
    var recordEntityName = Xrm.Page.data.entity.getEntityName();
    var recordId = Xrm.Page.data.entity.getId();
    Process.callWorkflow(workflowId, recordId,
							function () {
							  //Reload the form.
							  Xrm.Utility.openEntityForm(recordEntityName, recordId);
							},
							function () {
							  alert("Error executing action : " + error);
							},
							serverURL
						);
  },

  //C'est une fonction à exécuter sur le clic du bouton docusign
  onClickDocusignButton: function () {//PackageContrat.Contrat.onClickDocusignButton
    alert("Un peu de patience. DOCUSIGN  arrive bientôt");
  },

  //C'est une fonction à exécuter sur la charge du formulaire d'entité Contrat
  onContratLoad: function () {//PackageContrat.Contrat.onContratLoad
    var formType = Xrm.Page.ui.getFormType();
    if (formType != 1) //Not in Create mode
    {
      Xrm.Page.ui.controls.get("resi_concernerid").setDisabled(true);
      Xrm.Page.ui.controls.get("resi_opportuniteid").setDisabled(true);
    }
  },

  //Il s'agit d'une fonction à effectuer sur la sauvegarde du formulaire de l'entité Contrat
  onContratSave: function () {//PackageContrat.Contrat.onContratSave
    //Set the Field values to "0" in case of null
    var fraisDeNotaireOffert = getAttributeValue("resi_fraisdenotaireoffert");
    var tMAOffert = getAttributeValue("resi_tmaoffert");
    var autres = getAttributeValue("resi_autres");
    var appelsDeFondsAdaptes = getAttributeValue("resi_appelsdefondsadaptes");
    var baisseDePrixDirect = getAttributeValue("resi_aidesalaventeenbaissedeprix");

    if (fraisDeNotaireOffert == null) {
      setAttributeValue("resi_fraisdenotaireoffert", 0);
    }
    if (tMAOffert == null) {
      setAttributeValue("resi_tmaoffert", 0);
    }
    if (autres == null) {
      setAttributeValue("resi_autres", 0);
    }
    if (appelsDeFondsAdaptes == null) {
      setAttributeValue("resi_appelsdefondsadaptes", 0);
    }
    if (baisseDePrixDirect == null) {
      setAttributeValue("resi_aidesalaventeenbaissedeprix", 0);
    }
  },

  //C'est une fonction à effectuer sur la modification du champ "Signé par les clients"?
  onSigneParLesClientsChange: function () {//PackageContrat.Contrat.onSigneParLesClientsChange
    //Update "Date de Signature" on "Signé par les clients?" change
    var signeParLesClients = getAttributeValue("resi_issigneparlesclients");

    if (signeParLesClients == true) {
      setAttributeValue("resi_datedesignaturepapier", Date.now());
    }
  },
}
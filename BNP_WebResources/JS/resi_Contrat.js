// Programme codé en JavaScript permettant de gérer l'entité Contrat

if (typeof (PackageContrat) == "undefined")
{ PackageContrat = {}; }
PackageContrat.Contrat = {

  // Il s'agit de la fonction pour résoudre le Contrat 
  onClickCloseButton: function () {//PackageContrat.Contrat.onClickCloseButton
    var actionName = "resi_CancelContrat";
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
          alert("Contrat annulé avec succès");//Contrat Cancelled successfully.
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

  //Il s'agit d'une fonction appelant le workflow pour générer le mot document
  onClickGenerateDocumentButton: function () {//PackageContrat.Contrat.onClickGenerateDocumentButton
    var recordId = Xrm.Page.data.entity.getId();
    recordId = removeCurlyBraceFromGuid(recordId);
    //var workflowId = "3C0BA9DC-131F-4D83-9B6D-5E5241EF8AFE";//Workflow: Generate and attach Contrat document
    //startWorkflow(recordId, workflowId);

    var actionName = "resi_SetContratDocumentTemplate";
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

  //C'est une fonction à exécuter sur le clic du bouton docusign
  onClickDocusignButton: function () {//PackageContrat.Contrat.onClickDocusignButton
    alert("Un peu de patience. DOCUSIGN  arrive bientôt");
  },
}
// Programme codé en JavaScript permettant de gérer l'entité Opportunity

if (typeof (PackageOpportunity) == "undefined")
{ PackageOpportunity = {}; }
PackageOpportunity.Opportunity = {

  //Refresh "WebResource_MultiCanalContact" HTML Web Resource on Change of Contact Lookup
  onContactChange: function () {//PackageOpportunity.Opportunity.onContactChange
    Xrm.Page.data.entity.save(null);
    var wrRelatedOpportunities = Xrm.Page.ui.controls.get("WebResource_RelatedOpportunities");
    var src = wrRelatedOpportunities.getSrc();
    wrRelatedOpportunities.setSrc(null);
    wrRelatedOpportunities.setSrc(src);
  },

  // Il s'agit de la fonction pour résoudre le opportunity en utilisant la boite de dialogue
  onClickCloseButton: function () {//PackageOpportunity.Opportunity.onClickCloseButton
    var opportunityId = Xrm.Page.data.entity.getId();
    var entityName = Xrm.Page.data.entity.getEntityName();
    opportunityId = opportunityId.replace("{", "").replace("}", "");
    Xrm.Utility.confirmDialog("Toutes les activités ouvertes associées à cette demande seront annulées si vous décidez de l'annuler. Voulez-vous continuer", function () {
      //Open a dialogue.
      var dialogId = "8907D82E-7B75-4CA3-AA2B-7A3178EF8A69"; // Boite de dialogue pour clôre la demande
      var entityName = "opportunity";

      var url = Xrm.Page.context.getClientUrl() +
       "/cs/dialog/rundialog.aspx?DialogId=" +
       dialogId + "&EntityName=" +
       entityName + "&ObjectId=" +
       opportunityId;
      window.showModalDialog(encodeURI(url), "", "dialogWidth:550px; dialogHeight:300px;resizable:yes;");
      Xrm.Utility.openEntityForm(entityName, opportunityId);
    })
  },

  // Remplir «Agence» sur Prescripteur Change
  onPrescripteurChange: function () {//PackageOpportunity.Opportunity.onPrescripteurChange
    var prescripteurId = getLookupAttributeId("resi_conseillerbddfid");
    if (prescripteurId != null) {
      prescripteurId = removeCurlyBraceFromGuid(prescripteurId);
      var contactFetchQuery = ["<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'> ",
																	" <entity name='contact'> ",
																	"   <attribute name='contactid' /> ",
																	"   <attribute name='parentcustomerid' /> ",
																	"   <filter type='and'> ",
																	"     <condition attribute='contactid' operator='eq' uitype='contact' value='" + prescripteurId + "' /> ",
																	"   </filter> ",
																	" </entity> ",
																	"</fetch>"].join("");
      var contactFetchExpression = new Sdk.Query.FetchExpression(contactFetchQuery);
      var fetchedContacts = Sdk.Sync.retrieveMultiple(contactFetchExpression);
      if (fetchedContacts != null) {
        var fetchedContactsCount = fetchedContacts.getEntities().getCount();
        if (fetchedContactsCount == 1) {
          var fetchedContact = fetchedContacts.getEntities().getByIndex(0);
          if (fetchedContact != null) {
            var parentCustomerId = getSdkCallFieldValue(fetchedContact, "parentcustomerid");
            if (parentCustomerId != null &&
								parentCustomerId != "") {
              setLookupAttributeValue("resi_agenceid", parentCustomerId.getId(), parentCustomerId.getName(), parentCustomerId.getType());
            }
            else {
              setAttributeValue("resi_agenceid", null);
            }
          }
        }
      }
    }
    else {
      setAttributeValue("resi_agenceid", null);
    }
  },
  onOpportunityLoad: function () {//PackageOpportunity.Opportunity.onOpportunityLoad
    var formType = Xrm.Page.ui.getFormType();
    if (formType == 1) //Only in Create mode
    {
      setAttributeValue("resi_origine", "Passage BV");
      setAttributeValue("resi_searchdate", new Date());
    }
  },
}
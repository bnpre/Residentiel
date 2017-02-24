// Programme cod� en JavaScript permettant de g�rer l'entit� Opportunity

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

  // Il s'agit de la fonction pour r�soudre le opportunity en utilisant la boite de dialogue
  onClickCloseButton: function () {//PackageOpportunity.Opportunity.onClickCloseButton
    var opportunityId = Xrm.Page.data.entity.getId();
    var entityName = Xrm.Page.data.entity.getEntityName();
    opportunityId = opportunityId.replace("{", "").replace("}", "");
    Xrm.Utility.confirmDialog("Toutes les activit�s ouvertes associ�es � cette demande seront annul�es si vous d�cidez de l'annuler. Voulez-vous continuer", function () {
      //Open a dialogue.
      var dialogId = "8907D82E-7B75-4CA3-AA2B-7A3178EF8A69"; // Boite de dialogue pour cl�re la demande
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
}
//Refresh "WebResource_MultiCanalContact" HTML Web Resource on Change of Contact Lookup
function onContactChange() {
  Xrm.Page.data.entity.save(null);
  var wrRelatedOpportunities = Xrm.Page.ui.controls.get("WebResource_RelatedOpportunities");
  var src = wrRelatedOpportunities.getSrc();
  wrRelatedOpportunities.setSrc(null);
  wrRelatedOpportunities.setSrc(src);
}
// Programme codé en JavaScript permettant de gérer l'entité Echéancier

if (typeof (PackageEcheancier) == "undefined")
{ PackageEcheancier = {}; }
PackageEcheancier.Echeancier = {
  onEcheancierSave: function () {//PackageEcheancier.Echeancier.onEcheancierSave
    //First get the program or contract to identofy that,
    //Either we use Initial or Adapte
    var programLookup = getAttributeValue("resi_programmeid");
    var contratLookup = getAttributeValue("resi_contratid");
    var fieldToCalculateSum = "";
    var currentValue = 0;
    var currentRecordId = Xrm.Page.data.entity.getId();
    currentRecordId = removeCurlyBraceFromGuid(currentRecordId);

    if (contratLookup != null) {
      fieldToCalculateSum = "resi_adapte";
      currentValue = getAttributeValue("resi_adapte");
    }
    else if (programLookup != null) {
      fieldToCalculateSum = "resi_initial";
      currentValue = getAttributeValue("resi_initial");
    }

    //Check that is user changed the field then only fire the validation.
    var fieldIsDirty = getIsDirty(fieldToCalculateSum);
    var validateCumulePercent = false;
    var fetchXMLCondition = "";
    var currentRecordCondition = "";
    if (currentRecordId != null &&
        currentRecordId != "") {
      currentRecordCondition = " <condition attribute='resi_echeancierid' operator='ne' uiname='' uitype='resi_echeancier' value='" + currentRecordId + "' /> ";
    }

    if (fieldIsDirty == true) {
      //Check that the pregam or contrat entity have allowedsale = true or Contrat Valide = true.

      if (contratLookup != null) {
        //Fetch the contract record to validate contrat valide is true or false.
        var contratId = getLookupAttributeId("resi_contratid");
        contratId = removeCurlyBraceFromGuid(contratId);
        contratColumnset = new Sdk.ColumnSet("resi_iscontratvalide");
        fetchXMLCondition = " <condition attribute='resi_contratid' operator='eq' uiname='' uitype='resi_contrat' value='" + contratId + "' />";

        var contrat = Sdk.Sync.retrieve("resi_contrat", contratId, contratColumnset);
        if (contrat != null) {
          validateCumulePercent = getSdkCallFieldValue(contrat, "resi_iscontratvalide");
        }
      } else if (programLookup != null) {
        //Fetch the contract record to validate contrat valide is true or false.
        var programId = getLookupAttributeId("resi_programmeid");
        programId = removeCurlyBraceFromGuid(programId);
        programColumnset = new Sdk.ColumnSet("resi_allowedsale");
        fetchXMLCondition = " <condition attribute='resi_programmeid' operator='eq' uiname='' uitype='resi_programme' value='" + programId + "' />";

        var program = Sdk.Sync.retrieve("resi_programme", programId, programColumnset);
        if (program != null) {
          validateCumulePercent = getSdkCallFieldValue(program, "resi_allowedsale");
        }
      }

      //Fetch the sum of the specified field to alert the user.
      if (validateCumulePercent == true) {
        var echeancierFetchQuery = ["<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>",
                                    "  <entity name='resi_echeancier'>",
                                    "    <attribute name='resi_echeancierid' />",
                                    "    <attribute name='resi_initial' />",
                                    "    <attribute name='resi_cumule' />",
                                    "    <attribute name='resi_adapte' />",
                                    "    <filter type='and'>",
                                    fetchXMLCondition,
                                    currentRecordCondition,
                                    "    </filter>",
                                    "  </entity>",
                                    "</fetch>"].join("");
        var sumValue = 0;
        var echeancierFetchExpression = new Sdk.Query.FetchExpression(echeancierFetchQuery);
        var fetchedEcheanciers = Sdk.Sync.retrieveMultiple(echeancierFetchExpression);
        if (fetchedEcheanciers != null) {
          var fetchedEcheancierCount = fetchedEcheanciers.getEntities().getCount();
          for (i = 0; i < fetchedEcheancierCount; i++) {
            var fetchedEcheancier = fetchedEcheanciers.getEntities().getByIndex(i);
            if (fetchedEcheancier != null) {
              var fieldToCalculateSumValue = getSdkCallFieldValue(fetchedEcheancier, fieldToCalculateSum);
              if (fieldToCalculateSumValue != null &&
                  fieldToCalculateSumValue != "") {
                sumValue += fieldToCalculateSumValue;
              }
            }
          }

          //Add current record value.
          sumValue += currentValue;

          if (sumValue != 100) {
            Xrm.Utility.alertDialog("Le cumul pour l'ensemble des échéanciers ne fait pas 100%. Veuillez ajuster les étapes d’échéance pour avoir le cumul souhaité de 100%")
          }
        }
      }
    }
  },

  onEcheancierLoad: function () {//PackageEcheancier.Echeancier.onEcheancierLoad
    debugger;
    var programLookup = getAttributeValue("resi_programmeid");
    var contratLookup = getAttributeValue("resi_contratid");
    if (contratLookup != null) {
      setVisibleTabSection("Initial", "General", false);
      Xrm.Page.getAttribute("resi_adapte").setRequiredLevel("required");
    }
    else if (programLookup != null) {
      setVisibleTabSection("Adapte", "tab_3_section_1", false);
      Xrm.Page.getAttribute("resi_initial").setRequiredLevel("required");
    }
  },
}
using BNP_Model.Utils;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace BNP_Workflows
{
  public class SetWordTemplate : CodeActivity
  {
    /// <summary>
    /// Executes the specified execution context.
    /// </summary>
    /// <param name="executionContext">The execution context.</param>
    /// <exception cref="InvalidPluginExecutionException">The timeout elapsed while attempting to issue the request.</exception>
    protected override void Execute(CodeActivityContext context)
    {
      //Create the tracing service
      ITracingService tracingService = context.GetExtension<ITracingService>();

      //Create the context
      IWorkflowContext wfContext = context.GetExtension<IWorkflowContext>();
      IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
      IOrganizationService service = serviceFactory.CreateOrganizationService(null);

      tracingService.Trace("Getting contract");
      //Fetch the contract entity record and pick the Program and TVA
      Entity contract = CRMData.GetEntityRecordByGuid(service, "resi_contrat", wfContext.PrimaryEntityId, "resi_concernerid", "resi_tvatype");
      tracingService.Trace("Contract got");

      tracingService.Trace("Getting programme from Contract");
      EntityReference programmeEntityReference = CRMData.GetAttributeValue<EntityReference>(contract, null, "resi_concernerid");
      string programme = string.Empty;
      if (programmeEntityReference != null)
      { programme = programmeEntityReference.Name; }

      tracingService.Trace("Programme  = " + programme);

      tracingService.Trace("Getting tva from Contract");
      string tva = CRMData.GetFormattedAttributeValue(contract, null, "resi_tvatype");
      tracingService.Trace("TVA = " + tva);

      if (string.IsNullOrWhiteSpace(programme) == false &&
          string.IsNullOrWhiteSpace(tva) == false)
      {
        tracingService.Trace("Program and TVS Found");

        string contractNameCondition = string.Format("%{0}_{1}_", programme, tva);
        tracingService.Trace("contractNameCondition = " + contractNameCondition);

        tracingService.Trace("Fetching document templates ");
        QueryExpression documentTemplateQueryExpression = new QueryExpression("documenttemplate");
        documentTemplateQueryExpression.ColumnSet = new ColumnSet("name");
        documentTemplateQueryExpression.Criteria.AddCondition("documenttype", ConditionOperator.Equal, 2);//Word document
        documentTemplateQueryExpression.Criteria.AddCondition("name", ConditionOperator.Like, contractNameCondition);//Word document

        EntityCollection documentTemplateEntityCollection = service.RetrieveMultiple(documentTemplateQueryExpression);
        if (documentTemplateEntityCollection.Entities.Count > 0)
        {
          tracingService.Trace("Doeumnet exists");
          tracingService.Trace("Executing code to generate the document");

          OrganizationRequest organizationRequest = new OrganizationRequest("SetWordTemplate");
          organizationRequest["Target"] = new EntityReference(wfContext.PrimaryEntityName, wfContext.PrimaryEntityId);
          organizationRequest["SelectedTemplate"] = new EntityReference("documenttemplate", documentTemplateEntityCollection.Entities[0].Id);
          service.Execute(organizationRequest);
          tracingService.Trace("document generated successfully");
        }
        else
        { tracingService.Trace("Doeumnet does not exists"); }
      }
      else
      { tracingService.Trace("Program or TVS not Found"); }
    }
  }
}

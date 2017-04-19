using BNP_Model.Utils;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace BNP_Workflows
{
	public class SetDocumentTemplateId : CodeActivity
	{
		/// <summary>
		/// Gets or sets the input.
		/// </summary>
		/// <value>
		/// The input.
		/// </value>
		[Input("Input")]
		public InArgument<string> Input { get; set; }

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

			tracingService.Trace("Workflow Execution started");

			tracingService.Trace("Contract got");

			tracingService.Trace("Fetching document templates ");
			tracingService.Trace("Input :{0}", Input);
			QueryExpression documentTemplateQueryExpression = new QueryExpression("documenttemplate");
			documentTemplateQueryExpression.ColumnSet = new ColumnSet("documenttemplateid");
			documentTemplateQueryExpression.Criteria.AddCondition("name", ConditionOperator.Equal, this.Input.Get(context));

			EntityCollection documentTemplateEntityCollection = service.RetrieveMultiple(documentTemplateQueryExpression);
			if (documentTemplateEntityCollection.Entities.Count > 0)
			{
				tracingService.Trace("documentTemplateEntityCollection.Entities.Count:{0}", documentTemplateEntityCollection.Entities.Count);
				Guid documentTemplateId = CRMData.GetAttributeValue<Guid>(documentTemplateEntityCollection.Entities[0], null, "documenttemplateid");
				tracingService.Trace("documentTemplateId {0}", documentTemplateId.ToString());
				Entity mappingProgrammeDocument = new Entity("resi_mappingprogrammedocument");
				mappingProgrammeDocument.Id = wfContext.PrimaryEntityId;
				tracingService.Trace("Assigning the Id of the current entity");
				CRMData.AddAttribute<string>(mappingProgrammeDocument, "resi_documenttemplateid", documentTemplateId.ToString());
				service.Update(mappingProgrammeDocument);
			}
		}
	}
}

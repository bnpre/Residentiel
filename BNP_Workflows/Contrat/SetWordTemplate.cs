using BNP_Model.Utils;
using Microsoft.Crm.Sdk.Messages;
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

			EntityReference contractProgramme = CRMData.GetAttributeValue<EntityReference>(contract, null, "resi_concernerid");
			OptionSetValue contractTVA = CRMData.GetAttributeValue<OptionSetValue>(contract, null, "resi_tvatype");
			string documentTemplateId = string.Empty;
			string workflowIdToGenerateDocument = "7770ad66-800d-453a-a718-b069b52fefb6";//Workflow: Set Contrat Document Template

			if (contractProgramme != null && contractTVA != null)
			{
				tracingService.Trace("contractProgramme != null && contractTVA != null");
				QueryExpression mappingProgrammeDocumentQueryExpression = new QueryExpression("resi_mappingprogrammedocument");
				mappingProgrammeDocumentQueryExpression.ColumnSet = new ColumnSet("resi_documenttemplateid");
				mappingProgrammeDocumentQueryExpression.Criteria.AddCondition("resi_programmeid", ConditionOperator.Equal, contractProgramme.Id);
				mappingProgrammeDocumentQueryExpression.Criteria.AddCondition("resi_tvatype", ConditionOperator.Equal, contractTVA.Value);
				EntityCollection mappingProgrammeDocumentEntityCollection = service.RetrieveMultiple(mappingProgrammeDocumentQueryExpression);

				if (mappingProgrammeDocumentEntityCollection.Entities.Count > 0)
				{
					tracingService.Trace("mappingProgrammeDocumentEntityCollection.Entities.Count > 0");
					documentTemplateId = mappingProgrammeDocumentEntityCollection.Entities[0].GetAttributeValue<string>("resi_documenttemplateid");

					if (string.IsNullOrWhiteSpace(documentTemplateId) == false)
					{
						tracingService.Trace("documentTemplateId != null");

						#region Create XAML

						// Define the workflow XAML.
						string mainXaml;

						mainXaml = @"<Activity x:Class=""XrmWorkflow7770ad66800d453aa718b069b52fefb6"" xmlns=""http://schemas.microsoft.com/netfx/2009/xaml/activities"" xmlns:mcwa=""clr-namespace:Microsoft.Crm.Workflow.Activities;assembly=Microsoft.Crm.Workflow, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"" xmlns:mva=""clr-namespace:Microsoft.VisualBasic.Activities;assembly=System.Activities, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"" xmlns:mxs=""clr-namespace:Microsoft.Xrm.Sdk;assembly=Microsoft.Xrm.Sdk, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"" xmlns:mxswa=""clr-namespace:Microsoft.Xrm.Sdk.Workflow.Activities;assembly=Microsoft.Xrm.Sdk.Workflow, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"" xmlns:s=""clr-namespace:System;assembly=mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"" xmlns:scg=""clr-namespace:System.Collections.Generic;assembly=mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"" xmlns:sco=""clr-namespace:System.Collections.ObjectModel;assembly=mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"" xmlns:srs=""clr-namespace:System.Runtime.Serialization;assembly=System.Runtime.Serialization, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"" xmlns:this=""clr-namespace:"" xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml"">
				          <x:Members>
				            <x:Property Name=""InputEntities"" Type=""InArgument(scg:IDictionary(x:String, mxs:Entity))"" />
				            <x:Property Name=""CreatedEntities"" Type=""InArgument(scg:IDictionary(x:String, mxs:Entity))"" />
				          </x:Members>
				          <this:XrmWorkflow7770ad66800d453aa718b069b52fefb6.InputEntities>
				            <InArgument x:TypeArguments=""scg:IDictionary(x:String, mxs:Entity)"" />
				          </this:XrmWorkflow7770ad66800d453aa718b069b52fefb6.InputEntities>
				          <this:XrmWorkflow7770ad66800d453aa718b069b52fefb6.CreatedEntities>
				            <InArgument x:TypeArguments=""scg:IDictionary(x:String, mxs:Entity)"" />
				          </this:XrmWorkflow7770ad66800d453aa718b069b52fefb6.CreatedEntities>
				          <mva:VisualBasic.Settings>Assembly references and imported namespaces for internal implementation</mva:VisualBasic.Settings>
				          <mxswa:Workflow>
				            <mxswa:ActivityReference AssemblyQualifiedName=""Microsoft.Crm.Workflow.Activities.Composite, Microsoft.Crm.Workflow, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"" DisplayName=""InvokeSdkMessageStep1"">
				              <mxswa:ActivityReference.Properties>
				                <sco:Collection x:TypeArguments=""Variable"" x:Key=""Variables"">
				                  <Variable x:TypeArguments=""x:Object"" Name=""InvokeSdkMessageStep1_1"" />
				                  <Variable x:TypeArguments=""x:Object"" Name=""InvokeSdkMessageStep1_2"" />
				                  <Variable x:TypeArguments=""x:Object"" Name=""InvokeSdkMessageStep1_1_converted"" />
				                  <Variable x:TypeArguments=""x:Object"" Name=""InvokeSdkMessageStep1_3"" />
				                  <Variable x:TypeArguments=""x:Object"" Name=""InvokeSdkMessageStep1_4"" />
				                  <Variable x:TypeArguments=""x:Object"" Name=""InvokeSdkMessageStep1_3_converted"" />
				                </sco:Collection>
				                <sco:Collection x:TypeArguments=""Activity"" x:Key=""Activities"">
				                  <mxswa:ActivityReference AssemblyQualifiedName=""Microsoft.Crm.Workflow.Activities.EvaluateExpression, Microsoft.Crm.Workflow, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"" DisplayName=""EvaluateExpression"">
				                    <mxswa:ActivityReference.Arguments>
				                      <InArgument x:TypeArguments=""x:String"" x:Key=""ExpressionOperator"">CreateCrmType</InArgument>
				                      <InArgument x:TypeArguments=""s:Object[]"" x:Key=""Parameters"">[New Object() { Microsoft.Xrm.Sdk.Workflow.WorkflowPropertyType.Guid, ""{0}"", ""UniqueIdentifier"" }]</InArgument>
				                      <InArgument x:TypeArguments=""s:Type"" x:Key=""TargetType"">
				                        <mxswa:ReferenceLiteral x:TypeArguments=""s:Type"" Value=""mxs:EntityReference"" />
				                      </InArgument>
				                      <OutArgument x:TypeArguments=""x:Object"" x:Key=""Result"">[InvokeSdkMessageStep1_2]</OutArgument>
				                    </mxswa:ActivityReference.Arguments>
				                  </mxswa:ActivityReference>
				                  <mxswa:ActivityReference AssemblyQualifiedName=""Microsoft.Crm.Workflow.Activities.EvaluateExpression, Microsoft.Crm.Workflow, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"" DisplayName=""EvaluateExpression"">
				                    <mxswa:ActivityReference.Arguments>
				                      <InArgument x:TypeArguments=""x:String"" x:Key=""ExpressionOperator"">CreateCrmType</InArgument>
				                      <InArgument x:TypeArguments=""s:Object[]"" x:Key=""Parameters"">[New Object() { Microsoft.Xrm.Sdk.Workflow.WorkflowPropertyType.EntityReference, ""documenttemplate"", ""Modèle de Contrat - TVA 20"", InvokeSdkMessageStep1_2, ""Lookup"" }]</InArgument>
				                      <InArgument x:TypeArguments=""s:Type"" x:Key=""TargetType"">
				                        <mxswa:ReferenceLiteral x:TypeArguments=""s:Type"" Value=""mxs:EntityReference"" />
				                      </InArgument>
				                      <OutArgument x:TypeArguments=""x:Object"" x:Key=""Result"">[InvokeSdkMessageStep1_1]</OutArgument>
				                    </mxswa:ActivityReference.Arguments>
				                  </mxswa:ActivityReference>
				                  <mxswa:ActivityReference AssemblyQualifiedName=""Microsoft.Crm.Workflow.Activities.ConvertCrmXrmTypes, Microsoft.Crm.Workflow, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"" DisplayName=""ConvertCrmXrmTypes"">
				                    <mxswa:ActivityReference.Arguments>
				                      <InArgument x:TypeArguments=""x:Object"" x:Key=""Value"">[InvokeSdkMessageStep1_1]</InArgument>
				                      <InArgument x:TypeArguments=""s:Type"" x:Key=""TargetType"">
				                        <mxswa:ReferenceLiteral x:TypeArguments=""s:Type"" Value=""mxs:EntityReference"" />
				                      </InArgument>
				                      <OutArgument x:TypeArguments=""x:Object"" x:Key=""Result"">[InvokeSdkMessageStep1_1_converted]</OutArgument>
				                    </mxswa:ActivityReference.Arguments>
				                  </mxswa:ActivityReference>
				                  <mxswa:GetEntityProperty Attribute=""resi_contratid"" Entity=""[InputEntities(&quot;primaryEntity&quot;)]"" EntityName=""resi_contrat"" Value=""[InvokeSdkMessageStep1_4]"">
				                    <mxswa:GetEntityProperty.TargetType>
				                      <InArgument x:TypeArguments=""s:Type"">
				                        <mxswa:ReferenceLiteral x:TypeArguments=""s:Type"" Value=""mxs:EntityReference"" />
				                      </InArgument>
				                    </mxswa:GetEntityProperty.TargetType>
				                  </mxswa:GetEntityProperty>
				                  <mxswa:ActivityReference AssemblyQualifiedName=""Microsoft.Crm.Workflow.Activities.EvaluateExpression, Microsoft.Crm.Workflow, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"" DisplayName=""EvaluateExpression"">
				                    <mxswa:ActivityReference.Arguments>
				                      <InArgument x:TypeArguments=""x:String"" x:Key=""ExpressionOperator"">SelectFirstNonNull</InArgument>
				                      <InArgument x:TypeArguments=""s:Object[]"" x:Key=""Parameters"">[New Object() { InvokeSdkMessageStep1_4 }]</InArgument>
				                      <InArgument x:TypeArguments=""s:Type"" x:Key=""TargetType"">
				                        <mxswa:ReferenceLiteral x:TypeArguments=""s:Type"" Value=""mxs:EntityReference"" />
				                      </InArgument>
				                      <OutArgument x:TypeArguments=""x:Object"" x:Key=""Result"">[InvokeSdkMessageStep1_3]</OutArgument>
				                    </mxswa:ActivityReference.Arguments>
				                  </mxswa:ActivityReference>
				                  <mxswa:ActivityReference AssemblyQualifiedName=""Microsoft.Crm.Workflow.Activities.ConvertCrmXrmTypes, Microsoft.Crm.Workflow, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"" DisplayName=""ConvertCrmXrmTypes"">
				                    <mxswa:ActivityReference.Arguments>
				                      <InArgument x:TypeArguments=""x:Object"" x:Key=""Value"">[InvokeSdkMessageStep1_3]</InArgument>
				                      <InArgument x:TypeArguments=""s:Type"" x:Key=""TargetType"">
				                        <mxswa:ReferenceLiteral x:TypeArguments=""s:Type"" Value=""mxs:EntityReference"" />
				                      </InArgument>
				                      <OutArgument x:TypeArguments=""x:Object"" x:Key=""Result"">[InvokeSdkMessageStep1_3_converted]</OutArgument>
				                    </mxswa:ActivityReference.Arguments>
				                  </mxswa:ActivityReference>
				                  <mcwa:InvokeSdkMessageActivity SdkMessageEntityName=""{x:Null}"" DisplayName=""InvokeSdkMessageStep1"" SdkMessageId=""8a8aebf5-d462-4ccb-a340-3103fbe7c5c5"" SdkMessageName=""SetWordTemplate"" SdkMessageRequestSuffix="""">
				                    <mcwa:InvokeSdkMessageActivity.Arguments>
				                      <InArgument x:TypeArguments=""mxs:EntityReference"" x:Key=""SelectedTemplate"">[DirectCast(InvokeSdkMessageStep1_1_converted, Microsoft.Xrm.Sdk.EntityReference)]</InArgument>
				                      <InArgument x:TypeArguments=""mxs:EntityReference"" x:Key=""Target"">[DirectCast(InvokeSdkMessageStep1_3_converted, Microsoft.Xrm.Sdk.EntityReference)]</InArgument>
				                    </mcwa:InvokeSdkMessageActivity.Arguments>
				                    <mcwa:InvokeSdkMessageActivity.Properties>
				                      <sco:Collection x:TypeArguments=""x:String"" x:Key=""SelectedTemplate#InArgumentEntityType"">
				                        <x:String>documenttemplate</x:String>
				                      </sco:Collection>
				                      <x:String x:Key=""SelectedTemplate#InArgumentRequired"">SelectedTemplate</x:String>
				                      <sco:Collection x:TypeArguments=""x:String"" x:Key=""Target#InArgumentEntityType"">
				                        <x:String>account</x:String>
				                        <x:String>team</x:String>
				                        <x:String>campaignactivity</x:String>
				                        <x:String>quote</x:String>
				                        <x:String>resi_contrat</x:String>
				                        <x:String>email</x:String>
				                        <x:String>product</x:String>
				                        <x:String>recurringappointmentmaster</x:String>
				                        <x:String>entitlement</x:String>
				                        <x:String>incident</x:String>
				                        <x:String>systemuser</x:String>
				                        <x:String>task</x:String>
				                        <x:String>opportunity</x:String>
				                        <x:String>campaignresponse</x:String>
				                        <x:String>fax</x:String>
				                        <x:String>bookableresource</x:String>
				                        <x:String>letter</x:String>
				                        <x:String>knowledgearticle</x:String>
				                        <x:String>contact</x:String>
				                        <x:String>campaign</x:String>
				                        <x:String>salesliterature</x:String>
				                        <x:String>lead</x:String>
				                        <x:String>appointment</x:String>
				                        <x:String>bookableresourcecharacteristic</x:String>
				                        <x:String>competitor</x:String>
				                        <x:String>bookableresourcebookingheader</x:String>
				                        <x:String>socialactivity</x:String>
				                        <x:String>list</x:String>
				                        <x:String>bookableresourcebooking</x:String>
				                        <x:String>invoice</x:String>
				                        <x:String>salesorder</x:String>
				                        <x:String>kyc_kycpp</x:String>
				                        <x:String>test_dummy</x:String>
				                        <x:String>phonecall</x:String>
				                      </sco:Collection>
				                      <x:String x:Key=""Target#InArgumentRequired"">Target</x:String>
				                    </mcwa:InvokeSdkMessageActivity.Properties>
				                  </mcwa:InvokeSdkMessageActivity>
				                </sco:Collection>
				              </mxswa:ActivityReference.Properties>
				            </mxswa:ActivityReference>
				          </mxswa:Workflow>
				        </Activity>";

						#endregion Create XAML

						#region Create Workflow

						Entity workflowEntity = CRMData.GetEntityRecordByGuid(service, "workflow", new Guid(workflowIdToGenerateDocument), "xaml");
						if (workflowEntity != null)
						{
							tracingService.Trace("workflowEntity != null");
							//Replacing the document template in main XAML of the workflow with the documentId fetched from the Mapping entity
							string modifiedXaml = mainXaml.Replace("{0}", documentTemplateId);

							//Updating the workflow with the modified XAML
							Entity updateWorkflow = new Entity("workflow");
							updateWorkflow.Id = workflowEntity.Id;
							updateWorkflow.Attributes.Add("xaml", modifiedXaml);
							service.Update(updateWorkflow);
							tracingService.Trace("Workflow updated successfully");

							// Activate the workflow.
							CRMData.ActivateDeactivateRecord(workflowEntity, 1, 2, service);
							tracingService.Trace("Workflow activated successfully");

							//Create ExecuteWorkflowRequest object 
							ExecuteWorkflowRequest request = new ExecuteWorkflowRequest()
							{
								WorkflowId = workflowEntity.Id,
								EntityId = contract.Id
							};
							Console.Write("Created ExecuteWorkflow request, ");

							// Execute the workflow.
							ExecuteWorkflowResponse response = (ExecuteWorkflowResponse)service.Execute(request);
							tracingService.Trace("Workflow executed successfully");

							// De-activate the workflow.
							CRMData.ActivateDeactivateRecord(workflowEntity, 0, 1, service);
							tracingService.Trace("Workflow draft successfully");

							//Updating the documentTemplateId in the XAML to blank
							documentTemplateId = "00000000-0000-0000-0000-000000000000";
							modifiedXaml = mainXaml.Replace("{0}", documentTemplateId);

							//Updating the Workflow XAML with the blank documentTemplateId
							updateWorkflow = new Entity("workflow");
							updateWorkflow.Id = workflowEntity.Id;
							updateWorkflow.Attributes.Add("xaml", modifiedXaml);
							service.Update(updateWorkflow);
							tracingService.Trace("Workflow updated with the blank documentTemplateId successfully");
							#endregion Create Workflow
						}
					}
				}
			}
		}
	}
}

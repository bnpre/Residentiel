////-----------------------------
//// <copyright file="PreLeadCreate.cs" company="BNPRE">
//// Copyright (c) BNPRE. All rights reserved.
//// </copyright>
//// <summary>
//// This file contains operations for Lead class.
//// </summary>
////-------------------------------------------------------------------------------------------------------------------------------
namespace BNP_Plugins
{
	using BNP_Model.Utils;
	using Microsoft.Crm.Sdk.Messages;
	using Microsoft.Xrm.Sdk;
	using Microsoft.Xrm.Sdk.Query;
	using System;
	using System.ServiceModel;

	public class PreOpportunityCreate : IPlugin
	{
		#region Public Methods
		/// <summary>
		/// Executes plug-in code in response to an event.
		/// </summary>
		/// <param name="serviceProvider">Type: Returns_IServiceProvider. A container for service objects. Contains references to the plug-in execution context (<see cref="T:Microsoft.Xrm.Sdk.IPluginExecutionContext"></see>), tracing service (<see cref="T:Microsoft.Xrm.Sdk.ITracingService"></see>), organization service (<see cref="T:Microsoft.Xrm.Sdk.IOrganizationServiceFactory"></see>), and notification service (<see cref="T:Microsoft.Xrm.Sdk.IServiceEndpointNotificationService"></see>).</param>
		/// <exception cref="InvalidPluginExecutionException">
		/// Invalid 
		/// or
		/// An error occurred in the PreOpportunityCreate plug-in. + ex.Message
		/// </exception>
		public void Execute(IServiceProvider serviceProvider)
		{
			// Extract the tracing service for use in debugging sandboxed plug-ins.
			ITracingService iTracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

			// Obtain the execution context from the service provider.
			IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

			// Obtain the organization service reference.
			IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
			IOrganizationService iOrganizationService = serviceFactory.CreateOrganizationService(null);

			// The InputParameters collection contains all the data passed in the message request.
			if (CRMData.ValidateTargetAsEntity("opportunity", context))
			{
				try
				{
					iTracingService.Trace("Context contains target entity");
					// Obtain the target entity from the input parameters.
					Entity opportunity = (Entity)context.InputParameters["Target"];

					iTracingService.Trace("Context fetched");
					GetAndSetBPFStageName(opportunity, null, iOrganizationService, iTracingService, context);
				}
				catch (FaultException<OrganizationServiceFault> ex)
				{
					throw new InvalidPluginExecutionException("An error occurred in the PreOpportunityCreate plug-in." + ex.Message);
				}
			}
		}
		#endregion

		#region Internal Methods
		#region GetAndSetBPFStageName
		/// <summary>
		/// Set the BPF Stagename in the "Stepname" of the entity
		/// </summary>
		/// <param name="entity"></param>
		/// <param name="entityProcessId"></param>
		/// <param name="iOrganizationService"></param>
		/// <param name="iTracingService"></param>
		internal static void GetAndSetBPFStageName(Entity entity,
																							 Entity entityImage,
																							 IOrganizationService iOrganizationService,
																							 ITracingService iTracingService,
																							 IPluginExecutionContext context)
		{
			if (entity != null &&
					entity.Attributes.Contains("stageid"))
			{
				iTracingService.Trace("Entered into method - GetBPFStageName");
				string entityStageName = string.Empty;
				Guid entityProcessId = new Guid();
				Guid entityStageId = new Guid();

				//Get the ProcessId and StageId of the current entity record.
				entityProcessId = CRMData.GetAttributeValue<Guid>(entity, entityImage, "processid");
				entityStageId = CRMData.GetAttributeValue<Guid>(entity, entityImage, "stageid");

				iTracingService.Trace("entityProcessId = {0}", entityProcessId.ToString());
				iTracingService.Trace("entityStageId = {0}", entityStageId.ToString());

				if (entityProcessId != Guid.Empty &&
						entityStageId != Guid.Empty)
				{
					iTracingService.Trace("opportunityProcessId is not null");
					iTracingService.Trace("opportunityProcessId = {0}", entityProcessId.ToString());

					// Query Process Stage by passing ProcessId and get the name of the stage.
					QueryExpression processStageQueryExpression = new QueryExpression("processstage");
					processStageQueryExpression.ColumnSet = new ColumnSet(true);
					processStageQueryExpression.Criteria.AddCondition("primaryentitytypecode", ConditionOperator.Equal, entity.LogicalName);
					processStageQueryExpression.Criteria.AddCondition("processid", ConditionOperator.Equal, entityProcessId);
					processStageQueryExpression.Criteria.AddCondition("processstageid", ConditionOperator.Equal, entityStageId);
					EntityCollection processStageEntityCollection = iOrganizationService.RetrieveMultiple(processStageQueryExpression);

					iTracingService.Trace("retrieved Programme = {0} records", processStageEntityCollection.Entities.Count);

					if (processStageEntityCollection.Entities.Count > 0)
					{
						entityStageName = CRMData.GetAttributeValue<string>(processStageEntityCollection.Entities[0], null, "stagename");
						if (string.IsNullOrWhiteSpace(entityStageName) == false)
						{
							iTracingService.Trace("entityStageName is not null");
							iTracingService.Trace("entityStageName = {0}", entityStageName);
							switch (entityStageName)
							{
								case "Nouvelle Demande":
									CRMData.AddAttribute<OptionSetValue>(entity, "statecode", new OptionSetValue(0));
									CRMData.AddAttribute<OptionSetValue>(entity, "statuscode", new OptionSetValue(1));
									break;
								case "En cours de Traitement":
									CRMData.AddAttribute<OptionSetValue>(entity, "statecode", new OptionSetValue(0));
									CRMData.AddAttribute<OptionSetValue>(entity, "statuscode", new OptionSetValue(2));
									break;
								case "Qualification du besoin":
									CRMData.AddAttribute<OptionSetValue>(entity, "statecode", new OptionSetValue(0));
									CRMData.AddAttribute<OptionSetValue>(entity, "statuscode", new OptionSetValue(852810000));
									break;
								case "Négociation":
									CRMData.AddAttribute<OptionSetValue>(entity, "statecode", new OptionSetValue(0));
									CRMData.AddAttribute<OptionSetValue>(entity, "statuscode", new OptionSetValue(852810001));
									break;
								case "Contractualisation":
									//CRMData.AddAttribute<OptionSetValue>(entity, "statecode", new OptionSetValue(1));
									//CRMData.AddAttribute<OptionSetValue>(entity, "statuscode", new OptionSetValue(3));
									context.SharedVariables.Add("isOpportunityClose", true);
									break;
								default:
									break;
							}
							iTracingService.Trace("Statecode and StatusCode updated in Context");
							iTracingService.Trace("Execution completed for method GetBPFStageName");
						}
					}
				}
			}
		}
		#endregion
		#endregion
	}

	public class PostOpportunityCreate : IPlugin
	{
		#region Public Methods
		/// <summary>
		/// Executes plug-in code in response to an event.
		/// </summary>
		/// <param name="serviceProvider">Type: Returns_IServiceProvider. A container for service objects. Contains references to the plug-in execution context (<see cref="T:Microsoft.Xrm.Sdk.IPluginExecutionContext"></see>), tracing service (<see cref="T:Microsoft.Xrm.Sdk.ITracingService"></see>), organization service (<see cref="T:Microsoft.Xrm.Sdk.IOrganizationServiceFactory"></see>), and notification service (<see cref="T:Microsoft.Xrm.Sdk.IServiceEndpointNotificationService"></see>).</param>
		/// <exception cref="InvalidPluginExecutionException">
		/// Invalid 
		/// or
		/// An error occurred in the PostOpportunityCreate plug-in. + ex.Message
		/// </exception>
		public void Execute(IServiceProvider serviceProvider)
		{
			// Extract the tracing service for use in debugging sandboxed plug-ins.
			ITracingService iTracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

			// Obtain the execution context from the service provider.
			IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

			// Obtain the organization service reference.
			IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
			IOrganizationService iOrganizationService = serviceFactory.CreateOrganizationService(null);

			// The InputParameters collection contains all the data passed in the message request.
			if (CRMData.ValidateTargetAsEntity("opportunity", context))
			{
				try
				{
					bool isOpportunityClose = false;
					iTracingService.Trace("Context contains target entity");
					// Obtain the target entity from the input parameters.
					Entity opportunity = (Entity)context.InputParameters["Target"];

					if (context.SharedVariables.ContainsKey("isContactExists"))
					{
						iTracingService.Trace("context contains the key [isOpportunityClose]");
						isOpportunityClose = (bool)context.SharedVariables["isOpportunityClose"];
						iTracingService.Trace("isOpportunityClose ={0}", isOpportunityClose.ToString());
					}

					if (isOpportunityClose == true)
					{
						CloseOpportunityAsWon(opportunity, iOrganizationService, iTracingService, context);
					}
				}
				catch (FaultException<OrganizationServiceFault> ex)
				{
					throw new InvalidPluginExecutionException("An error occurred in the PostOpportunityCreate plug-in." + ex.Message);
				}
			}
		}
		#endregion

		#region Internal Methods
		#region CloseOpportunityAsWon
		/// <summary>
		/// Close the Opportunity as Won
		/// </summary>
		/// <param name="entity"></param>
		/// <param name="iOrganizationService"></param>
		/// <param name="iTracingService"></param>
		/// <param name="context"></param>
		internal static void CloseOpportunityAsWon(Entity entity,
																							 IOrganizationService iOrganizationService,
																							 ITracingService iTracingService,
																							 IPluginExecutionContext context)
		{
			if (entity != null)
			{
				WinOpportunityRequest request = new WinOpportunityRequest();
				iTracingService.Trace("WinOpportunityRequest request created");
				Entity opportunityClose = new Entity("opportunityclose");
				iTracingService.Trace("opportunityclose entity object created");
				opportunityClose.Attributes.Add("opportunityid", new EntityReference("opportunity", entity.Id));
				iTracingService.Trace("opportunityclose attribute added opportunity Id");
				request.OpportunityClose = opportunityClose;
				iTracingService.Trace("opportunity closed");
				request.RequestName = "WinOpportunity";
				iTracingService.Trace("request.RequestName");
				OptionSetValue opportunityStatusCode = new OptionSetValue();
				iTracingService.Trace("opportunityStatusCode");
				opportunityStatusCode.Value = 3;
				request.Status = opportunityStatusCode;
				iTracingService.Trace("status changed");
				WinOpportunityResponse response = (WinOpportunityResponse)iOrganizationService.Execute(request);
			}
		}
		#endregion
		#endregion
	}

	public class PreOpportunityUpdate : IPlugin
	{
		#region Public Methods
		/// <summary>
		/// Executes plug-in code in response to an event.
		/// </summary>
		/// <param name="serviceProvider">Type: Returns_IServiceProvider. A container for service objects. Contains references to the plug-in execution context (<see cref="T:Microsoft.Xrm.Sdk.IPluginExecutionContext"></see>), tracing service (<see cref="T:Microsoft.Xrm.Sdk.ITracingService"></see>), organization service (<see cref="T:Microsoft.Xrm.Sdk.IOrganizationServiceFactory"></see>), and notification service (<see cref="T:Microsoft.Xrm.Sdk.IServiceEndpointNotificationService"></see>).</param>
		/// <exception cref="InvalidPluginExecutionException">
		/// Invalid 
		/// or
		/// An error occurred in the PreOpportunityUpdate plug-in. + ex.Message
		/// </exception>
		public void Execute(IServiceProvider serviceProvider)
		{
			// Extract the tracing service for use in debugging sandboxed plug-ins.
			ITracingService iTracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

			// Obtain the execution context from the service provider.
			IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

			// Obtain the organization service reference.
			IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
			IOrganizationService iOrganizationService = serviceFactory.CreateOrganizationService(null);

			// The InputParameters collection contains all the data passed in the message request.
			if (CRMData.ValidateTargetAsEntity("opportunity", context))
			{
				try
				{
					iTracingService.Trace("Context contains target entity");
					// Obtain the target entity from the input parameters.
					Entity opportunity = (Entity)context.InputParameters["Target"];

					//Retrieving processId and StageId as the they can not be obtained from the image at the time of update.
					Entity opportunityRetrieved = iOrganizationService.Retrieve(opportunity.LogicalName, opportunity.Id, new ColumnSet("processid", "stageid"));

					iTracingService.Trace("opportunityRetrieved fetched");
					PreOpportunityCreate.GetAndSetBPFStageName(opportunity, opportunityRetrieved, iOrganizationService, iTracingService, context);
				}
				catch (FaultException<OrganizationServiceFault> ex)
				{
					throw new InvalidPluginExecutionException("An error occurred in the PreOpportunityUpdate plug-in." + ex.Message);
				}
			}
		}
		#endregion
	}

	public class PostOpportunityUpdate : IPlugin
	{
		#region Public Methods
		/// <summary>
		/// Executes plug-in code in response to an event.
		/// </summary>
		/// <param name="serviceProvider">Type: Returns_IServiceProvider. A container for service objects. Contains references to the plug-in execution context (<see cref="T:Microsoft.Xrm.Sdk.IPluginExecutionContext"></see>), tracing service (<see cref="T:Microsoft.Xrm.Sdk.ITracingService"></see>), organization service (<see cref="T:Microsoft.Xrm.Sdk.IOrganizationServiceFactory"></see>), and notification service (<see cref="T:Microsoft.Xrm.Sdk.IServiceEndpointNotificationService"></see>).</param>
		/// <exception cref="InvalidPluginExecutionException">
		/// Invalid 
		/// or
		/// An error occurred in the PostOpportunityUpdate plug-in. + ex.Message
		/// </exception>
		public void Execute(IServiceProvider serviceProvider)
		{
			// Extract the tracing service for use in debugging sandboxed plug-ins.
			ITracingService iTracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

			// Obtain the execution context from the service provider.
			IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

			// Obtain the organization service reference.
			IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
			IOrganizationService iOrganizationService = serviceFactory.CreateOrganizationService(null);

			// The InputParameters collection contains all the data passed in the message request.
			if (CRMData.ValidateTargetAsEntity("opportunity", context))
			{
				try
				{
					bool isOpportunityClose = false;
					iTracingService.Trace("Context contains target entity");
					// Obtain the target entity from the input parameters.
					Entity opportunity = (Entity)context.InputParameters["Target"];

					if (context.SharedVariables.ContainsKey("isOpportunityClose"))
					{
						iTracingService.Trace("context contains the key [isOpportunityClose]");
						isOpportunityClose = (bool)context.SharedVariables["isOpportunityClose"];
						iTracingService.Trace("isOpportunityClose ={0}", isOpportunityClose.ToString());
					}

					if (isOpportunityClose == true)
					{
						PostOpportunityCreate.CloseOpportunityAsWon(opportunity, iOrganizationService, iTracingService, context);
					}
				}
				catch (FaultException<OrganizationServiceFault> ex)
				{
					throw new InvalidPluginExecutionException("An error occurred in the PostOpportunityUpdate plug-in." + ex.Message);
				}
			}
		}
		#endregion
	}
}


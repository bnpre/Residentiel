﻿////-----------------------------
//// <copyright file="resi_Contrat.cs" company="BNPRE">
//// Copyright (c) BNPRE. All rights reserved.
//// </copyright>
//// <summary>
//// This file contains operations for Lead class.
//// </summary>
////-------------------------------------------------------------------------------------------------------------------------------
namespace BNP_Plugins
{
	using BNP_Model.Utils;
	using Microsoft.Xrm.Sdk;
	using Microsoft.Xrm.Sdk.Client;
	using Microsoft.Xrm.Sdk.Messages;
	using Microsoft.Xrm.Sdk.Query;
	using System;

	/// <summary>
	/// Contrat Class
	/// </summary>
	/// <seealso cref="Microsoft.Xrm.Sdk.IPlugin" />
	public class PostContratCreate : IPlugin
	{
		/// <summary>
		/// Executes plug-in code in response to an event.
		/// </summary>
		/// <param name="serviceProvider">Type: Returns_IServiceProvider. A container for service objects. Contains references to the plug-in execution context (<see cref="T:Microsoft.Xrm.Sdk.IPluginExecutionContext"></see>), tracing service (<see cref="T:Microsoft.Xrm.Sdk.ITracingService"></see>), organization service (<see cref="T:Microsoft.Xrm.Sdk.IOrganizationServiceFactory"></see>), and notification service (<see cref="T:Microsoft.Xrm.Sdk.IServiceEndpointNotificationService"></see>).</param>
		/// <exception cref="InvalidPluginExecutionException">
		/// Invalid 
		/// or
		/// An error occurred in the PostContratCreate plug-in. + ex.Message
		/// </exception>

		#region Public Methods
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
			if (CRMData.ValidateTargetAsEntity("resi_contrat", context))
			{
				try
				{
					iTracingService.Trace("Context contains target entity");
					// Obtain the target entity from the input parameters.
					Entity contrat = (Entity)context.InputParameters["Target"];
					iTracingService.Trace("Context fetched.");

					#region Fetching Programme Attribute from Context.
					EntityReference programmeOnContrat = CRMData.GetAttributeValue<EntityReference>(contrat, null, "resi_concernerid");
					iTracingService.Trace("Attributes fetched from Context.");
					#endregion

					if (programmeOnContrat != null)
					{
						iTracingService.Trace("Programme entity reference fetched from Contrat");

						#region Fetching "Echéancier" entity records related to the "Programme"
						iTracingService.Trace("Fetching 'Echéancier' entity records related to the 'Programme'");
						// Fetch "Echéancier" entity records related to the 'Programme'
						QueryExpression echeancierQueryExpression = new QueryExpression("resi_echeancier");
						echeancierQueryExpression.ColumnSet = new ColumnSet(true);
						echeancierQueryExpression.Criteria.AddCondition("resi_programmeid", ConditionOperator.Equal, programmeOnContrat.Id);
						EntityCollection echeancierEntityCollection = iOrganizationService.RetrieveMultiple(echeancierQueryExpression);
						#endregion

						if (echeancierEntityCollection.Entities.Count > 0)
						{
							iTracingService.Trace("retrieved Echéancier = {0} records", echeancierEntityCollection.Entities.Count);
							
							#region Initialize execute multiple request
							//Update all Items.
							ExecuteMultipleRequest executeMultipleRequest = null;

							// Create an ExecuteMultipleRequest object.
							executeMultipleRequest = new ExecuteMultipleRequest()
							{
								// Assign settings that define execution behavior: continue on error, return responses. 
								Settings = new ExecuteMultipleSettings()
								{
									ContinueOnError = true,
									ReturnResponses = false
								},
								// Create an empty organization request collection.
								Requests = new OrganizationRequestCollection()
							};
							#endregion

							foreach (Entity echeancier in echeancierEntityCollection.Entities)
							{
								//Fetch field values from current Echéancier entity
								string nom = CRMData.GetAttributeValue<string>(echeancier, null, "resi_declencheur");
								int ordre = CRMData.GetAttributeValue<int>(echeancier, null, "resi_ordre");
								DateTime? dateDEcheance = CRMData.GetAttributeValue<DateTime?>(echeancier, null, "resi_datedecheance");
								decimal initial = CRMData.GetAttributeValue<decimal>(echeancier, null, "resi_initial");
								iTracingService.Trace("nom:{0}, ordre:{1}, dateDEcheance:{2}, initial:{3}", nom, ordre, dateDEcheance, initial);
								//Create a new Echéancier entity record based on the values in the fields of current Echéancier entity
								Entity echeancierEntity = new Entity("resi_echeancier");
								CRMData.AddAttribute<bool>(echeancierEntity, "resi_isecheancierinitial", false); //echeancierInitial
								CRMData.AddAttribute<EntityReference>(echeancierEntity, "resi_contratid", new EntityReference(contrat.LogicalName, contrat.Id));
								CRMData.AddAttribute<string>(echeancierEntity, "resi_declencheur", nom);
								CRMData.AddAttribute<int>(echeancierEntity, "resi_ordre", ordre);
								CRMData.AddAttribute<DateTime?>(echeancierEntity, "resi_datedecheance", dateDEcheance);
								CRMData.AddAttribute<decimal>(echeancierEntity, "resi_initial", initial);
								iOrganizationService.Create(echeancierEntity);

								iTracingService.Trace("Added the CreateRequest to ExecuteMultipleRequest object");
								CreateRequest createRequest = new CreateRequest { Target = echeancierEntity };
								executeMultipleRequest.Requests.Add(createRequest);
							}
							// Execute all the requests in the request collection using a single web method call.
							ExecuteMultipleResponse responseWithResults =
									(ExecuteMultipleResponse)iOrganizationService.Execute(executeMultipleRequest);
							iTracingService.Trace("ExecuteMultipleRequest executed successfully");
						}
					}
				}
				catch (Exception ex)
				{
					throw ex;
				}
			}
			iTracingService.Trace("Plugin execution completed");
		}
		#endregion
	}
}

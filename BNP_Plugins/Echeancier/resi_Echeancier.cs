////-----------------------------
//// <copyright file="resi_Echeancier.cs" company="BNPRE">
//// Copyright (c) BNPRE. All rights reserved.
//// </copyright>
//// <summary>
//// This file contains operations for Echeancier class.
//// </summary>
////-------------------------------------------------------------------------------------------------------------------------------
namespace BNP_Plugins
{
	using BNP_Model.Utils;
	using Microsoft.Xrm.Sdk;
	using Microsoft.Xrm.Sdk.Query;
	using System;

	/// <summary>
	/// Echeancier class
	/// </summary>
	/// <seealso cref="Microsoft.Xrm.Sdk.IPlugin" />
	public class PreEcheancierCreate : IPlugin
	{
		/// <summary>
		/// Executes plug-in code in response to an event.
		/// </summary>
		/// <param name="serviceProvider">Type: Returns_IServiceProvider. A container for service objects. Contains references to the plug-in execution context (<see cref="T:Microsoft.Xrm.Sdk.IPluginExecutionContext"></see>), tracing service (<see cref="T:Microsoft.Xrm.Sdk.ITracingService"></see>), organization service (<see cref="T:Microsoft.Xrm.Sdk.IOrganizationServiceFactory"></see>), and notification service (<see cref="T:Microsoft.Xrm.Sdk.IServiceEndpointNotificationService"></see>).</param>
		/// <exception cref="InvalidPluginExecutionException">
		/// Invalid 
		/// or
		/// An error occurred in the PreEcheancierCreate plug-in. + ex.Message
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
			if (CRMData.ValidateTargetAsEntity("resi_echeancier", context))
			{
				try
				{
					iTracingService.Trace("Context contains target entity");
					// Obtain the target entity from the input parameters.
					Entity echeancier = (Entity)context.InputParameters["Target"];
					iTracingService.Trace("Context fetched.");
					if (echeancier.Contains("resi_programmeid"))
					{
						iTracingService.Trace("echeancier.Contains('resi_programmeid')");
						EntityReference programme = CRMData.GetAttributeValue<EntityReference>(echeancier, null, "resi_programmeid");
						int ordre = 1;
						if (programme != null)
						{
							iTracingService.Trace("programme != null");
							#region
							// Fetch "echeancier" entity records related to the current programme
							QueryExpression echeancierQueryExpression = new QueryExpression("resi_echeancier");
							echeancierQueryExpression.ColumnSet = new ColumnSet("resi_ordre");
							echeancierQueryExpression.AddOrder("createdon", OrderType.Descending);
							echeancierQueryExpression.Criteria.AddCondition("resi_programmeid", ConditionOperator.Equal, programme.Id);
							EntityCollection echeancierEntityCollection = iOrganizationService.RetrieveMultiple(echeancierQueryExpression);
							#endregion

							iTracingService.Trace("retrieved Echeancier = {0} records", echeancierEntityCollection.Entities.Count);
							if (echeancierEntityCollection.Entities.Count > 0)
							{
								iTracingService.Trace("echeancierEntityCollection.Entities.Count > 0");
								ordre = CRMData.GetAttributeValue<int>(echeancierEntityCollection.Entities[0], null, "resi_ordre");
								if (ordre > 0)
								{
									iTracingService.Trace("ordre > 0");
									ordre = ordre + 1;
									CRMData.AddAttribute(echeancier, "resi_ordre", ordre);
								}
							}
							iTracingService.Trace("echeancierEntityCollection.Entities.Count <= 0");
							CRMData.AddAttribute(echeancier, "resi_ordre", ordre);
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


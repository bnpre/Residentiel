////-----------------------------
//// <copyright file="resi_SalesLiterature.cs" company="BNPRE">
//// Copyright (c) BNPRE. All rights reserved.
//// </copyright>
//// <summary>
//// This file contains operations for SalesLiterature class.
//// </summary>
////-------------------------------------------------------------------------------------------------------------------------------
namespace BNP_Plugins
{
	using BNP_Model.Utils;
	using Microsoft.Xrm.Sdk;
	using Microsoft.Xrm.Sdk.Messages;
	using System;

	/// <summary>
	/// SalesLiterature Class
	/// </summary>
	/// <seealso cref="Microsoft.Xrm.Sdk.IPlugin" />
	public class PostSalesLiteratureCreate : IPlugin
	{
		/// <summary>
		/// Executes plug-in code in response to an event.
		/// </summary>
		/// <param name="serviceProvider">Type: Returns_IServiceProvider. A container for service objects. Contains references to the plug-in execution context (<see cref="T:Microsoft.Xrm.Sdk.IPluginExecutionContext"></see>), tracing service (<see cref="T:Microsoft.Xrm.Sdk.ITracingService"></see>), organization service (<see cref="T:Microsoft.Xrm.Sdk.IOrganizationServiceFactory"></see>), and notification service (<see cref="T:Microsoft.Xrm.Sdk.IServiceEndpointNotificationService"></see>).</param>
		/// <exception cref="InvalidPluginExecutionException">
		/// Invalid 
		/// or
		/// An error occurred in the PostSalesLiteratureCreate plug-in. + ex.Message
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
			if (CRMData.ValidateTargetAsEntity("salesliterature", context))
			{
				try
				{
					iTracingService.Trace("Context contains target entity");
					// Obtain the target entity from the input parameters.
					Entity salesLiterature = (Entity)context.InputParameters["Target"];
					iTracingService.Trace("Context fetched.");

					#region Fetching SalesLiterature Attribute from Context.
					string associatedProgrammeIdOnSalesLiterature = CRMData.GetAttributeValue<string>(salesLiterature, null, "resi_associatedprogrammeid");
					associatedProgrammeIdOnSalesLiterature = associatedProgrammeIdOnSalesLiterature.Replace("{", "").Replace("}","");
					iTracingService.Trace("Attributes fetched from Context.");
					#endregion

					if (associatedProgrammeIdOnSalesLiterature != null)
					{
						iTracingService.Trace("associatedProgrammeIdOnSalesLiterature != null");
						AssociateRequest associateSalesLiteratureProgramme = new AssociateRequest();

						//Target is the entity that you are associating your entities to.
						associateSalesLiteratureProgramme.Target = new EntityReference(salesLiterature.LogicalName, salesLiterature.Id);

						//RelatedEntities are the entities you are associating to your target (can be more than 1)
						associateSalesLiteratureProgramme.RelatedEntities = new EntityReferenceCollection();
						associateSalesLiteratureProgramme.RelatedEntities.Add(new EntityReference("resi_programme", new Guid(associatedProgrammeIdOnSalesLiterature)));

						//The relationship schema name in CRM you are using to associate the entities. 
						//found in settings - customization - entity - relationships
						associateSalesLiteratureProgramme.Relationship = new Relationship("resi_resi_programme_salesliterature");

						//execute the request
						iOrganizationService.Execute(associateSalesLiteratureProgramme);
						iTracingService.Trace("associateSalesLiteratureProgramme request exexcuted successfully");
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

////-----------------------------
//// <copyright file="resi_Account.cs" company="BNPRE">
//// Copyright (c) BNPRE. All rights reserved.
//// </copyright>
//// <summary>
//// This file contains operations for Account class.
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
	public class PreAccountCreate : IPlugin
	{
		/// <summary>
		/// Executes plug-in code in response to an event.
		/// </summary>
		/// <param name="serviceProvider">Type: Returns_IServiceProvider. A container for service objects. Contains references to the plug-in execution context (<see cref="T:Microsoft.Xrm.Sdk.IPluginExecutionContext"></see>), tracing service (<see cref="T:Microsoft.Xrm.Sdk.ITracingService"></see>), organization service (<see cref="T:Microsoft.Xrm.Sdk.IOrganizationServiceFactory"></see>), and notification service (<see cref="T:Microsoft.Xrm.Sdk.IServiceEndpointNotificationService"></see>).</param>
		/// <exception cref="InvalidPluginExecutionException">
		/// Invalid 
		/// or
		/// An error occurred in the PreAccountCreate plug-in. + ex.Message
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
			if (CRMData.ValidateTargetAsEntity("account", context))
			{
				try
				{
					iTracingService.Trace("Context contains target entity");
					// Obtain the target entity from the input parameters.
					Entity account = (Entity)context.InputParameters["Target"];
					iTracingService.Trace("Context fetched.");
					UpdateCodeGroupeInAccount(account, iTracingService, iOrganizationService);
				}
				catch (Exception ex)
				{
					throw ex;
				}
			}
			iTracingService.Trace("Plugin execution completed");
		}
		#endregion

		#region Private Methods
		internal static void UpdateCodeGroupeInAccount(Entity entity, ITracingService iTracingService, IOrganizationService iOrganizationService)
		{
			try
			{
				Entity parentAccount;
				string accountNumber = string.Empty;
				if (entity.Contains("parentaccountid"))
				{
					#region Fetching "Agence mère" Attribute from Context.
					EntityReference parentAccountOnAccount = CRMData.GetAttributeValue<EntityReference>(entity, null, "parentaccountid");
					iTracingService.Trace("Attributes fetched from entity Context.");
					#endregion

					if (parentAccountOnAccount != null)
					{
						iTracingService.Trace("parentAccount entity reference fetched from Account");

						#region Fetching "accountnumber" from the "parentAccount" obtained
						parentAccount = iOrganizationService.Retrieve("account", parentAccountOnAccount.Id, new ColumnSet("accountnumber"));
						#endregion

						if (parentAccount != null)
						{
							iTracingService.Trace("parentAccount != null");
							accountNumber = CRMData.GetAttributeValue<string>(parentAccount, null, "accountnumber");

							if (string.IsNullOrWhiteSpace(accountNumber) == false)
							{
								CRMData.AddAttribute<string>(entity, "resi_codegroupe", accountNumber);
								iTracingService.Trace("resi_codegroupe field value successfully updated in the context");
							}
							else
							{
								CRMData.AddAttribute<string>(entity, "resi_codegroupe", accountNumber);
								iTracingService.Trace("resi_codegroupe field value successfully updated in the context");
							}
						}
					}
					else
					{
						CRMData.AddAttribute<string>(entity, "resi_codegroupe", accountNumber);
						iTracingService.Trace("resi_codegroupe field value successfully updated in the context");
					}
				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
		#endregion
	}

	/// <summary>
	/// Contrat Class
	/// </summary>
	/// <seealso cref="Microsoft.Xrm.Sdk.IPlugin" />
	public class PreAccountUpdate : IPlugin
	{
		/// <summary>
		/// Executes plug-in code in response to an event.
		/// </summary>
		/// <param name="serviceProvider">Type: Returns_IServiceProvider. A container for service objects. Contains references to the plug-in execution context (<see cref="T:Microsoft.Xrm.Sdk.IPluginExecutionContext"></see>), tracing service (<see cref="T:Microsoft.Xrm.Sdk.ITracingService"></see>), organization service (<see cref="T:Microsoft.Xrm.Sdk.IOrganizationServiceFactory"></see>), and notification service (<see cref="T:Microsoft.Xrm.Sdk.IServiceEndpointNotificationService"></see>).</param>
		/// <exception cref="InvalidPluginExecutionException">
		/// Invalid 
		/// or
		/// An error occurred in the PreAccountUpdate plug-in. + ex.Message
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
			if (CRMData.ValidateTargetAsEntity("account", context))
			{
				try
				{
					iTracingService.Trace("Context contains target entity");
					// Obtain the target entity from the input parameters.
					Entity account = (Entity)context.InputParameters["Target"];
					iTracingService.Trace("Context fetched.");
					PreAccountCreate.UpdateCodeGroupeInAccount(account, iTracingService, iOrganizationService);
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


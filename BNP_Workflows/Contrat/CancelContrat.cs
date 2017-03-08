using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace BNP_Workflows
{
	public class CancelContrat : CodeActivity
	{
		//private string webserviceUrl = null, password = null, userName = null;

		/// <summary>
		/// Executes the specified execution context.
		/// </summary>
		/// <param name="executionContext">The execution context.</param>
		/// <exception cref="InvalidPluginExecutionException">The timeout elapsed while attempting to issue the request.</exception>
		protected override void Execute(CodeActivityContext context)
		{
			IWorkflowContext wfContext = context.GetExtension<IWorkflowContext>();
			IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
			IOrganizationService service = serviceFactory.CreateOrganizationService(null);

			try
			{
				if (wfContext == null)
				{
					throw new Exception("Workflow context est vide");
				}

				String contratName = wfContext.PrimaryEntityName;
				Guid contratId = wfContext.PrimaryEntityId;

				#region Deactivate the Contrat
				//StateCode = 1 and StatusCode = 2 for deactivating Contrat
				SetStateRequest setStateRequest = new SetStateRequest()
				{
					EntityMoniker = new EntityReference
					{
						Id = contratId,
						LogicalName = contratName,
					},
					State = new OptionSetValue(1),
					Status = new OptionSetValue(852810000)
				};
				service.Execute(setStateRequest);
				#endregion

				//#region Code to fetch the Web Service Url from the "Parametre" entity
				//if (string.IsNullOrEmpty(this.webserviceUrl))
				//{
				//	ConditionExpression condition1 = new ConditionExpression();
				//	condition1.AttributeName = "resi_cle";
				//	condition1.Operator = ConditionOperator.Equal;
				//	condition1.Values.Add("urlGetCommandesDetailSC");
				//	ConditionExpression condition2 = new ConditionExpression();
				//	condition2.AttributeName = "resi_cle";
				//	condition2.Operator = ConditionOperator.Equal;
				//	condition2.Values.Add("passwordRelai");
				//	ConditionExpression condition3 = new ConditionExpression();
				//	condition3.AttributeName = "resi_cle";
				//	condition3.Operator = ConditionOperator.Equal;
				//	condition3.Values.Add("userRelai");
				//	FilterExpression filter1 = new FilterExpression(LogicalOperator.Or);
				//	filter1.Conditions.AddRange(condition1, condition2, condition3);
				//	QueryExpression qe = new QueryExpression("resi_parametre");
				//	qe.NoLock = true;
				//	qe.ColumnSet.AddColumns("resi_cle", "resi_valeur");
				//	qe.Criteria.AddFilter(filter1);
				//	EntityCollection colSettings = service.RetrieveMultiple(qe);

				//	foreach (var setting in colSettings.Entities)
				//	{
				//		if (setting.GetAttributeValue<string>("resi_cle") == "urlGetCommandesDetailSC")
				//		{
				//			this.webserviceUrl = setting.GetAttributeValue<string>("resi_valeur");
				//		}

				//		if (setting.GetAttributeValue<string>("resi_cle") == "userRelai")
				//		{
				//			this.userName = setting.GetAttributeValue<string>("resi_valeur");
				//		}

				//		if (setting.GetAttributeValue<string>("resi_cle") == "passwordRelai")
				//		{
				//			this.password = setting.GetAttributeValue<string>("resi_valeur");
				//		}

				//		if (string.IsNullOrEmpty("webserviceUrl") || string.IsNullOrEmpty("userName") || string.IsNullOrEmpty("password"))
				//		{
				//			throw new InvalidPluginExecutionException("Webservice url, user name or password not defined in settings entity");
				//		}
				//	}

				//	try
				//	{
						
				//	}

				//	catch (FaultException<OrganizationServiceFault> ex)
				//	{
				//		throw new InvalidPluginExecutionException(ex.Message);

				//	}
				//}
				//#endregion
			}
			catch (Exception ex)
			{
				throw new Exception(ex.Message);
			}
		}
	}
}

using BNP_Model.Utils;
using BNP_Workflows.Model.Definitions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace BNP_Workflows
{
	public class GenerateContractXML : CodeActivity
	{
		private string webserviceUrl = string.Empty;
		private string password = string.Empty;
		private string userName = string.Empty;

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

			try
			{
				if (wfContext == null)
				{
					throw new Exception("Workflow context est vide");
				}

				#region Fetch the Webservice credentials from the "Parametre" Entity
				ConditionExpression condition1 = new ConditionExpression();
				condition1.AttributeName = "core_name";
				condition1.Operator = ConditionOperator.Equal;
				condition1.Values.Add("URL-ENVOILETTRES");
				ConditionExpression condition2 = new ConditionExpression();
				condition2.AttributeName = "core_name";
				condition2.Operator = ConditionOperator.Equal;
				condition2.Values.Add("passwordRelai");
				ConditionExpression condition3 = new ConditionExpression();
				condition3.AttributeName = "core_name";
				condition3.Operator = ConditionOperator.Equal;
				condition3.Values.Add("userRelai");
				FilterExpression filter1 = new FilterExpression(LogicalOperator.Or);
				filter1.Conditions.AddRange(condition1, condition2, condition3);
				QueryExpression qe = new QueryExpression("core_parametre");
				qe.ColumnSet.AddColumns("core_name", "core_valeur");
				qe.Criteria.AddFilter(filter1);
				EntityCollection colSettings = service.RetrieveMultiple(qe);
				foreach (var setting in colSettings.Entities)
				{
					if (setting.GetAttributeValue<string>("core_name") == "URL-ENVOILETTRES")
					{
						this.webserviceUrl = setting.GetAttributeValue<string>("core_valeur");
					}

					if (setting.GetAttributeValue<string>("core_name") == "userRelai")
					{
						this.userName = setting.GetAttributeValue<string>("core_valeur");
					}

					if (setting.GetAttributeValue<string>("core_name") == "passwordRelai")
					{
						this.password = setting.GetAttributeValue<string>("core_valeur");
					}

					if (string.IsNullOrEmpty(this.webserviceUrl) || string.IsNullOrEmpty(this.userName) || string.IsNullOrEmpty(this.password))
					{
						//throw new InvalidPluginExecutionException("Webservice url, user name or password not defined in settings entity");
					}
				}
				#endregion

				//Get the current Contrat record Id
				Guid contratId = wfContext.PrimaryEntityId;
				tracingService.Trace("Current Contrat Id fetched");
				BNP_Workflows.GestcomWebService.Dossier dossier = new GestcomWebService.Dossier();
				//Contrat Contrat = new Contrat();
				if (contratId != null)
				{
					//Fetch current Contrat entity record 
					Entity contratEntity = CRMData.GetEntityRecordByGuid(service, "resi_contrat", contratId, new string[] { "resi_concernerid", "resi_commissionnerid" });
					if (contratEntity != null)
					{
						tracingService.Trace("Current Contrat Entity fetched");
						//Fetch the Programme related with the Contrat
						EntityReference programme = CRMData.GetAttributeValue<EntityReference>(contratEntity, null, "resi_concernerid");
						if (programme != null)
						{
							tracingService.Trace("programme != null");
							//Retrieve the Programme entity
							Entity programmeEntity = service.Retrieve("resi_programme", programme.Id, new ColumnSet("resi_name", "resi_idgestcom", "resi_idwincom"));
							if (programmeEntity != null)
							{
								tracingService.Trace("programmeEntity != null");
								string programmeName = CRMData.GetAttributeValue<string>(programmeEntity, null, "resi_name");
								string programmeGestcomId = CRMData.GetAttributeValue<string>(programmeEntity, null, "resi_idgestcom");
								string programmeWincomId = CRMData.GetAttributeValue<string>(programmeEntity, null, "resi_idwincom");

								dossier.contrat.Program_name = programmeName;
								dossier.contrat.Programm_id_gestcom = programmeGestcomId;
								dossier.contrat.Program_id_wincom = programmeWincomId;
							}
						}

						#region Retrieving Acquereur (resi_purchaser) based on contrat
						//Retrieving Acquereur (resi_purchaser) based on contrat
						QueryExpression acquereurQueryExpression = new QueryExpression("resi_purchaser");
						acquereurQueryExpression.ColumnSet = new ColumnSet("resi_nom");
						acquereurQueryExpression.Criteria.AddCondition("resi_contratid", ConditionOperator.Equal, contratId);

						EntityCollection acquereurs = service.RetrieveMultiple(acquereurQueryExpression);
						tracingService.Trace("retrieved acquereurs");
						#endregion

						if (acquereurs.Entities.Count > 0)
						{
							tracingService.Trace("acquereurs.Entities.Count > 0");
							foreach (Entity acquereur in acquereurs.Entities)
							{
								string acquereurName = CRMData.GetAttributeValue<string>(acquereur, null, "resi_nom");
								dossier.contrat.acquereur = new GestcomWebService.acquereur();
								dossier.contrat.acquereur.Firstname = acquereurName;
							}
						}

						#region Retrieving Bien (resi_compositionslot) based on contrat
						//Retrieving Bien (resi_compositionslot) based on contrat
						QueryExpression compositionsLotQueryExpression = new QueryExpression("resi_compositionslot");
						compositionsLotQueryExpression.ColumnSet = new ColumnSet("resi_typedebientype");
						compositionsLotQueryExpression.Criteria.AddCondition("resi_contratid", ConditionOperator.Equal, contratId);

						EntityCollection compositionslots = service.RetrieveMultiple(compositionsLotQueryExpression);
						tracingService.Trace("retrieved compositionslots");
						#endregion

						if (compositionslots.Entities.Count > 0)
						{
							tracingService.Trace("compositionslots.Entities.Count > 0");
							foreach (Entity compositionslot in compositionslots.Entities)
							{
								string compositionsLotType = CRMData.GetFormattedAttributeValue(compositionslot, null, "resi_typedebientype");
								dossier.contrat.lot = new GestcomWebService.lot();
								dossier.contrat.lot.Type = compositionsLotType;
							}
						}

						#region Retrieving TMA (resi_product_fitting) based on contrat
						//Retrieving TMA (resi_product_fitting) based on contrat 
						QueryExpression productFittingQueryExpression = new QueryExpression("resi_product_fitting");
						productFittingQueryExpression.ColumnSet = new ColumnSet("resi_cout", "resi_designation", "resi_nom", "resi_amenagerid", "resi_estofferttype", "resi_isestoffert", "resi_reference");
						productFittingQueryExpression.Criteria.AddCondition("resi_contratid", ConditionOperator.Equal, contratId);

						EntityCollection productFittings = service.RetrieveMultiple(productFittingQueryExpression);
						tracingService.Trace("retrieved productFittings");
						#endregion

						List<GestcomWebService.TMA> lstTest = new List<GestcomWebService.TMA>();
						if (productFittings.Entities.Count > 0)
						{
							tracingService.Trace("productFittings.Entities.Count > 0");
							foreach (Entity productFitting in productFittings.Entities)
							{
								decimal productFittingCout = CRMData.ConvertMoneyToDecimal(productFitting, null, "resi_cout");
								string productFittingDesignation = CRMData.GetAttributeValue<string>(productFitting, null, "resi_designation");
								string productFittingName = CRMData.GetAttributeValue<string>(productFitting, null, "resi_nom");
								EntityReference productFittingAmenager = CRMData.GetAttributeValue<EntityReference>(productFitting, null, "resi_amenagerid");
								string productFittingEstOffert = CRMData.GetFormattedAttributeValue(productFitting, null, "resi_estofferttype");
								bool productFittingIsEstOffert = CRMData.GetAttributeValue<bool>(productFitting, null, "resi_isestoffert");
								string productFittingReference = CRMData.GetAttributeValue<string>(productFitting, null, "resi_reference");
								GestcomWebService.TMA tma = new GestcomWebService.TMA();
								tma.cost_ttc = productFittingCout;
								tma.Designation = productFittingDesignation;
								tma.is_offered = productFittingEstOffert;
								tma.ref_option = productFittingReference;
								lstTest.Add(tma);
								//lstTest.Add(new GestcomWebService.TMA() { cost_ttc = productFittingCout });
								//lstTest.Add(new GestcomWebService.TMA() { Designation = productFittingDesignation });
								//lstTest.Add(new GestcomWebService.TMA() { is_offered = productFittingEstOffert });
								//lstTest.Add(new GestcomWebService.TMA() { ref_option = productFittingReference });

								//Contrat.TMA.Nom = productFittingName;
								//Contrat.TMA.designation = productFittingDesignation;
								//Contrat.TMA.cout = productFittingCout;
								//if (productFittingAmenager != null)
								//{ Contrat.TMA.amenager = productFittingAmenager.Name; }
								//Contrat.TMA.est_offert = productFittingEstOffert;
								//Contrat.TMA.is_est_offert = productFittingIsEstOffert;
								//Contrat.TMA.Reference = productFittingReference;
							}
						}

						dossier.TMA = lstTest.ToArray();

						//Fetch the Commission related with the Contrat
						EntityReference commission = CRMData.GetAttributeValue<EntityReference>(contratEntity, null, "resi_commissionnerid");
						if (commission != null)
						{
							tracingService.Trace("commission != null");
							//Retrieve the Commission entity
							Entity commissionEntity = service.Retrieve("resi_comms", programme.Id, new ColumnSet("resi_commsid", "resi_name", "resi_partverseeelareservation", "resi_partverseelacte", "resi_primealacte", "resi_tauxderepartition", "resi_totalvariable"));
							if (commissionEntity != null)
							{
								tracingService.Trace("commissionEntity != null");
								string commissionName = CRMData.GetAttributeValue<string>(commissionEntity, null, "resi_name");
								decimal commissionPartVerseeALaReservation = CRMData.ConvertMoneyToDecimal(commissionEntity, null, "resi_partverseeelareservation");
								decimal commissionPartVerseeALActe = CRMData.ConvertMoneyToDecimal(commissionEntity, null, "resi_partverseelacte");
								decimal commissionPrimeALActe = CRMData.ConvertMoneyToDecimal(commissionEntity, null, "resi_primealacte");
								decimal commissionTauxDeRepartitionPercent = CRMData.GetAttributeValue<decimal>(commissionEntity, null, "resi_tauxderepartition");
								decimal commissionTotalVariable = CRMData.ConvertMoneyToDecimal(commissionEntity, null, "resi_totalvariable");

								dossier.commissions = new GestcomWebService.commissions();
								dossier.commissions.commission_rate = commissionTauxDeRepartitionPercent;
								dossier.commissions.commission_amount = commissionTotalVariable;
							}
						}
						//string contratXml = CreateXML(Contrat);
						//tracingService.Trace(contratXml);

						//Update "Commentaires sur le récapitulatif" field of the Current Contrat record to display the XML generated 
						//Entity contrat = new Entity("resi_contrat");
						//contrat.Id = contratEntity.Id;
						//contrat.Attributes.Add("resi_commentairessurlerecap", contratXml);
						//service.Update(contrat);

						string returnMessage = string.Empty;

						//Code to call a web service
						var serviceClient = new GestcomWebService.portTypeClient();
						serviceClient.wsResiDispatcherOp(dossier, out returnMessage);

						//Update "Commentaires sur le récapitulatif" field of the Current Contrat record to display the Web Service response 
						Entity contrat = new Entity("resi_contrat");
						contrat.Id = contratEntity.Id;
						contrat.Attributes.Add("resi_commentairessurlerecap", returnMessage);
						service.Update(contrat);
					}
				}
			}
			catch (Exception ex)
			{
				throw new Exception(ex.Message);
			}
		}

		private static string CreateXML(Contrat Contrat)
		{
			XmlDocument xmlDoc = new XmlDocument();   //Represents an XML document, 
																								// Initializes a new instance of the XmlDocument class.          
			XmlSerializer xmlSerializer = new XmlSerializer(Contrat.GetType());
			// Creates a stream whose backing store is memory. 
			using (MemoryStream xmlStream = new MemoryStream())
			{
				xmlSerializer.Serialize(xmlStream, Contrat);
				xmlStream.Position = 0;
				//Loads the XML document from the specified string.
				xmlDoc.Load(xmlStream);
				return xmlDoc.InnerXml;
			}
		}
	}
}

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
	public class MergeContratDocument : CodeActivity
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

			byte[] salesLiteratureItemDocumentBodyByteArray;
			tracingService.Trace("Workflow execution started");
			try
			{
				if (wfContext == null)
				{
					throw new Exception("Workflow context est vide");
				}

				//Get the current Contrat record Id
				Guid contratId = wfContext.PrimaryEntityId;
				tracingService.Trace("Current Contrat Id fetched");

				if (contratId != null)
				{
					tracingService.Trace("contratId is not null");
					//Fetch current Contrat entity record 
					Entity contratEntity = CRMData.GetEntityRecordByGuid(service, "resi_contrat", contratId, new string[] { "resi_concernerid", "resi_commissionnerid" });
					if (contratEntity != null)
					{
						tracingService.Trace("contratEntity is not null");
						#region retriveing Notes
						QueryExpression notesQueryExpression = new QueryExpression("annotation");
						notesQueryExpression.ColumnSet = new ColumnSet("filename", "subject", "annotationid", "documentbody", "objectid", "objecttypecode");
						notesQueryExpression.Criteria.AddCondition("objectid", ConditionOperator.Equal, contratId);
						notesQueryExpression.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, 1010);
						EntityCollection notesEntityCollection = service.RetrieveMultiple(notesQueryExpression);
						#endregion

						if (notesEntityCollection != null && notesEntityCollection.Entities.Count > 0)
						{
							tracingService.Trace("notesEntityCollection is not null and notesEntityCollection.Entities.Count > 0");
							tracingService.Trace(string.Format("Notes Counts = {0}", notesEntityCollection.Entities.Count));

							string contratNoteFileName = notesEntityCollection.Entities[0].Attributes["filename"].ToString();

							//converting document body content to bytes
							byte[] contratNoteDocumentByteArray = Convert.FromBase64String(notesEntityCollection.Entities[0].Attributes["documentbody"].ToString());

							if (contratNoteDocumentByteArray != null)
							{
								tracingService.Trace("Filename: {0}, file size in byte: {1}", contratNoteFileName, contratNoteDocumentByteArray.Length);
							}
						}
						tracingService.Trace("Current Contrat Entity fetched");

						//Fetch the Programme related with the Contrat
						EntityReference contratRelatedProgramme = CRMData.GetAttributeValue<EntityReference>(contratEntity, null, "resi_concernerid");
						if (contratRelatedProgramme != null)
						{
							tracingService.Trace("programme != null");
							//Retrieve the Programme entity
							//Entity programmeEntity = service.Retrieve("resi_programme", programme.Id, new ColumnSet("resi_name", "resi_idgestcom", "resi_idwincom"));

							//#region retriveing Sales Literature Items
							//QueryExpression salesLiteratureItemExpression = new QueryExpression("salesliteratureitem");
							//salesLiteratureItemExpression.ColumnSet = new ColumnSet("documentbody");
							//// Add link-entity QEsalesliteratureitem_resi_resi_programme_salesliterature
							//var salesliteratureitem_resi_resi_programme_salesliteratureLinkEntity = salesLiteratureItemExpression.AddLink("resi_resi_programme_salesliterature", "resi_resi_programme_salesliteratureid", "salesliteratureid");
							//salesliteratureitem_resi_resi_programme_salesliteratureLinkEntity.LinkCriteria.AddCondition("resi_programmeid", ConditionOperator.Equal, programmeEntity.Id);
							//EntityCollection salesliteratureitemEntityCollection = service.RetrieveMultiple(salesLiteratureItemExpression);
							//#endregion

							#region retriveing Sales Literature Items
							QueryExpression programmeExpression = new QueryExpression("resi_programme");
							programmeExpression.ColumnSet = new ColumnSet("createdby");
							programmeExpression.Criteria.AddCondition("resi_programmeid", ConditionOperator.Equal, contratRelatedProgramme.Id);
							// Add link-entity programmeAndSalesLiteratureLinkEntity
							var programmeAndSalesLiteratureLinkEntities = programmeExpression.AddLink("resi_resi_programme_salesliterature", "resi_programmeid", "resi_programmeid");
							// Add link-entity resi_programme_resi_resi_programme_salesliterature_salesliteratureitemLinkEntity
							var programmeAndSalesLiteratureAndSalesLiteratureItemLinkEntities = programmeAndSalesLiteratureLinkEntities.AddLink("salesliteratureitem", "salesliteratureid", "salesliteratureid", JoinOperator.LeftOuter);
							programmeAndSalesLiteratureAndSalesLiteratureItemLinkEntities.Columns.AddColumns("documentbody");
							EntityCollection programmeEntityCollection = service.RetrieveMultiple(programmeExpression);
							#endregion

							if (programmeEntityCollection.Entities.Count > 0)
							{
								tracingService.Trace("programmeEntityCollection.Entities.Count > 0");
								foreach (Entity resi_Programme in programmeEntityCollection.Entities)
								{
									string salesLiteratureItemDocumentBody = CRMData.GetAttributeValueFromAliasedValue<string>(resi_Programme, "salesliteratureitem2.documentbody");
									tracingService.Trace("salesLiteratureItemDocumentBody length: {0}", salesLiteratureItemDocumentBody.Length.ToString());
									
									//converting document body content to bytes
									salesLiteratureItemDocumentBodyByteArray = Convert.FromBase64String(notesEntityCollection.Entities[0].Attributes["salesliteratureitem2.documentbody"].ToString());
									tracingService.Trace("File size in byte: {0}", salesLiteratureItemDocumentBody.Length);
								}
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				throw new Exception(ex.Message);
			}
		}
	}
}

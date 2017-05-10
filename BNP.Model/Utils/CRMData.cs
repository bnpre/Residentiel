// -----------------------------------------------------------------------
// <copyright file="CRMData.cs" company="GroupeCRMDemo">
//      Copyright (c) Groupe CRMDemo.  All rights reserved.
// </copyright>
// -----------------------------------------------------------------------

namespace BNP_Model.Utils
{
	using System;
	using System.Collections.Generic;
	using Microsoft.Xrm.Sdk;
	using Microsoft.Xrm.Sdk.Query;
	using Microsoft.Crm.Sdk.Messages;
	using Microsoft.Xrm.Sdk.Metadata;
	using Microsoft.Xrm.Sdk.Messages;

	/// <summary>
	/// Represents common functions to create and update CRM data
	/// </summary>
	public sealed class CRMData
	{
		/// <summary>
		/// The organization service
		/// </summary>
		private IOrganizationService crmService;

		/// <summary>
		/// Initializes a new instance of the <see cref="CRMData"/> class.
		/// </summary>
		/// <param name="service">The service.</param>
		public CRMData(IOrganizationService service)
		{
			this.crmService = service;
		}

		/// <summary>
		/// Gets the entity records.
		/// </summary>
		/// <param name="entityName">Name of the entity.</param>
		/// <param name="conditionAttributeName">Name of the condition attribute.</param>
		/// <param name="conditionAttributeValue">The condition attribute value.</param>
		/// <param name="attributeNames">The attribute names.</param>
		/// <returns>
		/// Returns entity collection
		/// </returns>
		public EntityCollection GetEntityRecords(string entityName, string conditionAttributeName, object conditionAttributeValue, params string[] attributeNames)
		{
			QueryExpression qe = new QueryExpression();
			//#GRC-2185
			qe.NoLock = true;
			qe.EntityName = entityName;
			qe.ColumnSet = new ColumnSet(attributeNames);
			if (!string.IsNullOrEmpty(conditionAttributeName) && conditionAttributeValue != null)
			{
				qe.Criteria.AddCondition(new ConditionExpression(conditionAttributeName, ConditionOperator.Equal, conditionAttributeValue));
				qe.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
			}

			qe.AddOrder("createdon", OrderType.Descending);
			EntityCollection col = this.crmService.RetrieveMultiple(qe);
			return col;
		}

		/// <summary>
		/// Gets the entity records.
		/// </summary>
		/// <param name="entityName">Name of the entity.</param>
		/// <param name="conditionAttributeName">Name of the condition attribute.</param>
		/// <param name="conditionAttributeValue">The condition attribute value.</param>
		/// <param name="attributeNames">The attribute names.</param>
		/// <returns>
		/// Returns entity collection
		/// </returns>
		public EntityCollection GetEntityRecordsAll(string entityName, string conditionAttributeName, object conditionAttributeValue, params string[] attributeNames)
		{
			QueryExpression qe = new QueryExpression();
			//#GRC-2185
			qe.NoLock = true;
			qe.EntityName = entityName;
			qe.ColumnSet = new ColumnSet(attributeNames);
			if (!string.IsNullOrEmpty(conditionAttributeName) && conditionAttributeValue != null)
			{
				qe.Criteria.AddCondition(new ConditionExpression(conditionAttributeName, ConditionOperator.Equal, conditionAttributeValue));
			}

			qe.AddOrder("createdon", OrderType.Descending);
			EntityCollection col = this.crmService.RetrieveMultiple(qe);
			return col;
		}

		/// <summary>
		/// Gets the entity records.
		/// </summary>
		/// <param name="entityName">Name of the entity.</param>
		/// <param name="conditionAttributeName">Name of the condition attribute.</param>
		/// <param name="conditionAttributeValue">The condition attribute value.</param>
		/// <param name="linkEntities">The link entities.</param>
		/// <param name="attributeNames">The attribute names.</param>
		/// <returns>
		/// Returns EntityCollection
		/// </returns>
		public EntityCollection GetEntityRecords(string entityName, string conditionAttributeName, object conditionAttributeValue, List<LinkEntity> linkEntities, params string[] attributeNames)
		{
			QueryExpression qe = new QueryExpression();
			//#GRC-2185
			qe.NoLock = true;
			qe.EntityName = entityName;
			qe.ColumnSet = new ColumnSet(attributeNames);

			if (!string.IsNullOrEmpty(conditionAttributeName) && conditionAttributeValue != null)
			{
				qe.Criteria.AddCondition(new ConditionExpression(conditionAttributeName, ConditionOperator.Equal, conditionAttributeValue));
				qe.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
			}

			qe.AddOrder("createdon", OrderType.Descending);
			qe.LinkEntities.AddRange(linkEntities);
			EntityCollection col = this.crmService.RetrieveMultiple(qe);
			return col;
		}

		/// <summary>
		/// Gets the entity records.
		/// </summary>
		/// <param name="entityName">Name of the entity.</param>
		/// <param name="conditionExpressions">The condition expressions.</param>
		/// <param name="linkEntities">The link entities.</param>
		/// <param name="attributeNames">The attribute names.</param>
		/// <returns>Returns Entity Collection</returns>
		public EntityCollection GetEntityRecords(string entityName, List<ConditionExpression> conditionExpressions, List<LinkEntity> linkEntities, params string[] attributeNames)
		{
			QueryExpression qe = new QueryExpression();
			//#GRC-2185
			qe.NoLock = true;
			qe.EntityName = entityName;
			qe.ColumnSet = new ColumnSet(attributeNames);

			foreach (ConditionExpression expression in conditionExpressions)
			{
				qe.Criteria.AddCondition(expression);
			}

			qe.AddOrder("createdon", OrderType.Descending);
			qe.LinkEntities.AddRange(linkEntities);
			EntityCollection col = this.crmService.RetrieveMultiple(qe);
			return col;
		}

		/// <summary>
		/// Gets the entity records.
		/// </summary>
		/// <param name="entityName">Name of the entity.</param>
		/// <param name="conditionExpressions">The condition expressions.</param>
		/// <param name="attributeNames">The attribute names.</param>
		/// <returns>
		/// Returns EntityCollection
		/// </returns>
		public EntityCollection GetEntityRecords(string entityName, List<ConditionExpression> conditionExpressions, params string[] attributeNames)
		{
			QueryExpression qe = new QueryExpression();
			//#GRC-2185
			qe.NoLock = true;
			qe.EntityName = entityName;
			qe.ColumnSet = new ColumnSet(attributeNames);
			qe.Criteria.Conditions.AddRange(conditionExpressions);
			EntityCollection col = this.crmService.RetrieveMultiple(qe);
			return col;
		}

		/// <summary>
		/// Gets the entity record by unique identifier.
		/// </summary>
		/// <param name="entityName">Name of the entity.</param>
		/// <param name="guid">The unique identifier.</param>
		/// <param name="attributesNames">Name of the attributes.</param>
		/// <returns>
		/// Returns Entity Record
		/// </returns>
		public static Entity GetEntityRecordByGuid(IOrganizationService crmService, string entityName, Guid guid, params string[] attributesNames)
		{
			return crmService.Retrieve(entityName, guid, new ColumnSet(attributesNames));
		}

		/// <summary>
		/// Creates the entity record.
		/// </summary>
		/// <param name="entity">The entity.</param>
		/// <returns>Return Guid</returns>
		public Guid CreateEntityRecord(Entity entity)
		{
			return this.crmService.Create(entity);
		}

		/// <summary>
		/// Updates the entity record.
		/// </summary>
		/// <param name="entity">The entity.</param>
		public void UpdateEntityRecord(Entity entity)
		{
			this.crmService.Update(entity);
		}

		/// <summary>
		/// Executes the specified request.
		/// </summary>
		/// <param name="request">The request.</param>
		/// <returns>Returns OrganizationResponse</returns>
		public OrganizationResponse Execute(OrganizationRequest request)
		{
			return this.crmService.Execute(request);
		}

		/// <summary>
		/// Get the attribute value from specified entity, if not found in entity, then get it from entity image.
		/// </summary>
		/// <typeparam name="T">Type of Value</typeparam>
		/// <param name="entity">Entity for get attribute value.</param>
		/// <param name="entityImage">Entity image for get attribute value.</param>
		/// <param name="attributeName">Attribute name for get a value.</param>
		/// <returns>T type attribute value</returns>
		public static T GetAttributeValue<T>(Entity entity,
																				 Entity entityImage,
																				 String attributeName)
		{
			T attributeValue = default(T);

			try
			{
				if (entity != null &&
						entity.Attributes.Contains(attributeName))
				{
					attributeValue = entity.GetAttributeValue<T>(attributeName);
				}
				else if (entityImage != null &&
								 entityImage.Attributes.Contains(attributeName))
				{
					attributeValue = entityImage.GetAttributeValue<T>(attributeName);
				}
			}
			catch { } //It will return default.

			return attributeValue;
		}

		/// <summary>
		/// Add an item into queue.
		/// </summary>
		/// <param name="queueId">Destination queue id to change a queue.</param>
		/// <param name="target">Entity reference target type to add source queue from target queue.</param>
		/// <returns>false if fail to add otherwise true.</returns>
		public bool AddItemIntoQueue(Guid queueId, EntityReference target)
		{
			bool addedSuccessfully = false;
			try
			{
				addedSuccessfully = TransferItemFromSourceQueueIntoDestinationQueue(Guid.Empty, queueId, target);
			}
			catch (Exception exception)
			{
				throw new InvalidPluginExecutionException(string.Format("Error while adding item in queue. Error : {0}", exception.Message));
			}
			return addedSuccessfully;
		}

		/// <summary>
		/// Transfer an item into queue.
		/// </summary>
		/// <param name="sourceQueueId">Source queue id to change a queue.</param>
		/// <param name="destinationQueueId">Destination queue id to change a queue.</param>
		/// <param name="target">Entity reference target type to add source queue from target queue.</param>
		/// <returns>false if fail to add otherwise true.</returns>
		public bool TransferItemFromSourceQueueIntoDestinationQueue(Guid sourceQueueId, Guid destinationQueueId, EntityReference target)
		{
			bool addedSuccessfully = false;
			try
			{
				AddToQueueRequest routeRequest = new AddToQueueRequest()
				{
					Target = target,
					DestinationQueueId = destinationQueueId
				};

				if (sourceQueueId != Guid.Empty)
				{ routeRequest.SourceQueueId = sourceQueueId; }

				// Execute the Request
				Execute(routeRequest);
				addedSuccessfully = true;
			}
			catch (Exception exception)
			{
				throw new InvalidPluginExecutionException(string.Format("Error while adding item in queue. Error : {0}", exception.Message));
			}
			return addedSuccessfully;
		}

		/// <summary>
		/// Method to prepare a requaest for Colsing an incident
		/// </summary>
		/// <param name="service"></param>
		/// <param name="incidentId"></param>
		/// <param name="conclusion"></param>
		/// <returns>CloseIncidentRequest</returns>
		public static CloseIncidentRequest CreateNewIndicentResolutionRequest(IOrganizationService service, Guid incidentId, string conclusion, OptionSetValue resolved, string description)
		{
			CloseIncidentRequest closeRequest = null;
			try
			{
				Entity incidentResolution = new Entity("incidentresolution");

				incidentResolution.Attributes.Add("subject", conclusion);
				incidentResolution.Attributes.Add("incidentid", new EntityReference("incident", incidentId));
				incidentResolution.Attributes.Add("description", description);
				incidentResolution.Attributes.Add("statuscode", resolved);

				// Create the request to close the incident, and set its resolution to the 
				// resolution created above
				closeRequest = new CloseIncidentRequest();
				closeRequest.IncidentResolution = incidentResolution;
				closeRequest.Status = resolved;
			}
			catch { }

			return closeRequest;
		}

		/// Update Statuscode of Entity
		/// <summary>
		/// Activates/Deactivates the record.
		/// </summary>
		/// <param name="entity">The entity.</param>
		/// <param name="id">The identifier.</param>
		/// <param name="service">The organization service.</param>
		/// 
		public static void UpdateEntityStatusCode(Guid recordId, string recordEntityName, IOrganizationService service, int stateCode, int statuscode)
		{

			//tracingService.Trace("les infos à executer sont " + entity.Id + " " + entity.LogicalName + stateCode + statuscode);
			SetStateRequest setStateRequest = new SetStateRequest()
			{
				EntityMoniker = new EntityReference
				{
					Id = recordId,
					LogicalName = recordEntityName,
				},
				State = new OptionSetValue(stateCode),
				Status = new OptionSetValue(statuscode)
			};
			service.Execute(setStateRequest);

		}

		/// <summary>
		/// Adds the Attribute if does not exist else Updates the value.
		/// </summary>
		/// <typeparam name="T">Type of Value</typeparam>
		/// <param name="entity">Entity to which the value is to be added.</param>
		/// <param name="attributeName">The Name of the Attribute which is to be added.</param>
		/// <param name="value">Value of the Attribute.</param>
		public static void AddAttribute<T>(Entity entity,
																			 String attributeName,
																			 T value)
		{
			try
			{
				if (entity.Attributes.Contains(attributeName))
				{ entity.Attributes[attributeName] = value; }
				else
				{ entity.Attributes.Add(attributeName, value); }
			}
			catch { } //Suppress the Exception.
		}

		/// <summary>
		/// Validates that the Target is an Entity of the required LogicalName.
		/// </summary>
		/// <param name="entityName">The EntityName which is tobe verified as Target entity.</param>
		/// <param name="pluginExecutionContext">IPluginExecutionContext for validate.</param>
		/// <returns></returns>
		public static Boolean ValidateTargetAsEntity(String entityName,
																								 IPluginExecutionContext pluginExecutionContext)
		{
			bool returnValue = false;
			try
			{
				returnValue = (pluginExecutionContext != null &&
											 pluginExecutionContext.InputParameters.Contains("Target") &&
											 pluginExecutionContext.InputParameters["Target"] is Entity &&
											 ((Entity)pluginExecutionContext.InputParameters["Target"]).LogicalName.Equals(entityName));
			}
			catch { } //It will return default value.

			return returnValue;
		}

    /// <summary>
    /// Validates that the Target is an EntityReference of the required LogicalName.
    /// </summary>
    /// <param name="entityName">The EntityName which is tobe verified as Target EntityReference.</param>
    /// <param name="pluginExecutionContext">IPluginExecutionContext for validate.</param>
    /// <returns></returns>
    public static Boolean ValidateTargetAsEntityReference(String entityName,
                                                          IPluginExecutionContext pluginExecutionContext)
    {
      bool returnValue = false;
      try
      {
        returnValue = (pluginExecutionContext != null &&
                       pluginExecutionContext.InputParameters.Contains("Target") &&
                       pluginExecutionContext.InputParameters["Target"] is EntityReference &&
                       ((EntityReference)pluginExecutionContext.InputParameters["Target"]).LogicalName.Equals(entityName));
      }
      catch { } //It will return default value.

      return returnValue;
    }

    /// <summary>
    /// Get the option set id as per option set text.
    /// </summary>
    /// <param name="iOrganizationService">Organization service for retrive data.</param>
    /// <param name="optionSetName">Optionset name for get value.</param>
    /// <param name="optionSetText">Optionset text for get value.</param>
    /// <returns>Optionset value</returns>
    public static OptionSetValue GetOptionSetValue(IOrganizationService iOrganizationService,
																									 String optionSetName,
																									 String optionSetText)
		{
			OptionSetValue optionSetValue = null;

			try
			{
				// Get the current options list for the retrieved attribute.
				OptionMetadata[] optionSetValues = GetOptionSetValues(iOrganizationService, optionSetName);

				foreach (OptionMetadata optionMetadata in optionSetValues)
				{
					if (optionMetadata.Label.UserLocalizedLabel.Label.ToString().Equals(optionSetText, StringComparison.CurrentCultureIgnoreCase))
					{
						optionSetValue = new OptionSetValue(optionMetadata.Value.Value);
						break;
					}
				}
			}
			catch { } //It will return null.

			return optionSetValue;
		}

		/// <summary>
		/// Get the option set Values.
		/// </summary>
		/// <param name="iOrganizationService">Organization service for retrive data.</param>
		/// <param name="optionSetName">Optionset name for get Values.</param>
		/// <returns>Optionset Values</returns>
		private static OptionMetadata[] GetOptionSetValues(IOrganizationService iOrganizationService,
																											 string optionSetName)
		{
			OptionMetadata[] optionSetValues = null;

			try
			{
				RetrieveOptionSetRequest retrieveOptionSetRequest = new RetrieveOptionSetRequest();
				retrieveOptionSetRequest.Name = optionSetName;

				//Execute the request.
				RetrieveOptionSetResponse retrieveOptionSetResponse = (RetrieveOptionSetResponse)iOrganizationService.Execute(retrieveOptionSetRequest);

				//Access the retrieved OptionSetMetadata.
				OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;

				//Get the current options list for the retrieved attribute.
				optionSetValues = retrievedOptionSetMetadata.Options.ToArray();
			}
			catch (Exception ex)
			{
				throw ex;
			}

			return optionSetValues;
		}

		/// <summary>
		/// Assigning Entity Record to a Team or User
		/// </summary>
		/// <param name="assignee"></param>
		/// <param name="assigneeId"></param>
		/// <param name="target"></param>
		/// <param name="targetId"></param>
		/// <param name="iOrganizationService"></param>
		public static void AssignEntityRecord(EntityReference assignee,
																					 EntityReference target,
																					 IOrganizationService iOrganizationService)
		{
			try
			{
				AssignRequest assignRequest = new AssignRequest()
				{
					Assignee = new EntityReference
					{
						LogicalName = assignee.LogicalName,
						Id = assignee.Id
					},

					Target = new EntityReference(target.LogicalName, target.Id)
				};

				iOrganizationService.Execute(assignRequest);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// Get the attribute value from specified entity, if not found in entity, then get it from entity image.
		/// </summary>
		/// <typeparam name="T">Type of Value</typeparam>
		/// <param name="entity">Entity for get attribute value.</param>
		/// <param name="attributeName">Attribute name for get a value.</param>
		/// <returns>T type attribute value</returns>
		public static T GetAttributeValueFromAliasedValue<T>(Entity entity,
																												 String attributeName)
		{
			T attributeValue = default(T);

			try
			{
				attributeValue = GetAttributeValueFromAliasedValue<T>(entity,
																															null,
																															attributeName);
			}
			catch { } //It will return default.

			return attributeValue;
		}

		/// <summary>
		/// Get the attribute value from specified entity, if not found in entity, then get it from entity image.
		/// </summary>
		/// <typeparam name="T">Type of Value</typeparam>
		/// <param name="entity">Entity for get attribute value.</param>
		/// <param name="entityImage">Entity image for get attribute value.</param>
		/// <param name="attributeName">Attribute name for get a value.</param>
		/// <returns>T type attribute value</returns>
		public static T GetAttributeValueFromAliasedValue<T>(Entity entity,
																												 Entity entityImage,
																												 String attributeName)
		{
			T attributeValue = default(T);

			try
			{
				AliasedValue aliasedValue = GetAttributeValue<AliasedValue>(entity,
																																		entityImage,
																																		attributeName);
				if (aliasedValue != null)
				{ attributeValue = (T)aliasedValue.Value; }
			}
			catch { } //It will return default.

			return attributeValue;
		}

		/// <summary>
		/// Get the formatted attribute value from specified entity.
		/// </summary>
		/// <param name="entity">Entity for get attribute formatted value.</param>
		/// <param name="entityImage">EntityImage for get attribute formatted value.</param>
		/// <returns>string</returns>
		public static String GetFormattedAttributeValue(Entity entity,
																										Entity entityImage,
																										String attributeName)
		{
			string formattedAttributeValue = string.Empty;

			try
			{
				if (entity != null &&
						entity.FormattedValues.Contains(attributeName))
				{ formattedAttributeValue = entity.FormattedValues[attributeName]; }
				else if (entityImage != null &&
								 entityImage.FormattedValues.Contains(attributeName))
				{ formattedAttributeValue = entityImage.FormattedValues[attributeName]; }
			}
			catch { } //It will return default.

			return formattedAttributeValue;
		}

		/// <summary>
		/// Convert money value to decimal.
		/// </summary>
		/// <param name="entity">Entity contains the attribute is to be pick.</param>
		/// <param name="entityImage">Entity image contains the attribute is to be pick.</param>
		/// <param name="attributeName">Attribute is be picked.</param>
		/// <returns>Converted decimal value.</returns>
		public static Decimal ConvertMoneyToDecimal(Entity entity,
																								Entity entityImage,
																								String attributeName)
		{
			decimal returnValue = 0;

			try
			{
				Money returnValueMoney = null;
				if (entity != null &&
						entity.Attributes.Contains(attributeName))
				{
					//Pick the money type value from an entity.
					returnValueMoney = entity.GetAttributeValue<Money>(attributeName);

					//Convert the value.
					returnValue = ConvertMoneyToDecimal(returnValueMoney);
				}

				if (entityImage != null &&
						entityImage.Attributes.Contains(attributeName) &&
						returnValueMoney == null)
				{
					//Pick the money type value from an entity.
					returnValueMoney = entityImage.GetAttributeValue<Money>(attributeName);

					//Convert the value.
					returnValue = ConvertMoneyToDecimal(returnValueMoney);
				}
			}
			catch { } //It will return default.

			return returnValue;
		}

		/// <summary>
		/// Convert money value to decimal.
		/// </summary>
		/// <param name="value">Money value is to be convert.</param>
		/// <returns>Converted decimal value.</returns>
		public static Decimal ConvertMoneyToDecimal(Money value)
		{
			decimal returnValue = 0;

			try
			{
				//Check that the money type value contains any value or not.
				if (value != null)
				{ returnValue = value.Value; }
			}
			catch { } //It will return null.

			return returnValue;
		}

		/// <summary>
		/// Activate or De-activate a Record.
		/// </summary>
		/// <param name="state">statecode</param>
		/// <param name="status">statuscode</param>
		public static void ActivateDeactivateRecord(Entity entity, int state, int status, IOrganizationService service)
		{
			try
			{
				var activateRequest = new SetStateRequest
				{
					EntityMoniker = new EntityReference
											(entity.LogicalName, entity.Id),
					State = new OptionSetValue(state),
					Status = new OptionSetValue(status)
				};
				service.Execute(activateRequest);
			}
			catch { } //It will return null.
		}
	}
}
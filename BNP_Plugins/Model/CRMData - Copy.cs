// -----------------------------------------------------------------------
// <copyright file="CRMData.cs" company="GroupeCRMDemo">
//      Copyright (c) Groupe CRMDemo.  All rights reserved.
// </copyright>
// -----------------------------------------------------------------------

namespace BNP_Plugins
{
  using System;
  using System.Collections.Generic;
  using Microsoft.Xrm.Sdk;
  using Microsoft.Xrm.Sdk.Query;
  using Microsoft.Crm.Sdk.Messages;
  using Microsoft.Xrm.Sdk.Metadata;
  using Microsoft.Xrm.Sdk.Messages;/// <summary>
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

    ///// <summary>
    ///// Return the Dictionary containing the [key,value] for a list of keys.
    ///// </summary>
    ///// <param name="keys">list of keys</param>
    ///// <returns>dictionary of key-value</returns>
    //public Dictionary<string, string> GetCleValeur(string[] keys)
    //{
    //  List<ConditionExpression> conditions = new List<ConditionExpression>();
    //  conditions.Add(new ConditionExpression("ds_cle", ConditionOperator.In, keys));
    //  EntityCollection paramList = this.GetEntityRecords(ParametreDefinition.EntityName, conditions, new string[] { "ds_cle", "ds_valeur" });
    //  var parameters = new Dictionary<string, string>();
    //  foreach (Entity entity in paramList.Entities)
    //  {
    //    ParametreDefinition param = entity.ToEntity<ParametreDefinition>();
    //    parameters.Add(param.Attributes["ds_cle"].ToString(), param.Attributes["ds_valeur"].ToString());
    //  }

    //  return parameters;
    //}

    ///// <summary>
    ///// Return the Dictionary containing the [key,value] for a list of keys.
    ///// </summary>
    ///// <param name="key">key needs to be find</param>
    ///// <returns>string contains value</returns>
    //public static string GetCleValeur(IOrganizationService service, string key)
    //{
    //  QueryExpression parameterQueryExpression = new QueryExpression(ParametreDefinition.EntityName);
    //  parameterQueryExpression.ColumnSet = new ColumnSet("ds_cle", "ds_valeur");
    //  parameterQueryExpression.Criteria.AddCondition(new ConditionExpression("ds_cle", ConditionOperator.Equal, key));
    //  EntityCollection paramList = service.RetrieveMultiple(parameterQueryExpression);
    //  string value = string.Empty;
    //  if (paramList.Entities.Count == 1)
    //  { value = CRMData.GetAttributeValue<string>(paramList.Entities[0], null, "ds_valeur"); }

    //  return value;
    //}

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
    /// 
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

    ///// <summary>
    ///// Create a new send email request from template.
    ///// </summary>
    ///// <param name="service">Organization service instance.</param>
    ///// <param name="incident">Case entity to get the details.</param>
    ///// <param name="emailTemplateTitle">Email template to create request, pass "" if you want default template from Parameter entity.</param>
    ///// <returns></returns>
    //public static SendEmailFromTemplateRequest CreateNewSendEmailFromTemplateRequest(IOrganizationService service, Entity incident, string emailTemplateTitle, ITracingService tracingService = null)
    //{
    //  SendEmailFromTemplateRequest sendEmailFromTemplateRequest = null;
    //  try
    //  {
    //    if (tracingService != null)
    //    { tracingService.Trace("Entering CreateNewSendEmailFromTemplateRequest"); }
    //    //Get the contact from selected record.
    //    EntityReference contact = null;
    //    string title = string.Empty;

    //    if (tracingService != null)
    //    { tracingService.Trace("Picking Contact"); }
    //    if (incident.Attributes.Contains("customerid"))
    //    {
    //      tracingService.Trace("incident.Attributes.Contains(customerid)");
    //      contact = incident.GetAttributeValue<EntityReference>("customerid");
    //      tracingService.Trace("contact.LogicalName : " + contact.LogicalName);
    //      if (contact.LogicalName != "contact")
    //      { contact = null; }
    //    }

    //    if (incident.Attributes.Contains("title"))
    //    { title = incident.GetAttributeValue<string>("title"); }

    //    if (tracingService != null)
    //    { tracingService.Trace("Title : " + title); }

    //    if (contact != null)
    //    {
    //      if (tracingService != null)
    //      { tracingService.Trace("icnident customer id logical name : " + contact.LogicalName); }

    //      if (tracingService != null)
    //      { tracingService.Trace("contact != null"); }

    //      EntityReference to = contact;
    //      Entity toEntity = new Entity("activityparty");
    //      toEntity.Attributes.Add("partyid", to);
    //      List<Entity> lstTo = new List<Entity>();
    //      lstTo.Add(toEntity);

    //      EntityReference from = null;
    //      string backofficeRecetteQueueId = CRMData.GetCleValeur(service, "Backoffice File d'attente de Recette Id");
    //      if (string.IsNullOrWhiteSpace(backofficeRecetteQueueId) == false)
    //      { from = new EntityReference("queue", new Guid(backofficeRecetteQueueId)); }

    //      if (from != null)
    //      {
    //        if (tracingService != null)
    //        { tracingService.Trace("from != null"); }

    //        //EntityReference from = incident.GetAttributeValue<EntityReference>("ownerid");
    //        Entity fromEntity = new Entity("activityparty");
    //        fromEntity.Attributes.Add("partyid", from);
    //        List<Entity> lstFrom = new List<Entity>();
    //        lstFrom.Add(fromEntity);

    //        if (tracingService != null)
    //        { tracingService.Trace("Create an e-mail message"); }
    //        // Create an e-mail message.
    //        Entity email = new Entity("email");
    //        email.Attributes.Add("to", lstTo.ToArray());
    //        email.Attributes.Add("from", lstFrom.ToArray());
    //        email.Attributes.Add("subject", string.Format("Closing an incident '{0}'.", title));
    //        email.Attributes.Add("description", string.Format("Closing an incident '{0}'.", title));
    //        email.Attributes.Add("directioncode", true);
    //        email.Attributes.Add("regardingobjectid", new EntityReference(incident.LogicalName, incident.Id));

    //        if (tracingService != null)
    //        { tracingService.Trace("Checking emailTemplateTitle"); }
    //        Guid emailTemplateId = Guid.Empty;
    //        if (string.IsNullOrWhiteSpace(emailTemplateTitle) == false)
    //        {
    //          if (tracingService != null)
    //          { tracingService.Trace("string.IsNullOrWhiteSpace(emailTemplateTitle) == false"); }
    //          //Get the email template id based on title.
    //          QueryExpression emailTemplateQueryExpression = new QueryExpression("template");
    //          emailTemplateQueryExpression.ColumnSet = new ColumnSet("templateid", "templatetypecode");
    //          emailTemplateQueryExpression.Criteria.AddCondition("title", ConditionOperator.Equal, emailTemplateTitle);
    //          EntityCollection emailTemplateEntityCollection = service.RetrieveMultiple(emailTemplateQueryExpression);
    //          if (emailTemplateEntityCollection.Entities.Count > 0)
    //          { emailTemplateId = emailTemplateEntityCollection.Entities[0].Id; }
    //          if (tracingService != null)
    //          { tracingService.Trace("Got emailTemplateId : " + emailTemplateId.ToString()); }
    //        }
    //        else
    //        {
    //          if (tracingService != null)
    //          { tracingService.Trace("string.IsNullOrWhiteSpace(emailTemplateTitle) == true"); }
    //          string strEmailTemplateId = CRMData.GetCleValeur(service, "Template Email de Base Service Client");
    //          if (string.IsNullOrWhiteSpace(strEmailTemplateId) == false)
    //          { emailTemplateId = new Guid(strEmailTemplateId); }
    //          if (tracingService != null)
    //          { tracingService.Trace("Got emailTemplateId : " + emailTemplateId.ToString()); }
    //        }

    //        if (emailTemplateId != Guid.Empty)
    //        {
    //          if (tracingService != null)
    //          { tracingService.Trace("emailTemplateId != Guid.Empty"); }
    //          sendEmailFromTemplateRequest = new SendEmailFromTemplateRequest();
    //          sendEmailFromTemplateRequest.Target = email;
    //          // Use a built-in Email Template of type "contact".
    //          sendEmailFromTemplateRequest.TemplateId = emailTemplateId;
    //          // The regarding Id is required, and must be of the same type as the Email Template.
    //          sendEmailFromTemplateRequest.RegardingId = incident.Id;
    //          sendEmailFromTemplateRequest.RegardingType = incident.LogicalName;
    //          if (tracingService != null)
    //          { tracingService.Trace("Created request sendEmailFromTemplateRequest"); }
    //        }
    //      }
    //      else
    //      { new InvalidPluginExecutionException("Backoffice Recette queue does not exists."); }
    //    }
    //  }
    //  catch (Exception ex)
    //  {
    //    if (tracingService != null)
    //    { tracingService.Trace("Exception : '" + ex.Message + "'"); }
    //  }

    //  return sendEmailFromTemplateRequest;
    //}

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

    #region ValidateTargetAsEntityReference
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
    #endregion

    #region ValidateTargetAsEntity
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
    #endregion

    #region GetOptionSetValue
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
    #endregion

    #region GetOptionSetValues
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
    #endregion

    #region AssignEntityRecord
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
    #endregion

    #region GetAttributeValueFromAliasedValue
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
    #endregion

    #region GetAttributeValueFromAliasedValue
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
		#endregion

		#region GetFormattedAttributeValue
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
		#endregion
	}
}
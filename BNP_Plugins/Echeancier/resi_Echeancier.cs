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
  public class PostEcheancierCreate : IPlugin
  {
    #region Public Methods
    /// <summary>
    /// Executes plug-in code in response to an event.
    /// </summary>
    /// <param name="serviceProvider">Type: Returns_IServiceProvider. A container for service objects. Contains references to the plug-in execution context (<see cref="T:Microsoft.Xrm.Sdk.IPluginExecutionContext"></see>), tracing service (<see cref="T:Microsoft.Xrm.Sdk.ITracingService"></see>), organization service (<see cref="T:Microsoft.Xrm.Sdk.IOrganizationServiceFactory"></see>), and notification service (<see cref="T:Microsoft.Xrm.Sdk.IServiceEndpointNotificationService"></see>).</param>
    /// <exception cref="InvalidPluginExecutionException">
    /// Invalid 
    /// or
    /// An error occurred in the PostEcheancierCreate plug-in. + ex.Message
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
      if (CRMData.ValidateTargetAsEntity("resi_echeancier", context))
      {
        try
        {
          iTracingService.Trace("Context contains target entity");
          // Obtain the target entity from the input parameters.
          Entity echeancier = (Entity)context.InputParameters["Target"];
          iTracingService.Trace("Calling CalculateCumule.");
          CalculateCumule(echeancier, null, iOrganizationService, iTracingService, false, false);
          iTracingService.Trace("Finished CalculateCumule.");
        }
        catch (Exception ex)
        {
          throw new InvalidPluginExecutionException("An error occurred in the PostEcheancierCreate plug-in." + ex.Message);
        }
      }
      iTracingService.Trace("PostEcheancierCreate Plugin execution completed");
    }
    #endregion

    #region Internal Methods

    #region CalculateCumule
    /// <summary>
    /// 
    /// </summary>
    /// <param name="iTracingService"></param>
    /// <param name="iOrganizationService"></param>
    /// <param name="echeancier"></param>
    internal static void CalculateCumule(Entity echeancier, Entity echeancierImage, IOrganizationService iOrganizationService, ITracingService iTracingService, bool isRecordUpdated, bool isRecordDeleted)
    {
      iTracingService.Trace("Incoming in method CalculateCumule");
      if ((echeancier != null &&
           (echeancier.Contains("resi_initial") ||
            echeancier.Contains("resi_adapte") ||
            echeancier.Contains("resi_ordre"))) ||
           isRecordDeleted == true)
      {
        iTracingService.Trace("echeancier.Contains('resi_initial' || 'resi_adapte' || 'resi_ordre')");
        EntityReference contratEntityReference = CRMData.GetAttributeValue<EntityReference>(echeancier, echeancierImage, "resi_contratid");
        EntityReference programEntityReference = CRMData.GetAttributeValue<EntityReference>(echeancier, echeancierImage, "resi_programmeid");

        QueryExpression echeancierQueryExpression = new QueryExpression(echeancier.LogicalName);
        echeancierQueryExpression.ColumnSet = new ColumnSet("resi_echeancierid", "resi_initial", "resi_cumule", "resi_adapte", "createdon");
        echeancierQueryExpression.AddOrder("resi_ordre", OrderType.Ascending);

        iTracingService.Trace("Validate contrat entity reference has value or not");
        if (contratEntityReference != null)
        {
          if (isRecordUpdated == true || isRecordDeleted == true)
          {
            iTracingService.Trace("contratEntityReference != null");
            echeancierQueryExpression.Criteria.AddCondition("resi_contratid", ConditionOperator.Equal, contratEntityReference.Id);
            EntityCollection echeancierEntityCollection = iOrganizationService.RetrieveMultiple(echeancierQueryExpression);

            //If value in the image in null in update mode then update all the echeancier adapte value with initial value.
            decimal? adapteOldValue = CRMData.GetAttributeValue<decimal?>(echeancierImage, null, "resi_adapte");

            //Update adapte value if user insert/update adapte in any extsing record.
            if (adapteOldValue == null &&
                isRecordUpdated == true)
            {
              foreach (Entity fetchedEcheancier in echeancierEntityCollection.Entities)
              {
                //Ignore current record
                if (fetchedEcheancier.Id != echeancier.Id)
                {
                  decimal? initalValue = CRMData.GetAttributeValue<decimal>(fetchedEcheancier, null, "resi_initial");
                  Entity updateEcheancier = new Entity(fetchedEcheancier.LogicalName);
                  updateEcheancier.Id = fetchedEcheancier.Id;
                  CRMData.AddAttribute<decimal?>(updateEcheancier, "resi_adapte", initalValue);
                  //Update "Adapte" in the current fetched entity collecion in order to avoid another fetch and get the updated "Adapte" in further loops.
                  CRMData.AddAttribute<decimal?>(fetchedEcheancier, "resi_adapte", initalValue);
                  iOrganizationService.Update(updateEcheancier);
                }
              }
            }

            decimal cumuleValue = 0;
            //Update all the values from Adapte.
            foreach (Entity fetchedEcheancier in echeancierEntityCollection.Entities)
            {
              decimal? adapteValue = CRMData.GetAttributeValue<decimal?>(fetchedEcheancier, null, "resi_adapte");
              if (adapteValue != null)
              { cumuleValue += adapteValue.Value; }
              
              Entity updateEcheancier = new Entity(fetchedEcheancier.LogicalName);
              updateEcheancier.Id = fetchedEcheancier.Id;
              CRMData.AddAttribute<decimal?>(updateEcheancier, "resi_cumule", cumuleValue);
              iOrganizationService.Update(updateEcheancier);
              iTracingService.Trace("Echeancier cumule updated successfully for contrat");
            }

            //Deciding the cumumule is 100% or not.
            bool isCumuleValue100 = (cumuleValue == 100);

            iTracingService.Trace("Updating contrat is valide");
            Entity updateContrat = new Entity(contratEntityReference.LogicalName);
            updateContrat.Id = contratEntityReference.Id;
            CRMData.AddAttribute<bool>(updateContrat, "resi_iscontratvalide", isCumuleValue100);
            iOrganizationService.Update(updateContrat);
            iTracingService.Trace("'resi_iscontratvalide' updated successfully");
          }
        }
        else if (programEntityReference != null)
        {
          iTracingService.Trace("programEntityReference != null");
          echeancierQueryExpression.Criteria.AddCondition("resi_programmeid", ConditionOperator.Equal, programEntityReference.Id);
          EntityCollection echeancierEntityCollection = iOrganizationService.RetrieveMultiple(echeancierQueryExpression);

          decimal cumuleValue = 0;
          //Update all the values from Initial.
          foreach (Entity fetchedEcheancier in echeancierEntityCollection.Entities)
          {
            decimal? initalValue = CRMData.GetAttributeValue<decimal>(fetchedEcheancier, null, "resi_initial");
            if (initalValue != null)
            { cumuleValue += initalValue.Value; }

            Entity updateEcheancier = new Entity(fetchedEcheancier.LogicalName);
            updateEcheancier.Id = fetchedEcheancier.Id;
            CRMData.AddAttribute<decimal?>(updateEcheancier, "resi_cumule", cumuleValue);
            iOrganizationService.Update(updateEcheancier);
            iTracingService.Trace("Echeancier cumule updated successfully for program");
          }

          //Deciding the cumumule is 100% or not.
          bool isCumuleValue100 = (cumuleValue == 100);

          iTracingService.Trace("Updating program allowed sale");
          Entity updateProgram = new Entity(programEntityReference.LogicalName);
          updateProgram.Id = programEntityReference.Id;
          CRMData.AddAttribute<bool>(updateProgram, "resi_allowedsale", isCumuleValue100);
          iOrganizationService.Update(updateProgram);
          iTracingService.Trace("'resi_allowedsale' updated successfully");
        }
      }
    }
    #endregion

    #endregion
  }

  /// <summary>
  /// Echeancier class
  /// </summary>
  /// <seealso cref="Microsoft.Xrm.Sdk.IPlugin" />
  public class PostEcheancierUpdate : IPlugin
  {
    #region Public Methods
    /// <summary>
    /// Executes plug-in code in response to an event.
    /// </summary>
    /// <param name="serviceProvider">Type: Returns_IServiceProvider. A container for service objects. Contains references to the plug-in execution context (<see cref="T:Microsoft.Xrm.Sdk.IPluginExecutionContext"></see>), tracing service (<see cref="T:Microsoft.Xrm.Sdk.ITracingService"></see>), organization service (<see cref="T:Microsoft.Xrm.Sdk.IOrganizationServiceFactory"></see>), and notification service (<see cref="T:Microsoft.Xrm.Sdk.IServiceEndpointNotificationService"></see>).</param>
    /// <exception cref="InvalidPluginExecutionException">
    /// Invalid 
    /// or
    /// An error occurred in the PostEcheancierUpdate plug-in. + ex.Message
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
      if (CRMData.ValidateTargetAsEntity("resi_echeancier", context) &&
          context.Depth == 1) //This condition in needed to stop infinity.
      {
        try
        {
          iTracingService.Trace("Context contains target entity");
          // Obtain the target entity from the input parameters.
          Entity echeancier = (Entity)context.InputParameters["Target"];
          Entity echeancierImage = null;
          if (context.PreEntityImages.Contains("PreImagePostEcheancierUpdate") &&
              context.PreEntityImages["PreImagePostEcheancierUpdate"] is Entity)
          { echeancierImage = (Entity)context.PreEntityImages["PreImagePostEcheancierUpdate"]; }
          iTracingService.Trace("Calling CalculateCumule.");
          PostEcheancierCreate.CalculateCumule(echeancier, echeancierImage, iOrganizationService, iTracingService, true, false);
          iTracingService.Trace("Finished CalculateCumule.");
        }
        catch (Exception ex)
        {
          throw new InvalidPluginExecutionException("An error occurred in the PostEcheancierUpdate plug-in." + ex.Message);
        }
      }
      iTracingService.Trace("PostEcheancierUpdate Plugin execution completed");
    }
    #endregion
  }

  /// <summary>
  /// Echeancier class
  /// </summary>
  /// <seealso cref="Microsoft.Xrm.Sdk.IPlugin" />
  public class PostEcheancierDelete : IPlugin
  {
    #region Public Methods
    /// <summary>
    /// Executes plug-in code in response to an event.
    /// </summary>
    /// <param name="serviceProvider">Type: Returns_IServiceProvider. A container for service objects. Contains references to the plug-in execution context (<see cref="T:Microsoft.Xrm.Sdk.IPluginExecutionContext"></see>), tracing service (<see cref="T:Microsoft.Xrm.Sdk.ITracingService"></see>), organization service (<see cref="T:Microsoft.Xrm.Sdk.IOrganizationServiceFactory"></see>), and notification service (<see cref="T:Microsoft.Xrm.Sdk.IServiceEndpointNotificationService"></see>).</param>
    /// <exception cref="InvalidPluginExecutionException">
    /// Invalid 
    /// or
    /// An error occurred in the PostEcheancierDelete plug-in. + ex.Message
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
      if (CRMData.ValidateTargetAsEntityReference("resi_echeancier", context))
      {
        try
        {
          iTracingService.Trace("Context contains target entity");
          // Obtain the target entity from the input parameters.
          EntityReference echeancierEntityReference = (EntityReference)context.InputParameters["Target"];
          Entity echeancierImage = null;
          if (context.PreEntityImages.Contains("PreImagePostEcheancierDelete") &&
              context.PreEntityImages["PreImagePostEcheancierDelete"] is Entity)
          { echeancierImage = (Entity)context.PreEntityImages["PreImagePostEcheancierDelete"]; }
          iTracingService.Trace("Calling CalculateCumule.");
          PostEcheancierCreate.CalculateCumule(new Entity(echeancierEntityReference.LogicalName, echeancierEntityReference.Id), echeancierImage, iOrganizationService, iTracingService, false, true);
          iTracingService.Trace("Finished CalculateCumule.");
        }
        catch (Exception ex)
        {
          throw new InvalidPluginExecutionException("An error occurred in the PostEcheancierDelete plug-in." + ex.Message);
        }
      }
      iTracingService.Trace("PostEcheancierDelete Plugin execution completed");
    }
    #endregion
  }
}
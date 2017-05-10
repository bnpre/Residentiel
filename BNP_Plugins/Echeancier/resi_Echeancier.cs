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
          iTracingService.Trace("Calling CalculateCumule.");
          CalculateCumule(echeancier, null, iOrganizationService, iTracingService, false, false);
          iTracingService.Trace("Finished CalculateCumule.");
        }
        catch (Exception ex)
        {
          throw ex;
        }
      }
      iTracingService.Trace("Plugin execution completed");
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
            echeancier.Contains("resi_adapte"))) ||
           isRecordDeleted == true)
      {
        iTracingService.Trace("echeancier.Contains('resi_initial' || 'resi_adapte')");
        EntityReference contratEntityReference = CRMData.GetAttributeValue<EntityReference>(echeancier, echeancierImage, "resi_contratid");
        EntityReference programEntityReference = CRMData.GetAttributeValue<EntityReference>(echeancier, echeancierImage, "resi_programmeid");

        string fieldToCalculateSum = string.Empty;
        decimal? currentValue = null;
        string fetchXMLCondition = string.Empty;

        iTracingService.Trace("Getting current changed value from adapte or initial");
        if (contratEntityReference != null)
        {
          fieldToCalculateSum = "resi_adapte";
          currentValue = CRMData.GetAttributeValue<decimal?>(echeancier, null, "resi_adapte");
          fetchXMLCondition = "<condition attribute='resi_contratid' operator='eq' uiname='' uitype='resi_contrat' value='" + contratEntityReference.Id + @"' />";
        }
        else if (programEntityReference != null)
        {
          fieldToCalculateSum = "resi_initial";
          currentValue = CRMData.GetAttributeValue<decimal?>(echeancier, null, "resi_initial");
          fetchXMLCondition = "<condition attribute='resi_programmeid' operator='eq' uiname='' uitype='resi_programme' value='" + programEntityReference.Id + @"' />";
        }

        if (currentValue.HasValue ||
            isRecordDeleted == true)
        {
          string echeancierFetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='resi_echeancier'>
                                            <attribute name='resi_echeancierid' />
                                            <attribute name='resi_initial' />
                                            <attribute name='resi_cumule' />
                                            <attribute name='resi_adapte' />
                                            <attribute name='createdon'/>
                                            <order descending='false' attribute='createdon'/>
                                            <filter type='and'>
                                              " + fetchXMLCondition + @" />
                                              <condition attribute='resi_echeancierid' operator='ne' uiname='' uitype='resi_echeancier' value='" + echeancier.Id + @"' />
                                            </filter>
                                          </entity>
                                        </fetch>";

          //Get all the echeancier based on above fetch xml.
          FetchExpression echeancierFetchExpression = new FetchExpression(echeancierFetchXML);
          EntityCollection echeancierEntityCollection = iOrganizationService.RetrieveMultiple(echeancierFetchExpression);

          decimal sumValue = 0;

          foreach (Entity fetchedEcheancier in echeancierEntityCollection.Entities)
          { sumValue += CRMData.GetAttributeValue<decimal>(fetchedEcheancier, null, fieldToCalculateSum); }

          //Add current value into the sum.
          sumValue += currentValue.Value;

          //Update the sum value in Cumule%.
          if (isRecordDeleted == false)
          { CRMData.AddAttribute<decimal>(echeancier, "resi_cumule", sumValue); }

          bool isSumValue100 = (sumValue == 100);

          //Update the related entity record that indicates the cumule percent is not 100%.
          if (contratEntityReference != null)
          {
            Entity updateContrat = new Entity(contratEntityReference.LogicalName);
            updateContrat.Id = contratEntityReference.Id;
            CRMData.AddAttribute<bool>(updateContrat, "resi_iscontratvalide", isSumValue100);
            iOrganizationService.Update(updateContrat);
          }
          else if (programEntityReference != null)
          {
            Entity updateProgram = new Entity(programEntityReference.LogicalName);
            updateProgram.Id = programEntityReference.Id;
            CRMData.AddAttribute<bool>(updateProgram, "resi_allowedsale", isSumValue100);
            iOrganizationService.Update(updateProgram);
          }

          if (contratEntityReference != null &&
              isRecordUpdated == true)
          {
            //If value in the image in null in update mode then update all the echeancier adapte value with initial value.
            decimal? adapteOldValue = CRMData.GetAttributeValue<decimal?>(echeancierImage, null, "resi_adapte");
            if (adapteOldValue.HasValue == false)
            {
              decimal cumuleValue = 0;
              //Update all the values from Initial.
              foreach (Entity fetchedEcheancier in echeancierEntityCollection.Entities)
              {
                decimal? adapteValue = CRMData.GetAttributeValue<decimal?>(fetchedEcheancier, null, "resi_adapte");
                if (adapteValue.HasValue == false)
                {
                  decimal initalValue = CRMData.GetAttributeValue<decimal>(fetchedEcheancier, null, "resi_initial");
                  cumuleValue += initalValue;
                  Entity updateEcheancier = new Entity(fetchedEcheancier.LogicalName);
                  updateEcheancier.Id = fetchedEcheancier.Id;
                  CRMData.AddAttribute<decimal>(updateEcheancier, "resi_adapte", initalValue);
                  CRMData.AddAttribute<decimal>(updateEcheancier, "resi_cumule", cumuleValue);
                  iOrganizationService.Update(updateEcheancier);
                }
              }
            }
          }
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
  public class PreEcheancierUpdate : IPlugin
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
      if (CRMData.ValidateTargetAsEntity("resi_echeancier", context) &&
          context.Depth == 1)
      {
        try
        {
          iTracingService.Trace("Context contains target entity");
          // Obtain the target entity from the input parameters.
          Entity echeancier = (Entity)context.InputParameters["Target"];
          Entity echeancierImage = null;
          if (context.PreEntityImages.Contains("PreImagePreEcheancierUpdate") &&
              context.PreEntityImages["PreImagePreEcheancierUpdate"] is Entity)
          { echeancierImage = (Entity)context.PreEntityImages["PreImagePreEcheancierUpdate"]; }
          iTracingService.Trace("Calling CalculateCumule.");
          PreEcheancierCreate.CalculateCumule(echeancier, echeancierImage, iOrganizationService, iTracingService, true, false);
          iTracingService.Trace("Finished CalculateCumule.");
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

  /// <summary>
  /// Echeancier class
  /// </summary>
  /// <seealso cref="Microsoft.Xrm.Sdk.IPlugin" />
  public class PostEcheancierDelete : IPlugin
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
          PreEcheancierCreate.CalculateCumule(new Entity(echeancierEntityReference.LogicalName, echeancierEntityReference.Id), echeancierImage, iOrganizationService, iTracingService, false, true);
          iTracingService.Trace("Finished CalculateCumule.");
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
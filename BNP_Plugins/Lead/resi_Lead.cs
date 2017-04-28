////-----------------------------
//// <copyright file="PreLeadCreate.cs" company="BNPRE">
//// Copyright (c) BNPRE. All rights reserved.
//// </copyright>
//// <summary>
//// This file contains operations for Lead class.
//// </summary>
////-------------------------------------------------------------------------------------------------------------------------------
namespace BNP_Plugins
{
	using BNP_Model.Utils;
	using Microsoft.Crm.Sdk.Messages;
	using Microsoft.Xrm.Sdk;
	using Microsoft.Xrm.Sdk.Query;
	using System;
	using System.Collections.Generic;
	using System.ServiceModel;

	/// <summary>
	/// Lead Class
	/// </summary>
	/// <seealso cref="Microsoft.Xrm.Sdk.IPlugin" />
	public class PreLeadCreate : IPlugin
	{
		/// <summary>
		/// Executes plug-in code in response to an event.
		/// </summary>
		/// <param name="serviceProvider">Type: Returns_IServiceProvider. A container for service objects. Contains references to the plug-in execution context (<see cref="T:Microsoft.Xrm.Sdk.IPluginExecutionContext"></see>), tracing service (<see cref="T:Microsoft.Xrm.Sdk.ITracingService"></see>), organization service (<see cref="T:Microsoft.Xrm.Sdk.IOrganizationServiceFactory"></see>), and notification service (<see cref="T:Microsoft.Xrm.Sdk.IServiceEndpointNotificationService"></see>).</param>
		/// <exception cref="InvalidPluginExecutionException">
		/// Invalid 
		/// or
		/// An error occurred in the PreLeadCreate plug-in. + ex.Message
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
			if (CRMData.ValidateTargetAsEntity("lead", context))
			{
				try
				{
					iTracingService.Trace("Context contains target entity");
					// Obtain the target entity from the input parameters.
					Entity lead = (Entity)context.InputParameters["Target"];
					iTracingService.Trace("Context fetched.");
					bool proceedFurtherInPostPlugin = false;

					#region Fetching Attributes from Context.
					string leadName = CRMData.GetAttributeValue<string>(lead, null, "fullname");
					ReplaceWithBlankIfStringContainsNull(ref leadName);
					string leadEmail = CRMData.GetAttributeValue<string>(lead, null, "emailaddress1");
					ReplaceWithBlankIfStringContainsNull(ref leadEmail);
					string leadTelePhone = CRMData.GetAttributeValue<string>(lead, null, "telephone1");
					ReplaceWithBlankIfStringContainsNull(ref leadTelePhone);
					string leadAddressLine1 = CRMData.GetAttributeValue<string>(lead, null, "address1_line1");
					ReplaceWithBlankIfStringContainsNull(ref leadAddressLine1);
					string leadAddressLine2 = CRMData.GetAttributeValue<string>(lead, null, "address1_line2");
					ReplaceWithBlankIfStringContainsNull(ref leadAddressLine2);
					string leadAddressLine3 = CRMData.GetAttributeValue<string>(lead, null, "address1_line3");
					ReplaceWithBlankIfStringContainsNull(ref leadAddressLine3);
					string leadCity = CRMData.GetAttributeValue<string>(lead, null, "address1_city");
					ReplaceWithBlankIfStringContainsNull(ref leadCity);
					string leadPostalCode = CRMData.GetAttributeValue<string>(lead, null, "address1_postalcode");
					ReplaceWithBlankIfStringContainsNull(ref leadPostalCode);
					string identifiantDuLeadSource = CRMData.GetAttributeValue<string>(lead, null, "resi_sourceleadid");
					ReplaceWithBlankIfStringContainsNull(ref identifiantDuLeadSource);
					string origine = CRMData.GetAttributeValue<string>(lead, null, "resi_origine");
					ReplaceWithBlankIfStringContainsNull(ref origine);
					origine = origine.Trim();
					string lignage = CRMData.GetAttributeValue<string>(lead, null, "resi_mktligneage");
					ReplaceWithBlankIfStringContainsNull(ref lignage);
					lignage = lignage.Trim();
					string transmiterSystemId = CRMData.GetAttributeValue<string>(lead, null, "resi_transmitersystemid");
					ReplaceWithBlankIfStringContainsNull(ref transmiterSystemId);
					transmiterSystemId = transmiterSystemId.Trim();
					string errorMessage = string.Empty;
					EntityReference leadParentContact = null;
					iTracingService.Trace("Attributes fetched from Context.");
					#endregion

					iTracingService.Trace("LeadName = {0}", leadName);

					//Check if "resi_origine" contains Value
					if (string.IsNullOrWhiteSpace(origine) == false)
					{
						proceedFurtherInPostPlugin = true;
						iTracingService.Trace("resi_origine contains Value");

						#region Update the field "Regle De Dispatch Type" on Lead entity with the value in the field "Regle De Dispatch Type" fetched "Parametre" entity.
						iTracingService.Trace("Fetching Regle De Dispatch from Parametre entity");
						// Fetch "Paramètres" entity records 
						QueryExpression parametreQueryExpression = new QueryExpression("core_parametre");
						parametreQueryExpression.ColumnSet = new ColumnSet("core_reglededispatchtype");
						if (string.IsNullOrWhiteSpace(origine) == false)
						{
							iTracingService.Trace("origineType int value changed");
							parametreQueryExpression.Criteria.AddCondition("core_origine", ConditionOperator.Equal, origine);
						}
						if (string.IsNullOrWhiteSpace(lignage) == false)
						{
							iTracingService.Trace("lignage is not null");
							parametreQueryExpression.Criteria.AddCondition("core_lignage", ConditionOperator.Equal, lignage);
						}
						if (string.IsNullOrWhiteSpace(transmiterSystemId) == false)
						{
							iTracingService.Trace("transmiterSystemId is not null");
							parametreQueryExpression.Criteria.AddCondition("core_transmitersystemid", ConditionOperator.Equal, transmiterSystemId);
						}
						EntityCollection parametreEntityCollection = iOrganizationService.RetrieveMultiple(parametreQueryExpression);

						iTracingService.Trace("retrieved Paramètres = {0} records", parametreEntityCollection.Entities.Count);

						if (parametreEntityCollection.Entities.Count > 0)
						{
							OptionSetValue regleDeDispatchType = CRMData.GetAttributeValue<OptionSetValue>(parametreEntityCollection.Entities[0], null, "core_reglededispatchtype");
							if (regleDeDispatchType != null)
							{
								iTracingService.Trace("Regle De Dispatch Type = {0}", regleDeDispatchType.Value);
								iTracingService.Trace("regleDeDispatchType is not null");
								CRMData.AddAttribute<OptionSetValue>(lead, "resi_reglesdedispatchtype", regleDeDispatchType);
								iTracingService.Trace("regleDeDispatchType updated");
							}
						}
						#endregion

						#region identifiantDuLeadSource is not null so fetch the lead records where "resi_sourceleadid" =  identifiantDuLeadSource
						bool validateContact = false;

						if (string.IsNullOrWhiteSpace(identifiantDuLeadSource) == false)
						{
							iTracingService.Trace("Entered in the region for fetching leads based on field 'Identifiant du lead source'");
							// Fetch "Lead" entity records based on "Identifiant du lead source" of Lead
							QueryExpression leadQueryExpression = new QueryExpression("lead");
							leadQueryExpression.ColumnSet = new ColumnSet("subject", "leadid", "parentcontactid");
							leadQueryExpression.Criteria.AddCondition("resi_sourceleadid", ConditionOperator.Equal, identifiantDuLeadSource);
							EntityCollection leadEntityCollection = iOrganizationService.RetrieveMultiple(leadQueryExpression);

							iTracingService.Trace("retrieved Leads = {0} records", leadEntityCollection.Entities.Count);

							if (leadEntityCollection.Entities.Count > 0)
							{
								#region Lead entity records exists with "resi_sourceleadid" =  identifiantDuLeadSource
								string leadSubject = CRMData.GetAttributeValue<string>(leadEntityCollection.Entities[0], null, "subject");
								iTracingService.Trace("Current lead is " + leadSubject + " ID = " + leadEntityCollection.Entities[0].Id.ToString());
								iTracingService.Trace("Validating lead parent contact");
								leadParentContact = CRMData.GetAttributeValue<EntityReference>(leadEntityCollection.Entities[0], null, "parentcontactid");
								if (leadParentContact != null)
								{
									iTracingService.Trace("Lead parent contact is not null");
									#region Fetch related open opportunities where "parentContactId" = parentContactLead
									iTracingService.Trace("Entered in the region for fetching leads based on field 'Identifiant du lead source' ");
									// Fetch "Lead" entity records based on "Identifiant du lead source" of Lead
									QueryExpression opportunityQueryExpression = new QueryExpression("opportunity");
									opportunityQueryExpression.ColumnSet = new ColumnSet("opportunityid");
									opportunityQueryExpression.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
									opportunityQueryExpression.Criteria.AddCondition("parentcontactid", ConditionOperator.Equal, leadParentContact.Id);
									EntityCollection opportunityEntityCollection = iOrganizationService.RetrieveMultiple(opportunityQueryExpression);

									iTracingService.Trace("retrieved Opportunity = {0} records", opportunityEntityCollection.Entities.Count);

									if (opportunityEntityCollection.Entities.Count > 0)
									{
										#region Related Open Opportunities does exists in CRM.
										iTracingService.Trace("Related Opportunity exists from parentContact of Source Lead found");
										context.SharedVariables.Add("isContactExists", false);
										//Do nothing
										#endregion
									}
									else
									{
										iTracingService.Trace("Related Opportunity does not exists from parentContact of Source Lead found.");
										validateContact = true;
									}
									#endregion
								}
								else
								{
									iTracingService.Trace("parentContactId on fetched lead is null");
									validateContact = true;
								}
								#endregion
							}
							else//lead entity records with Identifiant Du Lead Source does not exists.
							{
								iTracingService.Trace("lead entity records with Identifiant Du Lead Source does not exists");
								validateContact = true;
							}
						}
						else//"Identifiant Du Lead Source" is null
						{
							iTracingService.Trace("Identifiant Du Lead Source is null");
							validateContact = true;
						}

						if (validateContact == true)
						{
							#region Update the field "Contact parent du prospect" on Lead entity with the Contact if exists in CRM with the same "Email", "Address1_telephone1" and "Address1_PostalCode" else create a new contact.
							EntityReference contact = Recherche_Contact(lead,
																													leadName,
																													leadEmail,
																													leadTelePhone,
																													leadAddressLine1,
																													leadAddressLine2,
																													leadAddressLine3,
																													leadPostalCode,
																													leadCity,
																													iOrganizationService,
																													iTracingService,
																													context);

							if (contact != null)
							{
								iTracingService.Trace("Contact fetched or newly created");
								context.SharedVariables.Add("isContactExists", true);
							}
							else
							{
								iTracingService.Trace("Contact does not exists");
								context.SharedVariables.Add("isContactExists", false);
							}
							#endregion
						}

						#endregion
					}
					else//if "resi_origine" does not contain Value
					{
						iTracingService.Trace("resi_origine does not contains Value");
						errorMessage = "Le champ Origine n'est pas renseign";
						UpdateErrorOnPreCreate(lead, errorMessage, iOrganizationService, iTracingService);
					}
					context.SharedVariables.Add("ProceedFurtherInPostPlugin", proceedFurtherInPostPlugin);
				}
				catch (FaultException<OrganizationServiceFault> ex)
				{
					throw new InvalidPluginExecutionException("An error occurred in the PreLeadCreate plug-in." + ex.Message);
				}
			}
		}

		#endregion

		#region Private Methods

		#region Recherche_Contact 
		private EntityReference Recherche_Contact(Entity lead,
																							string leadName,
																							string leadEmail,
																							string leadTelephone,
																							string leadAddressLine1,
																							string leadAddressLine2,
																							string leadAddressLine3,
																							string leadPostalCode,
																							string leadCity,
																							IOrganizationService iOrganizationService,
																							ITracingService iTracingService,
																							IPluginExecutionContext iPluginExecutionContext)
		{
			EntityReference newContact = null;
			if (lead != null)
			{
				iTracingService.Trace("Entered in the region for fetching related Contacts and updating parent Contact on Lead entity");
				// Fetch "Contact" entity records based on "Fullname" of Lead
				QueryExpression contactQueryExpression = new QueryExpression("contact");
				contactQueryExpression.ColumnSet = new ColumnSet("contactid");
				contactQueryExpression.Criteria.AddCondition("fullname", ConditionOperator.Equal, leadName);
				EntityCollection contactEntityCollection = iOrganizationService.RetrieveMultiple(contactQueryExpression);

				iTracingService.Trace("retrieved Contacts = {0} records", contactEntityCollection.Entities.Count);

				if (contactEntityCollection.Entities.Count > 0)//If Contact does exists with the Fullname
				{
					#region If Contact does exists with the Fullname
					iTracingService.Trace("Contact does exists with the Fullname");
					bool fetchContactBasedOnLeadTelephone = false;
					if (string.IsNullOrWhiteSpace(leadEmail) == true)
					{
						iTracingService.Trace("Incoming lead doesn't contain email.");
						fetchContactBasedOnLeadTelephone = true;
					}
					else
					{
						iTracingService.Trace("Incoming lead contains an email and fetching contact based on email");
						contactQueryExpression.Criteria.Conditions.Clear();
						contactQueryExpression.Criteria.AddCondition("emailaddress1", ConditionOperator.Equal, leadEmail);
						contactEntityCollection = iOrganizationService.RetrieveMultiple(contactQueryExpression);
						iTracingService.Trace("retrieved Contacts = {0} records", contactEntityCollection.Entities.Count);

						if (contactEntityCollection.Entities.Count > 0)//Contacts exists with the Email
						{
							#region Contacts exists with the Email
							iTracingService.Trace("Contacts exists with the Email");
							Entity contactFetched = contactEntityCollection.Entities[0];
							newContact = contactFetched.ToEntityReference();
							lead.Attributes.Add("parentcontactid", newContact);
							#endregion
						}
						else
						{ fetchContactBasedOnLeadTelephone = true; }
					}

					if (fetchContactBasedOnLeadTelephone == true)
					{
						bool fetchContactBasedOnAddress = false;

						if (string.IsNullOrWhiteSpace(leadTelephone) == true)
						{
							iTracingService.Trace("Incoming lead doesn't contain Telephone.");
							fetchContactBasedOnAddress = true;
						}
						else
						{
							#region Contacts does not exists with the Email
							iTracingService.Trace("Contacts does not exists with the Email");
							contactQueryExpression.Criteria.Conditions.Clear();
							contactQueryExpression.Criteria.AddCondition("telephone1", ConditionOperator.Equal, leadTelephone);
							contactEntityCollection = iOrganizationService.RetrieveMultiple(contactQueryExpression);
							iTracingService.Trace("retrieved Contacts = {0} records", contactEntityCollection.Entities.Count);

							if (contactEntityCollection.Entities.Count > 0)//Contacts exists with the Telephone
							{
								#region Contacts exists with the Telephone
								iTracingService.Trace("Contacts exists with the Telephone");
								Entity contactFetched = contactEntityCollection.Entities[0];
								newContact = contactFetched.ToEntityReference();
								lead.Attributes.Add("parentcontactid", newContact);
								#endregion
							}
							else
							{ fetchContactBasedOnAddress = true; }
						}
						if (fetchContactBasedOnAddress == true)
						{
							bool createNewContact = false;

							if (string.IsNullOrWhiteSpace(leadAddressLine1) == true &&
									string.IsNullOrWhiteSpace(leadAddressLine2) == true &&
									string.IsNullOrWhiteSpace(leadAddressLine3) == true &&
									string.IsNullOrWhiteSpace(leadPostalCode) == true &&
									string.IsNullOrWhiteSpace(leadCity) == true)
							{
								iTracingService.Trace("Incoming lead doesn't contain address.");
								createNewContact = true;
							}
							else
							{
								#region Contacts does not exists with the Telephone
								iTracingService.Trace("Contacts does not exists with the Telephone");
								contactQueryExpression.Criteria.Conditions.Clear();
								contactQueryExpression.Criteria.AddCondition("address1_line1", ConditionOperator.Equal, leadAddressLine1);
								contactQueryExpression.Criteria.AddCondition("address1_line2", ConditionOperator.Equal, leadAddressLine2);
								contactQueryExpression.Criteria.AddCondition("address1_line3", ConditionOperator.Equal, leadAddressLine3);
								contactQueryExpression.Criteria.AddCondition("address1_postalcode", ConditionOperator.Equal, leadPostalCode);
								contactQueryExpression.Criteria.AddCondition("address1_city", ConditionOperator.Equal, leadCity);
								contactEntityCollection = iOrganizationService.RetrieveMultiple(contactQueryExpression);
								iTracingService.Trace("retrieved Contacts = {0} records", contactEntityCollection.Entities.Count);

								if (contactEntityCollection.Entities.Count > 0)//Contacts exists with the Address + PostalCode + City
								{
									#region Contacts exists with the Address + PostalCode + City
									iTracingService.Trace("Contacts exists with the Address + PostalCode + City");
									Entity contactFetched = contactEntityCollection.Entities[0];
									newContact = contactFetched.ToEntityReference();
									lead.Attributes.Add("parentcontactid", newContact);
									#endregion
								}
								else
								{ createNewContact = true; }
							}
							if (createNewContact == true)
							{
								#region Contacts does not exists with the Address + PostalCode + City
								iTracingService.Trace("Creating a new Contact from lead data");
								newContact = CreateNewContact(leadName,
																							leadEmail,
																							leadTelephone,
																							leadAddressLine1,
																							leadAddressLine2,
																							leadAddressLine3,
																							leadPostalCode,
																							leadCity,
																							iOrganizationService,
																							iTracingService);
								lead.Attributes.Add("parentcontactid", newContact);
								#endregion
							}
							#endregion
						}
						#endregion
					}
					#endregion
				}
				else//If Contact does not exists with the Fullname 
				{
					#region If Contact does not exists with the Fullname 
					iTracingService.Trace("contacts fetched based on fullname returns 0 records");
					contactQueryExpression.Criteria.Conditions.Clear();

					if (string.IsNullOrWhiteSpace(leadEmail) == false &&
					 string.IsNullOrWhiteSpace(leadTelephone) == false &&
					 string.IsNullOrWhiteSpace(leadAddressLine1) == false &&
					 string.IsNullOrWhiteSpace(leadAddressLine2) == false &&
					 string.IsNullOrWhiteSpace(leadAddressLine3) == false &&
					 string.IsNullOrWhiteSpace(leadPostalCode) == false &&
					 string.IsNullOrWhiteSpace(leadCity) == false)
					{
						#region All contains value
						string contactFetchQuery = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='contact'>
                                          <attribute name='fullname' />
                                          <attribute name='contactid' />
                                          <order attribute='createdon' descending='true' />
                                          <filter type='and'>
                                            <filter type='or'>
                                              <filter type='or'>
                                                <condition attribute='emailaddress1' operator='eq' value='{0}' />
                                                <condition attribute='telephone1' operator='eq' value='{1}' />
                                              </filter>
                                              <filter type='and'>
                                                <condition attribute='address1_line1' operator='eq' value='{2}' />
                                                <condition attribute='address1_line2' operator='eq' value='{3}' />
                                                <condition attribute='address1_line3' operator='eq' value='{4}' />
                                                <condition attribute='address1_city' operator='eq' value='{5}' />
                                                <condition attribute='address1_postalcode' operator='eq' value='{6}' />
                                              </filter>
                                            </filter>
                                          </filter>
                                        </entity>
                                      </fetch>", leadEmail, leadTelephone, leadAddressLine1, leadAddressLine2, leadAddressLine3, leadPostalCode, leadCity);
						FetchExpression contactsFetchExpression = new FetchExpression(contactFetchQuery);

						contactEntityCollection = iOrganizationService.RetrieveMultiple(contactsFetchExpression);
						iTracingService.Trace("retrieved Contacts = {0} records", contactEntityCollection.Entities.Count);

						if (contactEntityCollection.Entities.Count > 0)//if Contact does exists with Email, Telephone and Address.
						{
							newContact = new EntityReference(contactEntityCollection.Entities[0].LogicalName, contactEntityCollection.Entities[0].Id);
						}
						#endregion
					}
					else if (string.IsNullOrWhiteSpace(leadEmail) == false &&
									 string.IsNullOrWhiteSpace(leadTelephone) == false)
					{
						#region leadEmail && leadTelephone contains value
						string contactFetchQuery = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='contact'>
                                          <attribute name='fullname' />
                                          <attribute name='contactid' />
                                          <order attribute='createdon' descending='true' />
                                          <filter type='or'>
                                                <condition attribute='emailaddress1' operator='eq' value='{0}' />
                                                <condition attribute='telephone1' operator='eq' value='{1}' />
                                          </filter>
                                        </entity>
                                      </fetch>", leadEmail, leadTelephone);
						FetchExpression contactsFetchExpression = new FetchExpression(contactFetchQuery);

						contactEntityCollection = iOrganizationService.RetrieveMultiple(contactsFetchExpression);
						iTracingService.Trace("retrieved Contacts = {0} records", contactEntityCollection.Entities.Count);

						if (contactEntityCollection.Entities.Count > 0)//if Contact does exists with Email, Telephone and Address.
						{
							newContact = new EntityReference(contactEntityCollection.Entities[0].LogicalName, contactEntityCollection.Entities[0].Id);
						}
						#endregion
					}
					else if (string.IsNullOrWhiteSpace(leadTelephone) == false &&
									 string.IsNullOrWhiteSpace(leadAddressLine1) == false &&
									 string.IsNullOrWhiteSpace(leadAddressLine3) == false &&
									 string.IsNullOrWhiteSpace(leadAddressLine2) == false &&
									 string.IsNullOrWhiteSpace(leadPostalCode) == false &&
									 string.IsNullOrWhiteSpace(leadCity) == false)
					{
						#region telephone && address contains value
						string contactFetchQuery = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='contact'>
                                          <attribute name='fullname' />
                                          <attribute name='contactid' />
                                          <order attribute='createdon' descending='true' />
                                          <filter type='or'>
                                                <condition attribute='telephone1' operator='eq' value='{0}' />
                                                <condition attribute='address1_line1' operator='eq' value='{1}' />
                                                <condition attribute='address1_line2' operator='eq' value='{2}' />
                                                <condition attribute='address1_line3' operator='eq' value='{3}' />
                                                <condition attribute='address1_city' operator='eq' value='{4}' />
                                                <condition attribute='address1_postalcode' operator='eq' value='{5}' />
                                          </filter>
                                        </entity>
                                      </fetch>", leadTelephone, leadAddressLine1, leadAddressLine2, leadAddressLine3, leadCity, leadPostalCode);
						FetchExpression contactsFetchExpression = new FetchExpression(contactFetchQuery);

						contactEntityCollection = iOrganizationService.RetrieveMultiple(contactsFetchExpression);
						iTracingService.Trace("retrieved Contacts = {0} records", contactEntityCollection.Entities.Count);

						if (contactEntityCollection.Entities.Count > 0)//if Contact does exists with Email, Telephone and Address.
						{
							newContact = new EntityReference(contactEntityCollection.Entities[0].LogicalName, contactEntityCollection.Entities[0].Id);
						}
						#endregion
					}
					else if (string.IsNullOrWhiteSpace(leadTelephone) == false &&
									 string.IsNullOrWhiteSpace(leadAddressLine1) == false &&
									 string.IsNullOrWhiteSpace(leadAddressLine3) == false &&
									 string.IsNullOrWhiteSpace(leadAddressLine2) == false &&
									 string.IsNullOrWhiteSpace(leadPostalCode) == false &&
									 string.IsNullOrWhiteSpace(leadCity) == false)
					{
						#region leadEmail && address contains value
						string contactFetchQuery = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='contact'>
                                          <attribute name='fullname' />
                                          <attribute name='contactid' />
                                          <order attribute='createdon' descending='true' />
                                          <filter type='or'>
                                                <condition attribute='emailaddress1' operator='eq' value='{0}' />
                                                <condition attribute='address1_line1' operator='eq' value='{1}' />
                                                <condition attribute='address1_line2' operator='eq' value='{2}' />
                                                <condition attribute='address1_line3' operator='eq' value='{3}' />
                                                <condition attribute='address1_city' operator='eq' value='{4}' />
                                                <condition attribute='address1_postalcode' operator='eq' value='{5}' />
                                          </filter>
                                        </entity>
                                      </fetch>", leadEmail, leadAddressLine1, leadAddressLine2, leadAddressLine3, leadCity, leadPostalCode);
						FetchExpression contactsFetchExpression = new FetchExpression(contactFetchQuery);

						contactEntityCollection = iOrganizationService.RetrieveMultiple(contactsFetchExpression);
						iTracingService.Trace("retrieved Contacts = {0} records", contactEntityCollection.Entities.Count);

						if (contactEntityCollection.Entities.Count > 0)//if Contact does exists with Email, Telephone and Address.
						{
							newContact = new EntityReference(contactEntityCollection.Entities[0].LogicalName, contactEntityCollection.Entities[0].Id);
						}
						#endregion
					}
					else if (string.IsNullOrWhiteSpace(leadEmail) == false)
					{
						#region email contains value
						string contactFetchQuery = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='contact'>
                                          <attribute name='fullname' />
                                          <attribute name='contactid' />
                                          <order attribute='createdon' descending='true' />
                                          <filter>
                                            <condition attribute='emailaddress1' operator='eq' value='{0}' />
                                          </filter>
                                        </entity>
                                      </fetch>", leadEmail);
						FetchExpression contactsFetchExpression = new FetchExpression(contactFetchQuery);

						contactEntityCollection = iOrganizationService.RetrieveMultiple(contactsFetchExpression);
						iTracingService.Trace("retrieved Contacts = {0} records", contactEntityCollection.Entities.Count);

						if (contactEntityCollection.Entities.Count > 0)//if Contact does exists with Email, Telephone and Address.
						{
							newContact = new EntityReference(contactEntityCollection.Entities[0].LogicalName, contactEntityCollection.Entities[0].Id);
						}
						#endregion
					}
					else if (string.IsNullOrWhiteSpace(leadTelephone) == false)
					{
						#region telephone contains value
						string contactFetchQuery = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='contact'>
                                          <attribute name='fullname' />
                                          <attribute name='contactid' />
                                          <order attribute='createdon' descending='true' />
                                          <filter>
                                            <condition attribute='telephone1' operator='eq' value='{0}' />
                                          </filter>
                                        </entity>
                                      </fetch>", leadTelephone);
						FetchExpression contactsFetchExpression = new FetchExpression(contactFetchQuery);

						contactEntityCollection = iOrganizationService.RetrieveMultiple(contactsFetchExpression);
						iTracingService.Trace("retrieved Contacts = {0} records", contactEntityCollection.Entities.Count);

						if (contactEntityCollection.Entities.Count > 0)//if Contact does exists with Email, Telephone and Address.
						{
							newContact = new EntityReference(contactEntityCollection.Entities[0].LogicalName, contactEntityCollection.Entities[0].Id);
						}
						#endregion
					}
					else if (string.IsNullOrWhiteSpace(leadAddressLine1) == false &&
									 string.IsNullOrWhiteSpace(leadAddressLine3) == false &&
									 string.IsNullOrWhiteSpace(leadAddressLine2) == false &&
									 string.IsNullOrWhiteSpace(leadPostalCode) == false &&
									 string.IsNullOrWhiteSpace(leadCity) == false)
					{
						#region address contains value
						string contactFetchQuery = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='contact'>
                                          <attribute name='fullname' />
                                          <attribute name='contactid' />
                                          <order attribute='createdon' descending='true' />
                                          <filter type='or'>
                                            <condition attribute='address1_line1' operator='eq' value='{0}' />
                                            <condition attribute='address1_line2' operator='eq' value='{1}' />
                                            <condition attribute='address1_line3' operator='eq' value='{2}' />
                                            <condition attribute='address1_city' operator='eq' value='{3}' />
                                            <condition attribute='address1_postalcode' operator='eq' value='{4}' />
                                          </filter>
                                        </entity>
                                      </fetch>", leadAddressLine1, leadAddressLine2, leadAddressLine3, leadCity, leadPostalCode);
						FetchExpression contactsFetchExpression = new FetchExpression(contactFetchQuery);

						contactEntityCollection = iOrganizationService.RetrieveMultiple(contactsFetchExpression);
						iTracingService.Trace("retrieved Contacts = {0} records", contactEntityCollection.Entities.Count);

						if (contactEntityCollection.Entities.Count > 0)//if Contact does exists with Email, Telephone and Address.
						{
							newContact = new EntityReference(contactEntityCollection.Entities[0].LogicalName, contactEntityCollection.Entities[0].Id);
						}
						#endregion
					}
					#endregion
				}
			}
			return newContact;
		}
		#endregion

		#region CreateNewContact 
		internal static EntityReference CreateNewContact(string leadName,
																						 string leadEmail,
																						 string leadTelephone,
																						 string leadAddressLine1,
																						 string leadAddressLine2,
																						 string leadAddressLine3,
																						 string leadPostalCode,
																						 string leadCity,
																						 IOrganizationService iOrganizationService,
																						 ITracingService iTracingService)
		{
			iTracingService.Trace("Entered in Method CreateNewContact");
			iTracingService.Trace("Creating a new Contact from lead data");
			Entity contact = new Entity("contact");
			contact.Attributes.Add("lastname", leadName);
			contact.Attributes.Add("emailaddress1", leadEmail);
			contact.Attributes.Add("telephone1", leadTelephone);
			contact.Attributes.Add("address1_line1", leadAddressLine1);
			contact.Attributes.Add("address1_line2", leadAddressLine2);
			contact.Attributes.Add("address1_line3", leadAddressLine3);
			contact.Attributes.Add("address1_postalcode", leadPostalCode);
			contact.Attributes.Add("address1_city", leadCity);
			Guid newContactId = iOrganizationService.Create(contact);
			iTracingService.Trace("New Contact record created");
			EntityReference newContact = new EntityReference("contact", newContactId);
			return newContact;
		}
		#endregion

		#region CreateStringEmptyIFNull 
		internal static void ReplaceWithBlankIfStringContainsNull(ref string attribute)
		{
			if (string.IsNullOrWhiteSpace(attribute) == true)
			{ attribute = string.Empty; }
		}
		#endregion

		#region UpdateErrorOnPreCreate
		private void UpdateErrorOnPreCreate(Entity entity,
																 string errorMessage,
																 IOrganizationService iOrganizationService,
																 ITracingService iTracingService)
		{
			iTracingService.Trace("Entered in method UpdateErrorOnPreCreate.");
			if (entity != null &&
					string.IsNullOrWhiteSpace(errorMessage) == false)//Entity context is not null and errorCode is not empty
			{
				iTracingService.Trace("Entity context is not null");
				iTracingService.Trace("Error Message = {0}", errorMessage);
				//Updating the field "Message" on Lead entity with the Error Message
				CRMData.AddAttribute(entity, "resi_message", errorMessage);
				iTracingService.Trace("Error Message updated.");
			}
		}
		#endregion

		#endregion
	}

	/// <summary>
	/// Lead Class
	/// </summary>
	/// <seealso cref="Microsoft.Xrm.Sdk.IPlugin" />
	public class PostLeadCreate : IPlugin
	{
		/// <summary>
		/// Executes plug-in code in response to an event.
		/// </summary>
		/// <param name="serviceProvider">Type: Returns_IServiceProvider. A container for service objects. Contains references to the plug-in execution context (<see cref="T:Microsoft.Xrm.Sdk.IPluginExecutionContext"></see>), tracing service (<see cref="T:Microsoft.Xrm.Sdk.ITracingService"></see>), organization service (<see cref="T:Microsoft.Xrm.Sdk.IOrganizationServiceFactory"></see>), and notification service (<see cref="T:Microsoft.Xrm.Sdk.IServiceEndpointNotificationService"></see>).</param>
		/// <exception cref="InvalidPluginExecutionException">
		/// Invalid 
		/// or
		/// An error occurred in the PostLeadCreate plug-in. + ex.Message
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
			if (CRMData.ValidateTargetAsEntity("lead", context))
			{
				try
				{
					iTracingService.Trace("Context validating for proceed futher");
					bool proceedFurtherInPostPlugin = (bool)context.SharedVariables["ProceedFurtherInPostPlugin"];
					if (proceedFurtherInPostPlugin == true)
					{
						iTracingService.Trace("Context contains target entity");
						// Obtain the target entity from the input parameters.
						Entity lead = (Entity)context.InputParameters["Target"];
						iTracingService.Trace("Context fetched");
						string leadName = CRMData.GetAttributeValue<string>(lead, null, "fullname");
						iTracingService.Trace("leadName = {0}", leadName);

						OptionSetValue reglesDeDispatch = null;
						bool isContactExists = false;
						EntityReference leadParentContact = null;
						Entity opportunity = null;
						EntityReference zoneDeChalandise = null;
						EntityReference ownerOpportunity = null;

						iTracingService.Trace("lead is not null");
						reglesDeDispatch = CRMData.GetAttributeValue<OptionSetValue>(lead, null, "resi_reglesdedispatchtype");
						leadParentContact = CRMData.GetAttributeValue<EntityReference>(lead, null, "parentcontactid");

						if (context.SharedVariables.ContainsKey("isContactExists"))
						{
							iTracingService.Trace("context contains the key [isContactExists]");
							isContactExists = (bool)context.SharedVariables["isContactExists"];
							iTracingService.Trace("isContactExists ={0}", isContactExists.ToString());
						}

						if (isContactExists == true)
						{
							#region isContactExists == true 
							if (leadParentContact != null)
							{
								opportunity = Recherche_Demande(reglesDeDispatch,
																								leadParentContact,
																								iOrganizationService,
																								iTracingService);
							}
							if (opportunity != null)
							{
								zoneDeChalandise = CRMData.GetAttributeValue<EntityReference>(opportunity, null, "resi_zonedechalandiseid");
								ownerOpportunity = CRMData.GetAttributeValue<EntityReference>(opportunity, null, "ownerid");

								if (zoneDeChalandise != null && ownerOpportunity != null)
								{
									QualifyLeadAndUpdateOpportunity(lead,
																									zoneDeChalandise,
																									ownerOpportunity,
																									false,
																									iOrganizationService,
																									iTracingService);
								}
							}
							else
							{
								ValidatingCasesBasedOnRegleDeDispatch(lead,
																										isContactExists,
																										reglesDeDispatch,
																										iOrganizationService,
																										iTracingService);

							}
							#endregion
						}
						else
						{
							#region isContactExists == false 
							ValidatingCasesBasedOnRegleDeDispatch(lead,
																										isContactExists,
																										reglesDeDispatch,
																										iOrganizationService,
																										iTracingService);
							#endregion
						}
					}
				}
				catch (FaultException<OrganizationServiceFault> ex)
				{
					throw new InvalidPluginExecutionException("An error occurred in the PreLeadCreate plug-in." + ex.Message);
				}
			}
		}
		#endregion

		#region Private Methods

		#region ValidatingCasesBasedOnRegleDeDispatch
		private void ValidatingCasesBasedOnRegleDeDispatch(Entity lead,
																											 bool isNewContact,
																											 OptionSetValue reglesDeDispatch,
																											 IOrganizationService iOrganizationService,
																											 ITracingService iTracingService)
		{
			if (lead != null)
			{
				string codeDeAgenceDuClient = "";
				string postalCode = "";
				string nomDuProgramme = "";
				Entity zoneEntity = null;
				EntityReference zoneDeChalandiseOwner = null;
				EntityReference zoneDeChalandise = null;
				Guid zoneDeChalandiseId = new Guid();
				string errorMessage = string.Empty;
				EntityReference contactParentAccount = null;
				Entity account = null;
				//// Entity contact = null;
				EntityReference newContact = null;
				bool isCodeAgenceFoundInAccount = false;
				string accountNumber = string.Empty;
				string leadName = CRMData.GetAttributeValue<string>(lead, null, "fullname");
				PreLeadCreate.ReplaceWithBlankIfStringContainsNull(ref leadName);
				string leadEmail = CRMData.GetAttributeValue<string>(lead, null, "emailaddress1");
				PreLeadCreate.ReplaceWithBlankIfStringContainsNull(ref leadEmail);
				string leadTelePhone = CRMData.GetAttributeValue<string>(lead, null, "telephone1");
				PreLeadCreate.ReplaceWithBlankIfStringContainsNull(ref leadTelePhone);
				string leadAddressLine1 = CRMData.GetAttributeValue<string>(lead, null, "address1_line1");
				PreLeadCreate.ReplaceWithBlankIfStringContainsNull(ref leadAddressLine1);
				string leadAddressLine2 = CRMData.GetAttributeValue<string>(lead, null, "address1_line2");
				PreLeadCreate.ReplaceWithBlankIfStringContainsNull(ref leadAddressLine2);
				string leadAddressLine3 = CRMData.GetAttributeValue<string>(lead, null, "address1_line3");
				PreLeadCreate.ReplaceWithBlankIfStringContainsNull(ref leadAddressLine3);
				string leadCity = CRMData.GetAttributeValue<string>(lead, null, "address1_city");
				PreLeadCreate.ReplaceWithBlankIfStringContainsNull(ref leadCity);
				string leadPostalCode = CRMData.GetAttributeValue<string>(lead, null, "address1_postalcode");
				PreLeadCreate.ReplaceWithBlankIfStringContainsNull(ref leadPostalCode);
				string leadVendorId = CRMData.GetAttributeValue<string>(lead, null, "resi_vendorbddfid");
				bool isZoneFound = false;
				bool isContactToCreateOnLeadQualify = false;

				if (isNewContact == false)
				{
					isContactToCreateOnLeadQualify = true;
				}


				#region reglesDeDispatch == 3 (Code Agence)
				if (reglesDeDispatch != null &&
					reglesDeDispatch.Value == 3)
				{
					iTracingService.Trace("reglesDeDispatch is Code Agence");

					codeDeAgenceDuClient = CRMData.GetAttributeValue<string>(lead, null, "resi_bddfagencycode");
					if (string.IsNullOrWhiteSpace(codeDeAgenceDuClient) == false)
					{
						#region codeDeAgenceDuClient is not empty
						iTracingService.Trace("codeDeAgenceDuClient is not null");

						iTracingService.Trace("Fetch the Account record where Code is equal to codeDeAgenceDuClient");
						// Fetch Related "Account" entity records
						QueryExpression accountQueryExpression = new QueryExpression("account");
						accountQueryExpression.ColumnSet = new ColumnSet("accountid", "resi_sales_mgt_area", "name");
						accountQueryExpression.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, codeDeAgenceDuClient);
						EntityCollection accountEntityCollection = iOrganizationService.RetrieveMultiple(accountQueryExpression);

						iTracingService.Trace("retrieved accounts = {0} records", accountEntityCollection.Entities.Count);

						if (accountEntityCollection.Entities.Count > 0)
						{
							#region Account Records fetched with "accountnumber" = CodeDeAgence
							iTracingService.Trace("Account Records fetched with accountnumber = CodeDeAgence");

							if (string.IsNullOrWhiteSpace(leadVendorId) == false)
							{
								#region Field "Nom du commercial BDDF" is not null on lead entity.
								iTracingService.Trace("Field Nom du commercial BDDF is not null on lead entity");
								// Fetch all "Contact" entity records that have "resi_matricule" = leadVendorId
								QueryExpression contactQueryExpression = new QueryExpression("contact");
								contactQueryExpression.ColumnSet = new ColumnSet("contactid", "parentcustomerid");
								contactQueryExpression.Criteria.AddCondition("resi_matricule", ConditionOperator.Equal, leadVendorId);
								contactQueryExpression.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
								EntityCollection contactEntityCollection = iOrganizationService.RetrieveMultiple(contactQueryExpression);

								iTracingService.Trace("retrieved Contacts = {0} records", contactEntityCollection.Entities.Count);
								if (contactEntityCollection.Entities.Count > 0)
								{
									iTracingService.Trace("Field Nom du commercial BDDF is not null on lead entity");
									#region "Contact" entity records exists that contains "resi_matricule" = leadVendorId
									iTracingService.Trace("Contact entity records exists that contains resi_matricule = leadVendorId");
									newContact = new EntityReference(contactEntityCollection.Entities[0].LogicalName, contactEntityCollection.Entities[0].Id);

									#region Updating Lead with Parentcontact 
									Entity leadEntity = new Entity(lead.LogicalName);
									leadEntity.Id = lead.Id;
									CRMData.AddAttribute<EntityReference>(leadEntity, "parentcontactid", newContact);
									iOrganizationService.Update(leadEntity);
									#endregion

									contactParentAccount = CRMData.GetAttributeValue<EntityReference>(contactEntityCollection.Entities[0], null, "parentcustomerid");

									if (contactParentAccount != null)
									{
										iTracingService.Trace("contactParentAccount != null");
										account = iOrganizationService.Retrieve(contactParentAccount.LogicalName, contactParentAccount.Id, new ColumnSet("accountnumber"));
										accountNumber = CRMData.GetAttributeValue<string>(account, null, "accountnumber");

										if (string.IsNullOrWhiteSpace(accountNumber) == false)
										{
											iTracingService.Trace("accountNumber = {0}", accountNumber);
											if (accountNumber == codeDeAgenceDuClient)
											{
												iTracingService.Trace("accountNumber == codeDeAgenceDuClient");
												isCodeAgenceFoundInAccount = true;
											}
										}
									}
									if (isCodeAgenceFoundInAccount == true)
									{
										QualifyLeadAndUpdateOpportunityParentContact(lead, newContact, account, iOrganizationService, iTracingService);
									}
									else
									{
										#region Parent Account of the Conseiller bancaire does not contains the "accountnumber" = codeDeAgenceDuClient
										iTracingService.Trace("Parent Account of the Conseiller bancaire does not contains the 'accountnumber' = codeDeAgenceDuClient");
										errorMessage = "Le code d'agence du conseiller bancaire présent dans le CRM est différent de celui fourni par le fichier Source";
										UpdateErrorOnPostCreate(lead, errorMessage, iOrganizationService, iTracingService);
										#endregion
									}
									#endregion
								}
								else
								{
									#region "Contact" entity records does not exists that contains "resi_matricule" = leadVendorId
									newContact = PreLeadCreate.CreateNewContact(leadName,
																															leadEmail,
																															leadTelePhone,
																															leadAddressLine1,
																															leadAddressLine2,
																															leadAddressLine3,
																															leadPostalCode,
																															leadCity,
																															iOrganizationService,
																															iTracingService);

									#region Updating Lead with Parentcontact 
									Entity leadEntity = new Entity(lead.LogicalName);
									leadEntity.Id = lead.Id;
									CRMData.AddAttribute<EntityReference>(leadEntity, "parentcontactid", newContact);
									iOrganizationService.Update(leadEntity);
									#endregion

									account = accountEntityCollection.Entities[0];
									QualifyLeadAndUpdateOpportunityParentContact(lead, newContact, account, iOrganizationService, iTracingService);
									#endregion
								}
								#endregion
							}
							else
							{
								#region Field Nom du commercial BDDF is null on lead entity
								iTracingService.Trace("Field Nom du commercial BDDF is null on lead entity");
								errorMessage = "Le champ Commercial n'a pas été spécifié lors de l'import";
								UpdateErrorOnPostCreate(lead, errorMessage, iOrganizationService, iTracingService);
								#endregion
							}
							#endregion
						}
						else//Account Records does not exists with "accountnumber" = CodeDeAgence
						{
							#region Account Records does not exists with "accountnumber" = CodeDeAgence
							iTracingService.Trace("Account Records does not exists with accountnumber = CodeDeAgence");
							errorMessage = "Agence spécifiée n'existe pas dans le CRM";
							UpdateErrorOnPostCreate(lead, errorMessage, iOrganizationService, iTracingService);
							#endregion
						}
						#endregion
					}
					else //codeDeAgenceDuClient is empty
					{
						#region codeDeAgenceDuClient is empty
						iTracingService.Trace("codeDeAgenceDuClient is empty");
						errorMessage = "Le Code Agence n'est pas present dans le fichier source";
						UpdateErrorOnPostCreate(lead, errorMessage, iOrganizationService, iTracingService);
						#endregion
					}
				}
				#endregion

				#region reglesDeDispatch == 2 (Code Postale)
				if (reglesDeDispatch != null &&
					reglesDeDispatch.Value == 2)
				{
					iTracingService.Trace("reglesDeDispatch is Code Postale");
					postalCode = CRMData.GetAttributeValue<string>(lead, null, "address1_postalcode");
					if (string.IsNullOrWhiteSpace(postalCode) == false)
					{
						iTracingService.Trace("postalCode is not null");
						// Fetch "Code Postal" entity records 
						QueryExpression codePostalQueryExpression = new QueryExpression("resi_codepostal");
						codePostalQueryExpression.ColumnSet = new ColumnSet("resi_zonedechalandiseid");
						codePostalQueryExpression.Criteria.AddCondition("resi_name", ConditionOperator.Equal, postalCode);
						EntityCollection codePostalEntityCollection = iOrganizationService.RetrieveMultiple(codePostalQueryExpression);

						iTracingService.Trace("retrieved Postal Code = {0} records", codePostalEntityCollection.Entities.Count);

						if (codePostalEntityCollection.Entities.Count > 0)
						{
							zoneDeChalandise = CRMData.GetAttributeValue<EntityReference>(codePostalEntityCollection.Entities[0], null, "resi_zonedechalandiseid");
							if (zoneDeChalandise != null)
							{
								isZoneFound = true;
								iTracingService.Trace("zoneDeChalandise on Code Postale entity is not null");
								zoneEntity = iOrganizationService.Retrieve(zoneDeChalandise.LogicalName, zoneDeChalandise.Id, new ColumnSet("ownerid"));
								if (zoneEntity != null)
								{
									iTracingService.Trace("zoneEntity retrieved");
									zoneDeChalandiseOwner = CRMData.GetAttributeValue<EntityReference>(zoneEntity, null, "ownerid");
									if (zoneDeChalandiseOwner != null)
									{
										iTracingService.Trace("zoneDeChalandiseOwner fetched from zoneEntity retrieved");
										iTracingService.Trace("Calling method QualifyLeadAndUpdateOpportunity");
										QualifyLeadAndUpdateOpportunity(lead,
																										zoneDeChalandise,
																										zoneDeChalandiseOwner,
																										isContactToCreateOnLeadQualify,
																										iOrganizationService,
																										iTracingService);
									}
								}
							}
						}
					}

					if (isZoneFound == false)
					{
						// Fetch "Team" entity record for team "Team Sans CP "
						QueryExpression teamQueryExpression = new QueryExpression("team");
						teamQueryExpression.ColumnSet = new ColumnSet("teamid");
						teamQueryExpression.Criteria.AddCondition("name", ConditionOperator.Equal, "Team Sans CP");
						EntityCollection teamnEntityCollection = iOrganizationService.RetrieveMultiple(teamQueryExpression);

						if (teamnEntityCollection.Entities.Count > 0)
						{
							iTracingService.Trace("teamnEntityCollection.Entities.Count:{0}", teamnEntityCollection.Entities.Count);
							Guid teamSansCPId = CRMData.GetAttributeValue<Guid>(teamnEntityCollection.Entities[0], null, "teamid");
							iTracingService.Trace("teamSansCPId:{0}", teamSansCPId.ToString());

							// Fetch "Zone De Chalandise" entity record where owner = Teams Sans CP 
							QueryExpression zoneDeChalandiseQueryExpression = new QueryExpression("resi_salesarea");
							zoneDeChalandiseQueryExpression.ColumnSet = new ColumnSet("resi_salesareaid", "ownerid");
							zoneDeChalandiseQueryExpression.Criteria.AddCondition("ownerid", ConditionOperator.Equal, teamSansCPId);
							EntityCollection zoneDeChalandiseEntityCollection = iOrganizationService.RetrieveMultiple(zoneDeChalandiseQueryExpression);

							iTracingService.Trace("retrieved zoneDeChalandise = {0} records", zoneDeChalandiseEntityCollection.Entities.Count);

							if (zoneDeChalandiseEntityCollection.Entities.Count > 0)
							{
								iTracingService.Trace("zoneDeChalandiseEntityCollection fetched");
								zoneDeChalandiseId = CRMData.GetAttributeValue<Guid>(zoneDeChalandiseEntityCollection.Entities[0], null, "resi_salesareaid");
								zoneDeChalandiseOwner = CRMData.GetAttributeValue<EntityReference>(zoneDeChalandiseEntityCollection.Entities[0], null, "ownerid");
								if (zoneDeChalandiseId != Guid.Empty && zoneDeChalandiseOwner != null)
								{
									iTracingService.Trace("zoneDeChalandiseId and zoneDeChalandiseOwner is not null");
									zoneDeChalandise = new EntityReference("resi_salesarea", zoneDeChalandiseId);
									if (zoneDeChalandise != null)
									{
										iTracingService.Trace("zoneDeChalandise is not null");
										iTracingService.Trace("Calling method QualifyLeadAndUpdateOpportunity");
										QualifyLeadAndUpdateOpportunity(lead,
																										zoneDeChalandise,
																										zoneDeChalandiseOwner,
																										isContactToCreateOnLeadQualify,
																										iOrganizationService,
																										iTracingService);
									}
								}
							}
						}
					}
				}
				#endregion

				#region reglesDeDispatch == 1 (Programme)
				if (reglesDeDispatch != null &&
					reglesDeDispatch.Value == 1)
				{
					iTracingService.Trace("reglesDeDispatch is Code Programme");
					nomDuProgramme = CRMData.GetAttributeValue<string>(lead, null, "resi_programname");
					if (string.IsNullOrWhiteSpace(nomDuProgramme) == false)
					{
						iTracingService.Trace("nomDuProgramme is not null");
						// Fetch "Programme" entity records 
						QueryExpression programmeQueryExpression = new QueryExpression("resi_programme");
						programmeQueryExpression.ColumnSet = new ColumnSet("resi_zonedechalandiseid");
						programmeQueryExpression.Criteria.AddCondition("resi_name", ConditionOperator.Equal, nomDuProgramme);
						EntityCollection programmeEntityCollection = iOrganizationService.RetrieveMultiple(programmeQueryExpression);

						iTracingService.Trace("retrieved Programme = {0} records", programmeEntityCollection.Entities.Count);

						if (programmeEntityCollection.Entities.Count > 0)
						{
							zoneDeChalandise = CRMData.GetAttributeValue<EntityReference>(programmeEntityCollection.Entities[0], null, "resi_zonedechalandiseid");
							if (zoneDeChalandise != null)
							{
								isZoneFound = true;
								iTracingService.Trace("zoneDeChalandise on Programme entity is not null");
								zoneEntity = iOrganizationService.Retrieve(zoneDeChalandise.LogicalName, zoneDeChalandise.Id, new ColumnSet("ownerid"));
								if (zoneEntity != null)
								{
									iTracingService.Trace("zoneEntity retrieved");
									zoneDeChalandiseOwner = CRMData.GetAttributeValue<EntityReference>(zoneEntity, null, "ownerid");
									if (zoneDeChalandiseOwner != null)
									{
										iTracingService.Trace("zoneDeChalandiseOwner fetched from zoneEntity retrieved");
										iTracingService.Trace("Calling method QualifyLeadAndUpdateOpportunity");
										QualifyLeadAndUpdateOpportunity(lead,
																										zoneDeChalandise,
																										zoneDeChalandiseOwner,
																										isContactToCreateOnLeadQualify,
																										iOrganizationService,
																										iTracingService);
									}
								}
							}
						}
					}

					if (isZoneFound == false)
					{

						// Fetch "Team" entity record for team "Team Sans Programme "
						QueryExpression teamQueryExpression = new QueryExpression("team");
						teamQueryExpression.ColumnSet = new ColumnSet("teamid");
						teamQueryExpression.Criteria.AddCondition("name", ConditionOperator.Equal, "Team Sans Programme");
						EntityCollection teamnEntityCollection = iOrganizationService.RetrieveMultiple(teamQueryExpression);

						if (teamnEntityCollection.Entities.Count > 0)
						{
							iTracingService.Trace("teamnEntityCollection.Entities.Count:{0}", teamnEntityCollection.Entities.Count);
							Guid teamSansProgrammeId = CRMData.GetAttributeValue<Guid>(teamnEntityCollection.Entities[0], null, "teamid");
							iTracingService.Trace("teamSansCPId:{0}", teamSansProgrammeId.ToString());

							// Fetch "Zone De Chalandise" entity record where owner = Teams Sans Programme 
							QueryExpression zoneDeChalandiseQueryExpression = new QueryExpression("resi_salesarea");
							zoneDeChalandiseQueryExpression.ColumnSet = new ColumnSet("resi_salesareaid", "ownerid");
							zoneDeChalandiseQueryExpression.Criteria.AddCondition("ownerid", ConditionOperator.Equal, teamSansProgrammeId);
							EntityCollection zoneDeChalandiseEntityCollection = iOrganizationService.RetrieveMultiple(zoneDeChalandiseQueryExpression);

							iTracingService.Trace("retrieved zoneDeChalandise = {0} records", zoneDeChalandiseEntityCollection.Entities.Count);

							if (zoneDeChalandiseEntityCollection.Entities.Count > 0)
							{
								iTracingService.Trace("zoneDeChalandiseEntityCollection fetched");
								zoneDeChalandiseId = CRMData.GetAttributeValue<Guid>(zoneDeChalandiseEntityCollection.Entities[0], null, "resi_salesareaid");
								zoneDeChalandiseOwner = CRMData.GetAttributeValue<EntityReference>(zoneDeChalandiseEntityCollection.Entities[0], null, "ownerid");
								if (zoneDeChalandiseId != Guid.Empty && zoneDeChalandiseOwner != null)
								{
									iTracingService.Trace("zoneDeChalandiseId and zoneDeChalandiseOwner is not null");
									zoneDeChalandise = new EntityReference("resi_salesarea", zoneDeChalandiseId);
									if (zoneDeChalandise != null)
									{
										iTracingService.Trace("zoneDeChalandise is not null");
										iTracingService.Trace("Calling method QualifyLeadAndUpdateOpportunity");
										QualifyLeadAndUpdateOpportunity(lead,
																										zoneDeChalandise,
																										zoneDeChalandiseOwner,
																										isContactToCreateOnLeadQualify,
																										iOrganizationService,
																										iTracingService);
									}
								}
							}
						}
					}
				}
				#endregion

			}
		}
		#endregion

		#region QualifyLeadAndUpdateOpportunity
		public static void QualifyLeadAndUpdateOpportunity(Entity lead,
																								 EntityReference zoneDeChalandise,
																								 EntityReference owner,
																								 bool isContactToCreateOnLeadQualify,
																								 IOrganizationService iOrganizationService,
																								 ITracingService iTracingService)
		{
			if (lead != null)
			{
				iTracingService.Trace("Entered in function QualifyLeadAndUpdateOpportunity");
				EntityReference parentAccountEntity = null;
				EntityReference parentContactEntity = null;
				EntityReference opportunityEntity = null;

				var qualifyLead = new QualifyLeadRequest
				{
					CreateAccount = true,
					CreateContact = isContactToCreateOnLeadQualify,
					CreateOpportunity = true,
					LeadId = new EntityReference(lead.LogicalName, lead.Id),
					Status = new OptionSetValue(3)
				};
				iTracingService.Trace("qualify lead request populated");
				var qualifyIntoAccountContactRes = (QualifyLeadResponse)iOrganizationService.Execute(qualifyLead);
				iTracingService.Trace("qualify lead request executed");

				foreach (EntityReference created in qualifyIntoAccountContactRes.CreatedEntities)
				{
					//Check if Account record is created when lead qualified.
					if (created.LogicalName == "account")
					{
						iTracingService.Trace("parent account is created on lead qualify");
						parentAccountEntity = new EntityReference("account", created.Id);
					}

					//Check if Contact record is created when lead qualified.
					if (created.LogicalName == "contact")
					{
						iTracingService.Trace("parent contact is created on lead qualify");
						parentContactEntity = new EntityReference("contact", created.Id);
					}

					//Check if Oppo record is created when lead qualified.
					if (created.LogicalName == "opportunity")
					{
						iTracingService.Trace("opportunity is created on lead qualify");
						opportunityEntity = new EntityReference("opportunity", created.Id);
					}
				}

				if (opportunityEntity != null &&
						owner != null)
				{
					iTracingService.Trace("Updating the zoneDeChalandise field in Opportunity Record created");
					//Update the Opportunity record created by qualifying the lead
					Entity opportunity = new Entity("opportunity");
					opportunity.Id = opportunityEntity.Id;
					opportunity.Attributes.Add("resi_zonedechalandiseid", zoneDeChalandise);
					iOrganizationService.Update(opportunity);
					iTracingService.Trace("zoneDeChalandise field in the Opportunity created on qualifying lead updated.");

					#region Assigning the Opportunity created on qualifying lead
					CRMData.AssignEntityRecord(owner, opportunityEntity, iOrganizationService);
					#endregion
				}
			}
		}
		#endregion

		#region QualifyLeadAndUpdateOpportunityParentContact
		private void QualifyLeadAndUpdateOpportunityParentContact(Entity lead,
																															EntityReference parentContact,
																															Entity account,
																															IOrganizationService iOrganizationService,
																															ITracingService iTracingService)
		{
			if (lead != null)
			{
				iTracingService.Trace("Entered in function QualifyLeadAndUpdateOpportunityParentContact");
				EntityReference parentAccountEntity = null;
				EntityReference parentContactEntity = null;
				EntityReference opportunityEntity = null;
				EntityReference zoneDeChalandise = null;
				EntityReference zoneDeChalandiseSansAgence = new EntityReference("resi_salesarea", new Guid("FDE0F66F-E6F8-E611-8113-3863BB34D9E0"));

				var qualifyLead = new QualifyLeadRequest
				{
					CreateAccount = true,
					CreateContact = false,
					CreateOpportunity = true,
					LeadId = new EntityReference(lead.LogicalName, lead.Id),
					Status = new OptionSetValue(3)
				};
				iTracingService.Trace("qualify lead request populated");
				var qualifyIntoAccountContactRes = (QualifyLeadResponse)iOrganizationService.Execute(qualifyLead);
				iTracingService.Trace("qualify lead request executed");

				foreach (EntityReference created in qualifyIntoAccountContactRes.CreatedEntities)
				{
					//Check if Account record is created when lead qualified.
					if (created.LogicalName == "account")
					{
						iTracingService.Trace("parent account is created on lead qualify");
						parentAccountEntity = new EntityReference("account", created.Id);
					}

					//Check if Contact record is created when lead qualified.
					if (created.LogicalName == "contact")
					{
						iTracingService.Trace("parent contact is created on lead qualify");
						parentContactEntity = new EntityReference("contact", created.Id);
					}

					//Check if Oppo record is created when lead qualified.
					if (created.LogicalName == "opportunity")
					{
						iTracingService.Trace("opportunity is created on lead qualify");
						opportunityEntity = new EntityReference("opportunity", created.Id);
					}
				}

				if (account != null)
				{
					iTracingService.Trace("qualify lead request populated");
					string accountName = CRMData.GetAttributeValue<string>(account, null, "name");
					iTracingService.Trace("account name = {0}", accountName);
					zoneDeChalandise = CRMData.GetAttributeValue<EntityReference>(account, null, "resi_sales_mgt_area");
					if (zoneDeChalandise != null)
					{
						iTracingService.Trace("Updating the parentContact and ZoneDeChalandise fields in Opportunity Record created");
						//Update the Opportunity record created by qualifying the lead
						Entity opportunity = new Entity("opportunity");
						opportunity.Id = opportunityEntity.Id;
						opportunity.Attributes.Add("parentcontactid", parentContact);
						opportunity.Attributes.Add("resi_zonedechalandiseid", zoneDeChalandise);
						iOrganizationService.Update(opportunity);
						iTracingService.Trace("parentcontactid and ZoneDeChalandise fields updated in the Opportunity created on qualifying lead updated.");
					}
					else
					{
						iTracingService.Trace("Updating the parentContact field in Opportunity Record created");
						//Update the Opportunity record created by qualifying the lead
						Entity opportunity = new Entity("opportunity");
						opportunity.Id = opportunityEntity.Id;
						opportunity.Attributes.Add("parentcontactid", parentContact);
						iOrganizationService.Update(opportunity);
						iTracingService.Trace("parentcontactid updated  in the Opportunity created on qualifying lead updated.");

						#region Populating error on Message field as the 'Zone De Chalandise' is null on account.
						iTracingService.Trace("Populating error on Message field as the 'Zone De Chalandise' is null on account.");
						string errorMessage = "La valeur était absente de la source ou n’a pas été trouvé dans le CRM";
						UpdateErrorOnPostCreate(lead, errorMessage, iOrganizationService, iTracingService);
						#endregion
					}
				}

				//if (opportunityEntity != null &&
				//    parentContact != null)
				//{
				//  iTracingService.Trace("Updating the parentContact field in Opportunity Record created");
				//  //Update the Opportunity record created by qualifying the lead
				//  Entity opportunity = new Entity("opportunity");
				//  opportunity.Id = opportunityEntity.Id;
				//  opportunity.Attributes.Add("parentcontactid", parentContact);
				//  iOrganizationService.Update(opportunity);
				//  iTracingService.Trace("parentcontactid field in the Opportunity created on qualifying lead updated.");
				//}
			}
		}
		#endregion

		#region UpdateErrorOnPostCreate
		private void UpdateErrorOnPostCreate(Entity entity,
																 string errorMessage,
																 IOrganizationService iOrganizationService,
																 ITracingService iTracingService)
		{
			iTracingService.Trace("Entered in method UpdateErrorOnPostCreate.");
			if (entity != null &&
					string.IsNullOrWhiteSpace(errorMessage) == false)//Entity context is not null and errorCode is not empty
			{
				iTracingService.Trace("Entity context is not null");
				iTracingService.Trace("Error Message = {0}", errorMessage);
				try
				{
					//Updating the field "Message" on Lead entity with the Error Message
					Entity lead = new Entity("lead");
					lead.Id = entity.Id;
					lead.Attributes.Add("resi_message", errorMessage);
					iOrganizationService.Update(lead);
					iTracingService.Trace("Error Message updated.");
				}
				catch (Exception ex)
				{ throw ex; }
			}
		}
		#endregion

		#region Recherche_Demande
		private Entity Recherche_Demande(OptionSetValue reglesDeDispatch,
																		 EntityReference leadParentContact,
																		 IOrganizationService iOrganizationService,
																		 ITracingService iTracingService)
		{
			Entity opportunity = null;
			iTracingService.Trace("Entered in method Recherche_Demande.");

			if (reglesDeDispatch != null &&
					reglesDeDispatch.Value == 1)
			{
				#region reglesDeDispatch == 1 (Programme)
				// Fetch Related "Opportunity" entity records
				QueryExpression opportunityQueryExpression = new QueryExpression("opportunity");
				opportunityQueryExpression.ColumnSet = new ColumnSet("ownerid", "resi_zonedechalandiseid");
				opportunityQueryExpression.Criteria.AddCondition("resi_cipreferentid", ConditionOperator.NotNull);
				opportunityQueryExpression.Criteria.AddCondition("resi_searchdate", ConditionOperator.OlderThanXMonths, 6);
				opportunityQueryExpression.Criteria.AddCondition("parentcontactid", ConditionOperator.Equal, leadParentContact.Id);
				opportunityQueryExpression.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
				EntityCollection opportunityEntityCollection = iOrganizationService.RetrieveMultiple(opportunityQueryExpression);

				iTracingService.Trace("retrieved Related Opportunity = {0} records", opportunityEntityCollection.Entities.Count);
				if (opportunityEntityCollection.Entities.Count > 0)
				{
					opportunity = opportunityEntityCollection.Entities[0];
				}
				#endregion
			}
			return opportunity;
		}
		#endregion

		#endregion
	}

	/// <summary>
	/// Executes plug-in code in response to an event.
	/// </summary>
	/// <param name="serviceProvider">Type: Returns_IServiceProvider. A container for service objects. Contains references to the plug-in execution context (<see cref="T:Microsoft.Xrm.Sdk.IPluginExecutionContext"></see>), tracing service (<see cref="T:Microsoft.Xrm.Sdk.ITracingService"></see>), organization service (<see cref="T:Microsoft.Xrm.Sdk.IOrganizationServiceFactory"></see>), and notification service (<see cref="T:Microsoft.Xrm.Sdk.IServiceEndpointNotificationService"></see>).</param>
	/// <exception cref="InvalidPluginExecutionException">
	/// Invalid 
	/// or
	/// An error occurred in the PreLeadQualify plug-in. + ex.Message
	/// </exception>
	public class PreLeadQualify //:IPlugin
	{
		#region Variable Declaration

		private IOrganizationService iOrganizationService = null;
		private ITracingService iTracingService = null;

		#endregion
		/// <summary>
		/// Plugin execution point
		/// </summary>
		public void Execute(IServiceProvider iServiceProvider)
		{
			// Obtain the execution context from the service provider.
			IPluginExecutionContext iPluginExecutionContext = (IPluginExecutionContext)iServiceProvider.GetService(typeof(IPluginExecutionContext));

			// Obtain the organization service factory reference.
			IOrganizationServiceFactory iOrganizationServiceFactory = (IOrganizationServiceFactory)iServiceProvider.GetService(typeof(IOrganizationServiceFactory));

			// Obtain the organization service reference.
			iOrganizationService = iOrganizationServiceFactory.CreateOrganizationService(iPluginExecutionContext.UserId);

			// Obtain the tracing service reference.
			iTracingService = (ITracingService)iServiceProvider.GetService(typeof(ITracingService));
			iTracingService.Trace("Plugin entered");
			try
			{
				EntityReference leadId = (EntityReference)iPluginExecutionContext.InputParameters["LeadId"];
				iTracingService.Trace("leadId fetched from Input Parameters of Context");

				//Get "Existing Contact" lookup field value.
				Entity lead = iOrganizationService.Retrieve("lead", leadId.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("parentcontactid"));
				iTracingService.Trace("lead entity record retrieved from the LeadId fetched");
				EntityReference leadExistingContact = null;

				if (lead != null &&
					lead.Attributes.Contains("parentcontactid"))
				{
					iTracingService.Trace("lead is not null and contains parent contact id");
					leadExistingContact = CRMData.GetAttributeValue<EntityReference>(lead, null, "parentcontactid");
					if (leadExistingContact != null)
					{
						iTracingService.Trace("leadExistingContact is not null");
						//Set the system contact creation false by default.
						iPluginExecutionContext.InputParameters["CreateContact"] = false;
						iTracingService.Trace("Set the creation of contact on lead qualify to false");
					}
				}
			}
			catch (FaultException<OrganizationServiceFault> ex)
			{
				throw new InvalidPluginExecutionException("An error occurred in the PreLeadQualify plug-in." + ex.Message);
			}
		}
	}

	/// <summary>
	/// Executes plug-in code in response to an event.
	/// </summary>
	/// <param name="serviceProvider">Type: Returns_IServiceProvider. A container for service objects. Contains references to the plug-in execution context (<see cref="T:Microsoft.Xrm.Sdk.IPluginExecutionContext"></see>), tracing service (<see cref="T:Microsoft.Xrm.Sdk.ITracingService"></see>), organization service (<see cref="T:Microsoft.Xrm.Sdk.IOrganizationServiceFactory"></see>), and notification service (<see cref="T:Microsoft.Xrm.Sdk.IServiceEndpointNotificationService"></see>).</param>
	/// <exception cref="InvalidPluginExecutionException">
	/// Invalid 
	/// or
	/// An error occurred in the PreLeadQualify plug-in. + ex.Message
	/// </exception>
	public class PostLeadQualify //:IPlugin
	{
		#region Variable Declaration

		private IOrganizationService iOrganizationService = null;
		private ITracingService iTracingService = null;

		#endregion
		/// <summary>
		/// Plugin execution point
		/// </summary>
		public void Execute(IServiceProvider iServiceProvider)
		{
			// Obtain the execution context from the service provider.
			IPluginExecutionContext iPluginExecutionContext = (IPluginExecutionContext)iServiceProvider.GetService(typeof(IPluginExecutionContext));

			// Obtain the organization service factory reference.
			IOrganizationServiceFactory iOrganizationServiceFactory = (IOrganizationServiceFactory)iServiceProvider.GetService(typeof(IOrganizationServiceFactory));

			// Obtain the organization service reference.
			iOrganizationService = iOrganizationServiceFactory.CreateOrganizationService(iPluginExecutionContext.UserId);

			// Obtain the tracing service reference.
			iTracingService = (ITracingService)iServiceProvider.GetService(typeof(ITracingService));
			iTracingService.Trace("Plugin entered");
			try
			{
				iTracingService.Trace("lead entity validated as a Entity reference");
				EntityReference leadId = (EntityReference)iPluginExecutionContext.InputParameters["LeadId"];
				iTracingService.Trace("leadId fetched from Input Parameters of Context");

				//Get "Existing Contact" lookup field value.
				Entity lead = iOrganizationService.Retrieve("lead", leadId.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("statecode",
																																																						 "parentcontactid",
																																																						 "resi_reglesdedispatchtype",
																																																						 "resi_bddfagencycode",
																																																						 "address1_postalcode",
																																																						 "resi_programname"));

				iTracingService.Trace("lead entity record retrieved from the LeadId fetched");

				int leadStateCode = 0;
				EntityReference leadExistingContact = null;
				OptionSetValue reglesDeDispatch = null;
				string CodeDAgenceDuClient = "";
				string postalCode = "";
				string nomDuProgramme = "";
				EntityReference zoneDeChalandise = null;
				EntityReference zoneOwnerTeam = null;
				EntityReference parentAccountId = null;
				EntityReference parentContactId = null;
				EntityReference opportunityId = null;

				if (lead != null)
				{
					iTracingService.Trace("lead entity is not null");
					if (lead.Attributes.Contains("statecode"))
					{
						leadStateCode = CRMData.GetAttributeValue<int>(lead, null, "statecode");
						iTracingService.Trace("leadStateCode = {0}", leadStateCode.ToString());
					}
					iTracingService.Trace("leadStateCode = {0}", leadStateCode.ToString());
					if (lead.Attributes.Contains("parentcontactid"))
					{ leadExistingContact = CRMData.GetAttributeValue<EntityReference>(lead, null, "parentcontactid"); }

					if (lead.Attributes.Contains("resi_reglesdedispatchtype"))
					{ reglesDeDispatch = CRMData.GetAttributeValue<OptionSetValue>(lead, null, "resi_reglesdedispatchtype"); }

					if (lead.Attributes.Contains("resi_bddfagencycode"))
					{ CodeDAgenceDuClient = CRMData.GetAttributeValue<string>(lead, null, "resi_bddfagencycode"); }

					if (lead.Attributes.Contains("address1_postalcode"))
					{ postalCode = CRMData.GetAttributeValue<string>(lead, null, "address1_postalcode"); }

					if (lead.Attributes.Contains("resi_programname"))
					{ nomDuProgramme = CRMData.GetAttributeValue<string>(lead, null, "resi_programname"); }

					foreach (EntityReference created in (IEnumerable<object>)iPluginExecutionContext.OutputParameters["CreatedEntities"])
					{
						//Check if Account record is created when lead qualified.
						if (created.LogicalName == "account")
						{
							iTracingService.Trace("parent account is created on lead qualify");
							parentAccountId = new EntityReference("account", created.Id);
						}

						//Check if Contact record is created when lead qualified.
						if (created.LogicalName == "contact")
						{
							iTracingService.Trace("parent contact is created on lead qualify");
							parentContactId = new EntityReference("contact", created.Id);
						}

						//Check if Oppo record is created when lead qualified.
						if (created.LogicalName == "opportunity")
						{
							iTracingService.Trace("opportunity is created on lead qualify");
							opportunityId = new EntityReference("opportunity", created.Id);
						}
					}
				}

				if (opportunityId != null && leadExistingContact != null)
				{
					iTracingService.Trace("opportunity is created and lead entity qualify contains parent Contact");
					// Fetch Related "Opportunity" entity records
					QueryExpression opportunityQueryExpression = new QueryExpression("opportunity");
					opportunityQueryExpression.ColumnSet = new ColumnSet("ownerid");
					opportunityQueryExpression.Criteria.AddCondition("resi_origine", ConditionOperator.Equal, "BDDF");
					opportunityQueryExpression.Criteria.AddCondition("parentcontactid", ConditionOperator.Equal, leadExistingContact.Id);
					opportunityQueryExpression.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
					EntityCollection opportunityEntityCollection = iOrganizationService.RetrieveMultiple(opportunityQueryExpression);

					iTracingService.Trace("retrieved Related Opportunity = {0} records", opportunityEntityCollection.Entities.Count);

					/*If related Opportunity is found update the Opportunity record created by lead qualification and 
          assign it to the owner of the related Opportunity fetched.*/
					if (opportunityEntityCollection.Entities.Count > 0)
					{
						iTracingService.Trace("Openned opportunity exists");
						EntityReference ownerOpportunity = CRMData.GetAttributeValue<EntityReference>(opportunityEntityCollection.Entities[0], null, "ownerid");
						iTracingService.Trace("Owner of the opened opportunity retrieved");
						//Update the Opportunity record created by qualifying the lead
						Entity opportunity = new Entity("opportunity");
						opportunity.Id = opportunityId.Id;
						opportunity.Attributes.Add("parentcontactid", leadExistingContact);
						iOrganizationService.Update(opportunity);
						iTracingService.Trace("Parent Contact of the Opportunity updated populated with the parent contact of the lead qualified.");

						// Create assign request for enity
						AssignRequest assignRequest = new AssignRequest();
						assignRequest.Assignee = new EntityReference(opportunityId.LogicalName, opportunityId.Id);
						assignRequest.Target = new EntityReference(ownerOpportunity.LogicalName, ownerOpportunity.Id);
						iOrganizationService.Execute(assignRequest);
						iTracingService.Trace("assigned Opportunity to the owner of the open opportunity fetched.");
					}
					else
					{
						#region reglesDeDispatch == 3 (Code Agence)
						if (reglesDeDispatch != null &&
							reglesDeDispatch.Value == 3)
						{
							iTracingService.Trace("reglesDeDispatch is Code Agence");
							if (string.IsNullOrWhiteSpace(CodeDAgenceDuClient) == false)
							{
								//Entity leadContact = iOrganizationService.Retrieve(leadExistingContact.LogicalName, leadExistingContact.Id, new ColumnSet("parentcustomerid"));
								//if (leadContact!=null)
								//{
								//  EntityReference accountRelatedToLeadContactId = CRMData.GetAttributeValue<EntityReference>(leadContact, null, "parentcustomerid");
								//  Entity accountRelatedToLeadContact = iOrganizationService.Retrieve(accountRelatedToLeadContactId.LogicalName, accountRelatedToLeadContactId.Id, new ColumnSet("resi_zonedechalandiseid"));
								//  if (accountRelatedToLeadContact != null)
								//  {
								//    zoneDeChalandise = CRMData.GetAttributeValue<EntityReference>(accountRelatedToLeadContact, null, "resi_zonedechalandiseid");
								//    Entity zone = iOrganizationService.Retrieve("resi_salesarea", zoneDeChalandise.Id, new ColumnSet("ownerid"));
								//    if (zone != null)
								//    {
								//      zoneOwnerTeam = CRMData.GetAttributeValue<EntityReference>(zone, null, "ownerid");
								//    }
								//  }
								//}

								// Fetch "Account" entity records 
								//QueryExpression accountQueryExpression = new QueryExpression("account");
								//accountQueryExpression.ColumnSet = new ColumnSet("resi_sales_mgt_area");
								//accountQueryExpression.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, CodeDAgenceDuClient);
								//EntityCollection accountEntityCollection = iOrganizationService.RetrieveMultiple(accountQueryExpression);

								//iTracingService.Trace("retrieved Accounts = {0} records", accountEntityCollection.Entities.Count);

								//if (accountEntityCollection.Entities.Count > 0)
								//{
								//  zoneDeChalandise = CRMData.GetAttributeValue<EntityReference>(accountEntityCollection.Entities[0], null, "resi_zonedechalandiseid");
								//  Entity zone = iOrganizationService.Retrieve("resi_salesarea", zoneDeChalandise.Id, new ColumnSet("ownerid"));
								//  if (zone != null)
								//  {
								//    zoneOwnerTeam = CRMData.GetAttributeValue<EntityReference>(zone, null, "ownerid");
								//  }
								//}

								/**
                                                # Attribution du lead au responsable client de l'equipe du client
                                               **/
								AssignRequest assignRequest = new AssignRequest()
								{
									Assignee = new EntityReference
									{
										LogicalName = "team",
										Id = new Guid("8DB86DF8-06F4-E611-8113-3863BB350E28")
									},

									Target = new EntityReference(opportunityId.LogicalName, opportunityId.Id)
								};

								iOrganizationService.Execute(assignRequest);

								// Create assign request for enity
								//AssignRequest assignRequest = new AssignRequest();
								//assignRequest.Assignee = new EntityReference("team", new Guid("8DB86DF8-06F4-E611-8113-3863BB350E28"));
								//assignRequest.Target = new EntityReference(opportunityId.LogicalName, opportunityId.Id);
								//iOrganizationService.Execute(assignRequest);
							}
						}
						#endregion

						#region reglesDeDispatch == 2 (Code Postale)
						if (reglesDeDispatch != null &&
							reglesDeDispatch.Value == 2)
						{
							iTracingService.Trace("reglesDeDispatch is Code Postale");
							if (!string.IsNullOrWhiteSpace(postalCode))
							{
								// Fetch "Code Postal" entity records 
								QueryExpression codePostalQueryExpression = new QueryExpression("resi_codepostal");
								codePostalQueryExpression.ColumnSet = new ColumnSet("resi_zonedechalandiseid");
								codePostalQueryExpression.Criteria.AddCondition("resi_name", ConditionOperator.Equal, postalCode);
								EntityCollection codePostalEntityCollection = iOrganizationService.RetrieveMultiple(codePostalQueryExpression);

								iTracingService.Trace("retrieved Postal Code = {0} records", codePostalEntityCollection.Entities.Count);

								if (codePostalEntityCollection.Entities.Count > 0)
								{
									zoneDeChalandise = CRMData.GetAttributeValue<EntityReference>(codePostalEntityCollection.Entities[0], null, "resi_zonedechalandiseid");
									Entity zone = iOrganizationService.Retrieve("resi_salesarea", zoneDeChalandise.Id, new ColumnSet("ownerid"));
									if (zone != null)
									{
										zoneOwnerTeam = CRMData.GetAttributeValue<EntityReference>(zone, null, "ownerid");
									}
								}
							}
							else
							{
								AssignRequest assignRequest = new AssignRequest()
								{
									Assignee = new EntityReference
									{
										LogicalName = "team",
										Id = new Guid("D0FD6D90-07F4-E611-8113-3863BB350E28")
									},

									Target = new EntityReference(opportunityId.LogicalName, opportunityId.Id)
								};

								iOrganizationService.Execute(assignRequest);

								// Create assign request for enity
								//AssignRequest assignRequest = new AssignRequest();
								//assignRequest.Assignee = new EntityReference(opportunityId.LogicalName, opportunityId.Id);
								//assignRequest.Target = new EntityReference("team", new Guid("D0FD6D90-07F4-E611-8113-3863BB350E28"));
								//iOrganizationService.Execute(assignRequest);
							}
						}
						#endregion

						#region reglesDeDispatch == 1 (Programme)
						if (reglesDeDispatch != null &&
							reglesDeDispatch.Value == 1)
						{
							iTracingService.Trace("reglesDeDispatch is Code Programme");
							if (!string.IsNullOrWhiteSpace(nomDuProgramme))
							{
								// Fetch "Programme" entity records 
								QueryExpression programmeQueryExpression = new QueryExpression("Programme");
								programmeQueryExpression.ColumnSet = new ColumnSet("resi_zonedechalandiseid");
								programmeQueryExpression.Criteria.AddCondition("resi_name", ConditionOperator.Equal, nomDuProgramme);
								EntityCollection programmeEntityCollection = iOrganizationService.RetrieveMultiple(programmeQueryExpression);

								iTracingService.Trace("retrieved Programme = {0} records", programmeEntityCollection.Entities.Count);

								if (programmeEntityCollection.Entities.Count > 0)
								{
									zoneDeChalandise = CRMData.GetAttributeValue<EntityReference>(programmeEntityCollection.Entities[0], null, "resi_zonedechalandiseid");
									Entity zone = iOrganizationService.Retrieve("resi_salesarea", zoneDeChalandise.Id, new ColumnSet("ownerid"));
									if (zone != null)
									{
										zoneOwnerTeam = CRMData.GetAttributeValue<EntityReference>(zone, null, "ownerid");
									}

								}
							}
							else
							{
								AssignRequest assignRequest = new AssignRequest()
								{
									Assignee = new EntityReference
									{
										LogicalName = "team",
										Id = new Guid("01940A8E-0AF4-E611-8113-3863BB350E28")
									},

									Target = new EntityReference(opportunityId.LogicalName, opportunityId.Id)
								};

								iOrganizationService.Execute(assignRequest);

								// Create assign request for enity
								//AssignRequest assignRequest = new AssignRequest();
								//assignRequest.Assignee = new EntityReference(opportunityId.LogicalName, opportunityId.Id);
								//assignRequest.Target = new EntityReference("team", new Guid("01940A8E-0AF4-E611-8113-3863BB350E28"));
								//iOrganizationService.Execute(assignRequest);
							}
						}
						#endregion


						if (opportunityId != null &&
							zoneDeChalandise != null)
						{
							Entity opportunity = new Entity("opportunity");
							opportunity.Id = opportunityId.Id;
							opportunity.Attributes.Add("resi_zonedechalandiseid", zoneDeChalandise);
							iOrganizationService.Update(opportunity);
							iTracingService.Trace("Zone updated");
						}

						if (zoneOwnerTeam != null)
						{
							AssignRequest assignRequest = new AssignRequest()
							{
								Assignee = new EntityReference
								{
									LogicalName = zoneOwnerTeam.LogicalName,
									Id = zoneOwnerTeam.Id
								},

								Target = new EntityReference(opportunityId.LogicalName, opportunityId.Id)
							};

							iOrganizationService.Execute(assignRequest);
							// Create assign request for enity
							//AssignRequest assignRequest = new AssignRequest();
							//assignRequest.Assignee = new EntityReference(zoneOwnerTeam.LogicalName, zoneOwnerTeam.Id);
							//assignRequest.Target = new EntityReference(opportunityId.LogicalName, opportunityId.Id);
							//iOrganizationService.Execute(assignRequest);

							iTracingService.Trace("Entity record is successfully assigned to its owner of the Zone fetched.");
						}
					}
				}
			}
			catch (FaultException<OrganizationServiceFault> ex)
			{
				throw new InvalidPluginExecutionException("An error occurred in the PostLeadQualify plug-in." + ex.Message);
			}
		}
	}
}


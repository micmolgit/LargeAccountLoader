using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.IO;
using System.Collections.Generic;

using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;

using Microsoft.Crm.Sdk.ServiceHelper;

namespace LargeAccountLoader
{
    class LargeAccountLoader
    {
        private readonly DcrmConnector _dcrmConnector = new DcrmConnector();

        public LargeAccountLoader()
        {
            _dcrmConnector = new DcrmConnector();
        }

        public void Terminate()
        {
            if (_dcrmConnector != null)
                _dcrmConnector.Disconnect();

            Console.WriteLine("Press <Enter> to exit.");
            Console.ReadLine();
        }

        public async Task<int> Load(string accountGUID)
        {
            _dcrmConnector.Connect();
            var cptEntitiesCount = await GetPartyByGuid(accountGUID);
            return cptEntitiesCount;
        }

        public async Task<int> GetEntitiesCountAsync(QueryExpression queryParty)
        {
            int cptContact = 0;

            await Task<int>.Run(() =>
            {
                Stopwatch sw = new Stopwatch();
                try
                {
                    sw.Start();
                    using (StreamWriter output = new StreamWriter("RequestOutput.xml"))
                    {
                        SoapLoggerOrganizationService slos = new SoapLoggerOrganizationService(_dcrmConnector.ServiceConfig.OrganizationUri, _dcrmConnector.ServiceProxy, output);

                        //EntityCollection parties = _dcrmConnector.ServiceProxy.RetrieveMultiple(queryParty);
                        EntityCollection parties = slos.RetrieveMultiple(queryParty);
                        if (parties.Entities != null)
                        {
                            cptContact = parties.Entities.Count;
                        }
                        else
                        {
                            Console.WriteLine("GetEntitiesCountAsync() : PARTY_COLLECTION NULL");
                        }
                    }
                }
                catch (Exception Ex)
                {
                    var Innermessage = Ex.InnerException != null ? Ex.InnerException.Message : "";
                    Console.WriteLine($"[GetContactCountFromParty] {Ex.Message}  {Innermessage}");
                }
                finally
                {
                    sw.Stop();
                    Console.WriteLine($"GetEntitiesCountAsync() execution time  : {new DateTime(sw.ElapsedTicks).ToString("HH: mm:ss.fff")}");
                }
            });

            return cptContact;
        }


        public async Task<int> GetPartyByGuid(string guidParty)
        {
            var GetRolesClient = true;
            var GetMoyensDeContact = true;
            var GetConnexions = true;

            // Connect to the Organization service.
            // The using statement assures that the service proxy will be properly disposed.

            //<snippetQueryExpressionPagingWithCookie1>
            // Query using the paging cookie.
            // Define the paging attributes.
            // The number of records per page to retrieve.
            int queryCount = 1000;

            // Initialize the page number.
            int pageNumber = 1;

            // Initialize the number of records.
            int recordCount = 0;

            // Define the condition expression for retrieving records.
             //ConditionExpression pagecondition = new ConditionExpression();
             //pagecondition.AttributeName = "accountid";
             //pagecondition.Operator = ConditionOperator.Equal;
             //pagecondition.Values.Add(guidParty);

            // Define the order expression to retrieve the records.
            // OrderExpression order = new OrderExpression();
            // order.AttributeName = "name";
            // order.OrderType = OrderType.Ascending;

            // Create the query expression and add condition.
            QueryExpression queryParty = new QueryExpression();
            queryParty.EntityName = "account";
            queryParty.Criteria.AddCondition("accountid", ConditionOperator.Equal, guidParty);
            //queryParty.Criteria.AddCondition(pagecondition);
            //queryParty.Orders.Add(order);


            // Assign the pageinfo properties to the query expression.
            // queryParty.PageInfo = new PagingInfo();
            // queryParty.PageInfo.Count = queryCount;
            // queryParty.PageInfo.PageNumber = pageNumber;

            queryParty.ColumnSet = new ColumnSet("parentaccountid", "accountnumber", "accountid", "name", "crm_segment_id", "crm_ind_prenom", "crm_ind_nom",
                                      "crm_ind_date_naissance", "crm_ind_departement_naissance", "crm_nom_marital_usage", "crm_soc_raison_sociale",
                                      "crm_soc_siren", "crm_soc_siret", "crm_administration", "crm_type_code", "crm_datedernierecompletude", "crm_titre_code");

            // The current paging cookie. When retrieving the first page,
            // pagingCookie should be null.
            //queryParty.PageInfo.PagingCookie = null;
            Console.WriteLine("Retrieving accounts according to choosen criterias...\n");

            #region roleclient
            var linkRoleClient = new LinkEntity("account", "crm_roleclient", "accountid", "crm_account_id", JoinOperator.LeftOuter);
            linkRoleClient.EntityAlias = "roleClient";
            if (GetRolesClient == true)
            {
                linkRoleClient.Columns = new ColumnSet("crm_role", "crm_roleclientid", "crm_account_id");

                var linkClientFacturation = new LinkEntity("roleClient", "crm_clientdefacturation", "crm_clientdefacturation_id", "crm_clientdefacturationid", JoinOperator.LeftOuter);
                linkClientFacturation.EntityAlias = "clientFacturation";
                linkClientFacturation.Columns = new ColumnSet("crm_customer_id", "crm_custcode");

                var linkContrat = new LinkEntity("clientFacturation", "crm_contrat", "crm_clientdefacturationid", "crm_clientdefacturation_id", JoinOperator.LeftOuter);
                linkContrat.EntityAlias = "contrat";
                linkContrat.Columns = new ColumnSet("crm_contratid", "crm_msisdn", "crm_coid");

                var linkContratUtilisateur = new LinkEntity("roleClient", "crm_contrat", "crm_roleclientid", "crm_utilisateur_id", JoinOperator.LeftOuter);
                linkContratUtilisateur.EntityAlias = "roleClientUtil";
                linkContratUtilisateur.Columns = new ColumnSet("crm_contratid", "crm_msisdn", "crm_coid");

                linkClientFacturation.LinkEntities.Add(linkContrat);
                linkRoleClient.LinkEntities.Add(linkClientFacturation);
                linkRoleClient.LinkEntities.Add(linkContratUtilisateur);
            }
            queryParty.LinkEntities.Add(linkRoleClient);
            #endregion

            #region moyens de contact
            if (GetMoyensDeContact == true)
            {
                var linkContactUtilisateur = new LinkEntity("account", "contact", "accountid", "parentcustomerid", JoinOperator.LeftOuter);
                linkContactUtilisateur.EntityAlias = "contact";
                linkContactUtilisateur.Columns = new ColumnSet("crm_npai_oui_non", "crm_opt_in_out", "contactid", "crm_canal", "address1_line2",
                              "address1_line3", "address1_line1", "address1_postofficebox", "address1_city", "address1_postalcode", "address1_country",
                              "emailaddress1", "fax", "crm_login", "telephone1", "mobilephone", "contactid", "crm_type_voie_nom_voie", "crm_numero_voie");
                queryParty.LinkEntities.Add(linkContactUtilisateur);
            }
            #endregion

            #region Connexion
            if (GetConnexions == true)
            {
                var linkConnexionUtilisateurDe = new LinkEntity("partyUtilisateur", "connection", "accountid", "record1id", JoinOperator.LeftOuter);
                linkConnexionUtilisateurDe.EntityAlias = "connexion";
                linkConnexionUtilisateurDe.Columns = new ColumnSet("record2id", "record2roleid", "connectionid");
                queryParty.LinkEntities.Add(linkConnexionUtilisateurDe);
            }
            #endregion

            #region Connaissance client
            var linkConnaissanceClient =
              new LinkEntity("partyUtilisateur", "crm_connaissanceclient", "accountid", "crm_party_id", JoinOperator.LeftOuter);
            linkConnaissanceClient.EntityAlias = "ConnaissanceClient";
            linkConnaissanceClient.Columns = new ColumnSet("crm_party_id", "crm_nomgerantsociete");

            #endregion

            #region Mise Sous Responsabilité
            var linkMiseResponsabilite =
              new LinkEntity("ConnaissanceClient", "crm_misesousresponsabilite", "crm_connaissanceclientid", "crm_connaissanceclient_id", JoinOperator.LeftOuter);
            linkMiseResponsabilite.EntityAlias = "MiseResponsabilite";
            linkMiseResponsabilite.Columns = new ColumnSet("crm_connaissanceclient_id", "crm_misesousresponsabiliteid", "crm_date", "crm_motif", "crm_type", "crm_nom");
            linkConnaissanceClient.LinkEntities.Add(linkMiseResponsabilite);
            #endregion

            queryParty.LinkEntities.Add(linkConnaissanceClient);

            Console.WriteLine($"Querying DCRM for LargeAccount : {guidParty} ...");

            var entityCount = await GetEntitiesCountAsync(queryParty);

            // while (true)
            // {
            //     // Retrieve the page.
            //     EntityCollection results =  _dcrmConnector.ServiceProxy.RetrieveMultiple(queryParty);
            //     if (results.Entities != null)
            //     {
            //         // Retrieve all records from the result set.
            //         foreach (Account acct in results.Entities)
            //         {
            //             Console.WriteLine($"{++recordCount}.\t{acct.Name}\t{acct.Id}");
            //         }
            //         return results.Entities.Count;
            //     }

            //     // Check for more records, if it returns true.
            //     if (results.MoreRecords)
            //     {
            //         Console.WriteLine("\n****************\nPage number {0}\n****************", queryParty.PageInfo.PageNumber);

            //         // Increment the page number to retrieve the next page.
            //         queryParty.PageInfo.PageNumber++;

            //         // Set the paging cookie to the paging cookie returned from current results.
            //         queryParty.PageInfo.PagingCookie = results.PagingCookie;
            //     }
            //     else
            //     {
            //         // If no more records are in the result nodes, exit the loop.
            //         break;
            //     }
            // }
            return entityCount;
        }
    }
}
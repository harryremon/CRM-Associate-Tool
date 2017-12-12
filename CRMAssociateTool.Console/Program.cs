using System;
using System.Activities.Statements;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using CRMAssociateTool.Console.Helpers;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace CRMAssociateTool.Console
{
    class Program
    {
        private static IOrganizationService _orgService;

        [STAThread]
        static void Main(string[] args)
        {
            Initialize();
        }

        private static void Initialize()
        {
            System.Console.Clear();
            System.Console.WriteLine("========================");
            System.Console.WriteLine("== CRM Associate Tool ==");
            System.Console.WriteLine("========================");
            System.Console.WriteLine();
            System.Console.WriteLine("Please Choose an Action:");
            System.Console.WriteLine();
            System.Console.WriteLine("1) Connect to CRM");
            System.Console.WriteLine("2) View Saved Connections");
            System.Console.WriteLine("3) Add a new Connection");
            var firstChoise = System.Console.ReadKey(true);
            System.Console.WriteLine();
            switch (firstChoise.KeyChar)
            {
                case '1':
                    ConnectToCrm();
                    break;
                case '2':
                    ViewConnections();
                    break;
                case '3':
                    AddConnection();
                    break;
                default:
                    ConsoleHelpers.WriteError("Invalid Selection");
                    System.Console.WriteLine("Press any key to start over ...");
                    System.Console.ReadKey(true);
                    Initialize();
                    break;
            }
        }

        private static void ConnectToCrm()
        {
            System.Console.Clear();
            var index = 1;
            System.Console.WriteLine("Connect to CRM");
            System.Console.WriteLine();
            System.Console.WriteLine("Choose the connection:");
            System.Console.WriteLine();
            ConfigurationManager.RefreshSection("connectionStrings");
            if (ConfigurationManager.ConnectionStrings.Count < 1)
                AddConnection();
            foreach (ConnectionStringSettings connectionString in ConfigurationManager.ConnectionStrings)
            {
                System.Console.WriteLine($"{index++}) {connectionString.Name}");
            }

            if (int.TryParse(System.Console.ReadKey(true).KeyChar.ToString(), out var connectionIndex) && connectionIndex <= index && connectionIndex > 0)
            {
                var connectionString = ConfigurationManager.ConnectionStrings[--connectionIndex];
                var chosenString = connectionString.ConnectionString;
                System.Console.WriteLine();
                System.Console.Write($"Connecting to {connectionString.Name} ");
                ConsoleHelpers.PlayLoadingAnimation();
                try
                {
                    var crmServiceClient = new CrmServiceClient(chosenString);
                    ConsoleHelpers.StopLoadingAnimation();
                    if (!crmServiceClient.IsReady)
                    {
                        ConsoleHelpers.WriteError("Couldn't Connect");
                        System.Console.WriteLine();
                        System.Console.WriteLine("Press any key to start over ...");
                        System.Console.ReadKey(true);
                        ConnectToCrm();
                    }
                    ConsoleHelpers.WriteSuccess("Connected Successfully");
                    System.Console.WriteLine();
                    _orgService = (IOrganizationService)crmServiceClient.OrganizationWebProxyClient ??
                                  crmServiceClient.OrganizationServiceProxy;
                    System.Console.Title = $"CRM Bulk Association Tool - Connected: {connectionString.Name}";

                    ViewActions();
                }
                catch (Exception exception)
                {
                    ConsoleHelpers.StopLoadingAnimation();
                    ConsoleHelpers.WriteError(exception.Message);
                    System.Console.WriteLine();
                    System.Console.WriteLine("Press any key to start over ...");
                    System.Console.ReadKey(true);
                    ConnectToCrm();
                }

            }
            else
            {
                System.Console.WriteLine("Invalid Input");
                ConnectToCrm();
            }
        }

        private static void ViewActions()
        {
            System.Console.WriteLine("Choose Action:");
            System.Console.WriteLine();
            System.Console.WriteLine("1) Export N:N Relationships");
            System.Console.WriteLine("2) Import N:N Relationships");
            System.Console.WriteLine("3) Disconnect from CRM");
            switch (System.Console.ReadKey(true).KeyChar)
            {
                case '1':
                    StartExport();
                    break;
                case '2':
                    Startimport();
                    break;
                case '3':
                    _orgService = null;
                    ConnectToCrm();
                    break;
            }
        }

        private static void Startimport()
        {
            System.Console.Clear();
            System.Console.WriteLine("Starting Import:");
            System.Console.WriteLine();
            System.Console.WriteLine("Enter CSV file location:");
            var pathLine = System.Console.ReadLine();
            System.Console.WriteLine();
            if (!pathLine.EndsWith("csv") && !pathLine.EndsWith("csv\""))
            {
                ConsoleHelpers.WriteError("only csv files are supported.");
                System.Console.WriteLine();
                System.Console.WriteLine("Press any key to go back ...");
                System.Console.ReadKey(true);
                System.Console.Clear();
                ViewActions();
            }
            var csvLines = File.ReadAllLines(pathLine);
            System.Console.WriteLine($"Found {csvLines.Length} records in file");
            System.Console.WriteLine();
            var index = 1;
            var good = 0;
            var bad = 0;
            var badOnes = new List<int>();
            foreach (var line in csvLines)
            {
                System.Console.WriteLine($"Processing Record {index} of {csvLines.Length}");
                var cells = line.Split(',');
                if (cells.Length != 5)
                {
                    bad++;
                    badOnes.Add(index);
                    ConsoleHelpers.WriteError("Invalid Data in File, Skipping Record");
                    System.Console.WriteLine();
                    continue;
                }
                try
                {
                    var mainName = cells[0];
                    var mainId = Guid.Parse(cells[1]);
                    var relatedName = cells[2];
                    var relatedId = Guid.Parse(cells[3]);
                    var relationship = cells[4];
                    System.Console.WriteLine("Associating Record");
                    _orgService.Associate(mainName, mainId, new Relationship(relationship),
                        new EntityReferenceCollection { new EntityReference(relatedName, relatedId) });
                    good++;
                    ConsoleHelpers.WriteSuccess("Associated Successfully");
                    System.Console.WriteLine();
                }
                catch (Exception exception)
                {
                    bad++;
                    badOnes.Add(index);
                    ConsoleHelpers.WriteError($"{exception.Message}, Skipping Record");
                    System.Console.WriteLine();
                }
                index++;
            }
            System.Console.WriteLine($"Finished - {good} Done, {bad} Skipped");
            System.Console.Write("Skipped Records: ");
            badOnes.ForEach(i => System.Console.Write($"{i}, "));
            System.Console.WriteLine();
            System.Console.WriteLine("Press any key to go back ...");
            System.Console.ReadKey(true);
            System.Console.Clear();
            Initialize();
        }

        private static void StartExport()
        {
            System.Console.Clear();
            System.Console.WriteLine("Starting Export:");
            System.Console.WriteLine();
            System.Console.Write("Please type the main entity logical name: ");
            var mainEntity = System.Console.ReadLine();
            System.Console.WriteLine();
            var filteredIds = CheckForFiltration(mainEntity);
            System.Console.Write($"Retrieveing {mainEntity} relationships");
            ConsoleHelpers.PlayLoadingAnimation();
            var customRelationships = CrmHelpers.GetCustomRelationships(_orgService, mainEntity,
                new[] { CrmHelpers.RelationType.ManyToManyRelationships });
            ConsoleHelpers.StopLoadingAnimation();
            ConsoleHelpers.WriteSuccess("Retrieved Relationships");
            System.Console.WriteLine();
            Export(customRelationships, filteredIds, mainEntity);
        }

        private static List<Guid> CheckForFiltration(string mainEntity)
        {
            System.Console.WriteLine("Add Filtration ? (Y = Yes, Any Key = No)");
            var getFiltration = System.Console.ReadKey(true);
            if (getFiltration.Key != ConsoleKey.Y) return null;
            System.Console.WriteLine();
            System.Console.WriteLine("Paste the filter part of the FetchXML:");
            string readLine;
            var conditionXml = "";
            do
            {
                readLine = System.Console.ReadLine();
                conditionXml += $"{readLine}\r\n";
            } while (!string.IsNullOrWhiteSpace(readLine?.Trim().Replace("\r\n", "")));

            if (Helpers.Helpers.ValidateXml(conditionXml))
            {
                System.Console.WriteLine();
                System.Console.Write("Retrieving filtered Ids");
                ConsoleHelpers.PlayLoadingAnimation();
                var fetchxml = $"<fetch>\r\n  <entity name=\"{mainEntity}\" >\r\n {conditionXml} </entity>\r\n</fetch>";
                var fetch = new FetchExpression(fetchxml);
                var filteredIds = _orgService.RetrieveMultiple(fetch);
                ConsoleHelpers.StopLoadingAnimation();
                return filteredIds.Entities.Select(entity => entity.Id).ToList();
            }
            ConsoleHelpers.WriteError("Pasted Text is not valid xml");
            System.Console.WriteLine();
            CheckForFiltration(mainEntity);
            return null;
        }

        public static void Export(List<RelationshipMetadataBase> customRelationships, List<Guid> filteredIds, string typedMainEntity)
        {
            System.Console.WriteLine("Choose Relationship:");
            var index = 1;
            customRelationships.ForEach(relation => System.Console.WriteLine($"{index++}) {relation.SchemaName}"));
            if (int.TryParse(System.Console.ReadKey(true).KeyChar.ToString(), out var connectionIndex) &&
                connectionIndex <= index && connectionIndex > 0)
            {
                System.Console.WriteLine();
                var relationshipMetadata = (ManyToManyRelationshipMetadata)customRelationships[--connectionIndex];
                var relationshipSchemaName = relationshipMetadata.SchemaName;
                var relationshipName = relationshipMetadata.IntersectEntityName;
                var mainEntity = relationshipMetadata.Entity1LogicalName;
                var relatedEntity = relationshipMetadata.Entity2LogicalName;
                if (!string.Equals(mainEntity, typedMainEntity))
                    StringHelpers.Swap(ref mainEntity, ref relatedEntity);
                var query = new QueryExpression(relationshipName) { ColumnSet = new ColumnSet(true) };
                if (filteredIds != null && filteredIds.Any())
                    query.Criteria = new FilterExpression
                    {
                        Conditions = { new ConditionExpression($"{mainEntity}id", ConditionOperator.In, filteredIds) }
                    };
                System.Console.Write("Retrieving Relationship Records");
                ConsoleHelpers.PlayLoadingAnimation();
                var relationshipRecords = _orgService.RetrieveMultiple(query);
                ConsoleHelpers.StopLoadingAnimation();
                System.Console.WriteLine();
                ConsoleHelpers.WriteSuccess($"Retrieved {relationshipRecords.Entities.Count} Records");
                System.Console.WriteLine();
                System.Console.Write("Saving to file");
                ConsoleHelpers.PlayLoadingAnimation();
                var relationshipsDetailed = relationshipRecords.Entities
                    .Select(
                        relationship =>
                            $"{mainEntity},{relationship.Attributes[$"{mainEntity}id"]},{relatedEntity},{relationship.Attributes[$"{relatedEntity}id"]},{relationshipSchemaName}")
                    .ToList();
                var path = AppDomain.CurrentDomain.BaseDirectory;
                var csvPath = $"{path}{DateTime.Now.ToFileTime()}.csv";
                File.WriteAllLines(csvPath, relationshipsDetailed);
                ConsoleHelpers.StopLoadingAnimation();
                System.Console.WriteLine();
                System.Windows.Forms.Clipboard.SetText(csvPath);
                ConsoleHelpers.WriteSuccess("File saved, Path copied to clipboard");
                System.Console.WriteLine();
                System.Console.WriteLine("Press any key to go back...");
                System.Console.ReadKey(true);
                System.Console.Clear();
                Initialize();
            }
            else
            {
                System.Console.WriteLine();
                ConsoleHelpers.WriteError("Invalid Option");
                System.Console.WriteLine();
                Export(customRelationships, filteredIds, typedMainEntity);
            }
        }

        private static void ViewConnections()
        {
            System.Console.Clear();
            System.Console.WriteLine("Saved Connections:");
            System.Console.WriteLine();
            foreach (ConnectionStringSettings connectionString in ConfigurationManager.ConnectionStrings)
            {
                System.Console.WriteLine(connectionString.Name);
                System.Console.WriteLine($"==> {connectionString}");
                System.Console.WriteLine();
            }
            System.Console.WriteLine();
            ConsoleHelpers.WriteSuccess("End of Connection Strings, Press any key to go back ...");
            System.Console.ReadKey(true);
            Initialize();
        }

        private static void AddConnection()
        {
            System.Console.Clear();
            System.Console.WriteLine("Add a New Connection:");
            System.Console.WriteLine();
            System.Console.Write("Enter Organization Url (http://<url>/<organization>/): ");
            var organizationUrl = System.Console.ReadLine();
            System.Console.WriteLine();
            System.Console.Write("Enter AD Domain: ");
            var domain = System.Console.ReadLine();
            System.Console.WriteLine();
            System.Console.Write("Enter Username: ");
            var username = System.Console.ReadLine();
            System.Console.WriteLine();
            System.Console.Write("Enter Password: ");
            var password = ConsoleHelpers.ReadPassword();
            System.Console.WriteLine();
            var connectionString =
                $"AuthType=IFD;Url={organizationUrl}; Domain={domain}; Username={username}; Password={password}";
            System.Console.WriteLine();
            System.Console.Write("Testing Connection ");
            ConsoleHelpers.PlayLoadingAnimation();
            try
            {
                var crmServiceClient = new CrmServiceClient(connectionString);
                ConsoleHelpers.StopLoadingAnimation();
                if (!crmServiceClient.IsReady)
                {
                    ConsoleHelpers.WriteError("Couldn't Connect");
                    System.Console.WriteLine();
                    System.Console.WriteLine("Press any key to start over ...");
                    System.Console.ReadKey(true);
                    AddConnection();
                }
                ConsoleHelpers.WriteSuccess("Connected Successfully");
                System.Console.WriteLine();
                System.Console.Write("Enter a Name for this connection:  ");
                var connectionName = System.Console.ReadLine();
                var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                config.ConnectionStrings.ConnectionStrings.Add(new ConnectionStringSettings(connectionName, connectionString));
                config.Save(ConfigurationSaveMode.Modified);
                System.Console.WriteLine();
                ConsoleHelpers.WriteSuccess("Connection Saved, Press any key to go back ...");
                System.Console.ReadKey(true);
                Initialize();
            }
            catch (Exception exception)
            {
                ConsoleHelpers.StopLoadingAnimation();
                ConsoleHelpers.WriteError(exception.Message);
                System.Console.WriteLine();
                System.Console.WriteLine("Press any key to start over ...");
                System.Console.ReadKey(true);
                AddConnection();
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using CRMAssociateTool.Console.Helpers;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;

namespace CRMAssociateTool.Console
{
    class Program
    {
        private static IOrganizationService _orgService;

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
                    ConsoleHelpers.WriteSuccess("Connected Succefully");
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
            switch (System.Console.ReadKey(true).KeyChar)
            {
                case '1':
                    StartExport();
                    break;
                case '2':
                    Startimport();
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
                    ConsoleHelpers.WriteSuccess("Associated Succefully");
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
            ViewActions();
        }

        private static void StartExport()
        {
            System.Console.Clear();
            ConsoleHelpers.WriteError("Not Currently Implemented");
            System.Console.WriteLine();
            System.Console.Write("Press any key to go back ...");
            System.Console.ReadKey(true);
            System.Console.Clear();
            ViewActions();
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
                $"AuthType=AD;Url={organizationUrl}; Domain={domain}; Username={username}; Password={password}";
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
                ConsoleHelpers.WriteSuccess("Connected Succefully");
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

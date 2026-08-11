using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelDataImporter;
using ExcelDataImporter.Handler;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics.X86;
using System.Text.Json.Nodes;

class Program
{

    static void Main()
    {
        try
        {
            Console.CancelKeyPress += OnCancelKeyPress;
            Dictionary<string, string> _errorMap;
            try
            {
                var path = Path.Combine(AppContext.BaseDirectory, "error-mapping.json");
                if (File.Exists(path))
                {
                    var json = File.ReadAllText(path);
                    _errorMap = JsonConvert.DeserializeObject<Dictionary<string, string>>(json)
                                ?? new Dictionary<string, string>();
                }
                else
                {
                    Console.WriteLine($"Error mapping file not found: {path}");
                    Console.ReadLine();
                    return;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception while loading error-mapping.json - " + e.Message);
                Console.ReadLine();
                return;
            }

            var existingDbCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            var config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            var tenantId = config["KeyVault:TenantId"];
            var clientId = config["KeyVault:ClientId"];
            var clientSecret = config["KeyVault:ClientSecret"];
            var vaultUrl = config["KeyVault:VaultUrl"];
            var aesKeyName = config["KeyVault:aeskeyname"];

            var missingKeys = new List<string>();

            if (string.IsNullOrWhiteSpace(tenantId)) missingKeys.Add("KeyVault:TenantId");
            if (string.IsNullOrWhiteSpace(clientId)) missingKeys.Add("KeyVault:ClientId");
            if (string.IsNullOrWhiteSpace(clientSecret)) missingKeys.Add("KeyVault:ClientSecret");
            if (string.IsNullOrWhiteSpace(vaultUrl)) missingKeys.Add("KeyVault:VaultUrl");
            if (string.IsNullOrWhiteSpace(aesKeyName)) missingKeys.Add("KeyVault:aeskeyname");

            if (missingKeys.Count > 0)
            {
                Console.WriteLine($"Missing or empty configuration keys: {string.Join(", ", missingKeys)}");
                Console.ReadLine();
            }



            if (!Uri.TryCreate(vaultUrl, UriKind.Absolute, out var vaultUri))
            {
                Console.WriteLine($"KeyVault:VaultUrl '{vaultUrl}' is not a valid URI.");
                Console.ReadLine();
            }

            var keyVaultCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            var keyVaultClient = new SecretClient(vaultUri, keyVaultCredential);
            var encryptionService = new AesEncryptionService(keyVaultClient, aesKeyName);




            var availableUseCases = new List<string>

{   "Prerequisites",
    "Company","Supervisory Org","Cost Center","Pay Group","Department","Location",
    "Compensation Element","Compensation Grade","Compensation Grade Profile","Worker Compensation Code","Compensation Plan",
    "Job Family Group","Job Family","Job Profile","Payment Election Rules","Position","State",
    "Worker", "Worker Bank Account Payment Election","Worker Dependents",
    "Worker Beneficiary",
    "Worker Emergency Contacts","Org Manager Update","Worker Compensation Grade",
    "Worker Company Assignment",
     "Worker Location Assignment",
     "Worker Supervisory Org Assignment",
     "Worker Pay Group Assignment",
     "Worker Cost Center Assignment",
     "Worker Department Assignment",
    //"Update Worker Date Of Birth",
    "Company Event","Cost Center Event","Department Event",
    "Position Event","Supervisory Org Event","Location Event","Job Profile Event","Worker Event","Compensation Plan Event","Period Schedule Export Event","Long Term Leave Plan Export Event","Time Calculation Tag Export Event","Worker Payment Election Event","Worker Bank Accounts Event","Worker Long Term Leave Plan Event",
    "Valid Values Export",
    "Valid Values Import",
    "Localized Valid Values Export",
    "Localized Valid Values Import",
    "HCM Valid Values Export",
    "HCM Localized Valid Values Export",
    "State Export",
    "Currency Export",
    "Balance Period Export",
    "Period Schedule Export",
    "Time Calculation Tag Export",
    "Worker Lists Export",
    "Work Schedule Calendar Export",
    "Worker Eligibility Rules Export",
    "Holiday Calendar Export",
    "Time Off Plan Export",
    "Long Term Leave Plan Export",
    "Time Entry Template Export",
     "Custom Report Export",
      "User Feature Permission Constraint Export",
    //"Role Export",
    //"User Role Mapping Export",
    //"Workflow Authorized Role Export",
   
   
    "Benefit Plan",
    "Currency Conversion",
    "Worker Time Off",
    "Worker Long Term Leave Plan",
    "Worker Work Schedule",
    "Worker Holiday Calendar",
    "Worker Time Entry Template",
    "Worker Active Long Term Leave Transaction",
    //delete usecases
    "Delete System Process",
"Delete Symmetry Integration",
"Delete Notification Message Queue",
"Delete Document Info",
"Delete Biz Events",
"Delete Worker Transaction",
"Delete Worker Compensation Grade",
"Delete Worker Org Mapping",
"Delete Worker Benefit Plans",
"Delete Worker Emergency Contacts",
"Delete Worker Dependents",
"Delete Worker Beneficiary",
"Delete Hcm Export",
"Delete Worker",
"Delete Position",
"Delete Prerequisites",
"Delete Audit Log",
"Delete Table Data",
"Delete Worker Lists",
"Before Delete Table Data",
"After Delete Table Data",




     };

            var prerequisiteEntities = new List<string>
        {
            "Company","Supervisory Org","Cost Center","Pay Group","Department","Location",
            "Compensation Element","Compensation Grade","Compensation Grade Profile","Worker Compensation Code","Compensation Plan",
            "Job Family Group","Job Family","Job Profile","Payment Election Rules","Position"
        };


            var workerOrgMappingEntities = new List<string>
        {
                "Worker Company Assignment","Worker Location Assignment","Worker Supervisory Org Assignment","Worker Pay Group Assignment","Worker Cost Center Assignment",
                "Worker Department Assignment"
        };

            while (true)
            {
            start:
                Console.Clear();
                Console.WriteLine("===== Excel Data Importer =====\n");
                string connStr = config["ConnectionString"]!;
                var builder = new NpgsqlConnectionStringBuilder(connStr);
                string host = builder.Host;
                string InstanceName = config["Instance"];
                Console.WriteLine($"Version:47.3,Released On:11/08/2026");

                try
                {
                    using (var conn = new NpgsqlConnection(connStr))
                    {
                        conn.Open();

                        string databaseName;
                        string schemaName;

                        using (var cmd = new NpgsqlCommand("SELECT current_database();", conn))
                        {
                            databaseName = cmd.ExecuteScalar()?.ToString() ?? "Unknown";
                        }

                        using (var cmd = new NpgsqlCommand("SELECT current_schema();", conn))
                        {
                            schemaName = cmd.ExecuteScalar()?.ToString() ?? "Unknown";
                        }

                        Console.WriteLine("\n======= Database Details =======");
                        Console.WriteLine($"Database Name : {databaseName}");
                        Console.WriteLine($"Schema Name   : {schemaName}");
                        //Console.WriteLine("================================");
                    }
                }
                catch(Exception ex)
                {
                    Console.WriteLine($"Error:{ ex.Message}");
                }

                Console.WriteLine("\n");

                Console.WriteLine("=======Instance Details========");
                Console.WriteLine($"IP Address:{host} ,Instance Name:{InstanceName}");
                Console.WriteLine("===============================");
                Console.WriteLine("\nAvailable Use Cases:\n");

                Console.WriteLine("==========PREREQUISITES ENTITIES==========");
                Console.WriteLine(string.Join(", ", availableUseCases.Take(1)));
                Console.WriteLine(string.Join(", ", availableUseCases.Skip(1).Take(17)));
                //Console.WriteLine(string.Join(", ", availableUseCases.Skip(77).Take(1)));

                Console.WriteLine();
                Console.WriteLine("==========WORKER ENTITIES==========");
                Console.WriteLine(string.Join(", ", availableUseCases.Skip(18).Take(7)));
                

                Console.WriteLine();
                Console.WriteLine("==========Worker Org Mapping==========");
                Console.WriteLine(string.Join(", ", availableUseCases.Skip(25).Take(6)));



                Console.WriteLine();
                Console.WriteLine("==========EVENTS==========");
                Console.WriteLine(string.Join(", ", availableUseCases.Skip(31).Take(15)));
                Console.WriteLine();
                Console.WriteLine("==========VALID VALUES ENTITIES==========");
                Console.WriteLine(string.Join(", ", availableUseCases.Skip(46).Take(6)));
                Console.WriteLine();
                Console.WriteLine("==========EXPORT ENTITIES==========");
                Console.WriteLine(string.Join(", ", availableUseCases.Skip(52).Take(14)));
               
                Console.WriteLine();
                Console.WriteLine("==========WORKER OTHER ENTITIES==========");
                Console.WriteLine(string.Join(", ", availableUseCases.Skip(66).Take(8)));
                Console.WriteLine();
                Console.WriteLine("==========DELETE ENTITIES==========");
                Console.WriteLine(string.Join(", ", availableUseCases.Skip(74).Take(21)));
                //Console.WriteLine(string.Join(", ", availableUseCases.Skip(80).Take(6)));

                Console.WriteLine();

                string? useCaseInput;
                string? singlePrerequisite = null;
                bool isPrerequisites = false;
                bool isworkerorgmappingentites = false;
                string? singleworkerorgmappingentites = null;


                while (true)
                {

                    Console.Write("\nEnter UseCase:");
                    useCaseInput = Console.ReadLine()?.Trim();

                    if (string.IsNullOrEmpty(useCaseInput))
                    {
                        Console.WriteLine("Input cannot be empty. Please enter a valid use case.");
                        continue;
                    }

                    if (useCaseInput.Equals("exit", StringComparison.OrdinalIgnoreCase))
                        return;

                    if (useCaseInput.Equals("prerequisites", StringComparison.OrdinalIgnoreCase))
                    {
                        isPrerequisites = true;
                        break;
                    }

                    if (prerequisiteEntities.Any(e => e.Equals(useCaseInput, StringComparison.OrdinalIgnoreCase)))
                    {
                        singlePrerequisite = prerequisiteEntities
                            .First(e => e.Equals(useCaseInput, StringComparison.OrdinalIgnoreCase));
                        isPrerequisites = true;
                        break;
                    }


                    if (useCaseInput.Equals("worker org mapping", StringComparison.OrdinalIgnoreCase))
                    {
                        isworkerorgmappingentites = true;
                        break;
                    }

                    if (workerOrgMappingEntities.Any(e => e.Equals(useCaseInput, StringComparison.OrdinalIgnoreCase)))
                    {
                        singleworkerorgmappingentites = workerOrgMappingEntities
                            .First(e => e.Equals(useCaseInput, StringComparison.OrdinalIgnoreCase));
                        isworkerorgmappingentites = true;
                        break;
                    }
                    if (availableUseCases.Any(u => u.Equals(useCaseInput, StringComparison.OrdinalIgnoreCase)))
                    {
                        break;
                    }

                    Console.WriteLine($"Invalid use case: '{useCaseInput}'.");
                }

                if (useCaseInput.ToLower() == "valid values export")
                {
                    try
                    {
                        ValidValuesExportHandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}"); }

                }
                else if (useCaseInput.ToLower() == "localized valid values export")
                {
                    try
                    {
                        LocalizedValidValuesExportHandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}"); }
                }
                else if (useCaseInput.ToLower() == "delete worker")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {

                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        int deleteWorkerCount = 100;

                        if (!int.TryParse(config["DeleteWorkerRecordsCount"], out deleteWorkerCount))
                        {
                            deleteWorkerCount = 100;
                        }

                        var deleteWrkrHndlr = new DeleteWorkerHandler(config);
                        deleteWrkrHndlr.Run(deleteWorkerCount);



                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}"); }

                }
                else if (useCaseInput.ToLower() == "delete prerequisites")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deletePrerequisiteHandler = new DeletePrerequisiteHandler(config);
                        deletePrerequisiteHandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }
                else if (useCaseInput.ToLower() == "delete position")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deletePositionHandler = new DeletePositionHandler(config);
                        deletePositionHandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }
                else if (useCaseInput.ToLower() == "delete worker beneficiary")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deleteworkerHandler = new DeleteWorkerBeneficiaryhandler(config);
                        deleteworkerHandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }
                else if (useCaseInput.ToLower() == "delete worker dependents")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deleteworkerdependent = new deleteworkerdepnednt(config);
                        deleteworkerdependent.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }
                else if (useCaseInput.ToLower() == "delete worker emergency contacts")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deletePositionHandler = new deletememrgency(config);
                        deletePositionHandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }

                else if (useCaseInput.ToLower() == "delete worker benefit plans")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deleteWorkerBenefitplans = new DeleteWorkerBenefitplans(config);
                        deleteWorkerBenefitplans.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }
                else if (useCaseInput.ToLower() == "reset worker compensation amount")
                {

                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a 'update' operation.");
                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with update operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var resetWrkrCompHandler = new ResetWorkerCompAmtHandler(config, encryptionService);
                        resetWrkrCompHandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }

                }


                else if (useCaseInput.ToLower() == "delete benefit plans")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deleteBenefitPlansHandler = new DeleteBenefitPlansHandler(config);
                        deleteBenefitPlansHandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }


                else if (useCaseInput.ToLower() == "delete worker org mapping")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deleteWorkerOrgMappingHandler = new DeleteWorkerOrgMappingHandler(config);
                        deleteWorkerOrgMappingHandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }


                else if (useCaseInput.ToLower() == "delete worker compensation grade")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deleteWorkerCompensationGradeHandler = new DeleteWorkerCompensationGradeHandler(config);
                        deleteWorkerCompensationGradeHandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }


                else if (useCaseInput.ToLower() == "delete hcm export")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var hCMExportHandler = new DeleteHCMExportHandler(config);
                        hCMExportHandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }

                else if (useCaseInput.ToLower() == "delete worker transaction")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deleteWorkerTransctionHandler = new DeleteWorkerTransctionHandler(config);
                        deleteWorkerTransctionHandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }

                else if (useCaseInput.ToLower() == "delete biz events")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deletebizeventshandler = new deletebizeventshandler(config);
                        deletebizeventshandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }

                else if (useCaseInput.ToLower() == "delete document info")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deleteDocumentInfohandler = new DeleteDocumentInfohandler(config);
                        deleteDocumentInfohandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }


                else if (useCaseInput.ToLower() == "delete notification message queue")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deleteNotoficationhandler = new DeleteNotoficationhandler(config);
                        deleteNotoficationhandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }


                else if (useCaseInput.ToLower() == "delete symmetry integration")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deleteSymmetryintegrationhandler = new DeleteSymmetryintegrationhandler(config);
                        deleteSymmetryintegrationhandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }



                else if (useCaseInput.ToLower() == "delete system process")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deleteSystemProcesshandler = new DeleteSystemProcesshandler(config);
                        deleteSystemProcesshandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }



                else if (useCaseInput.ToLower() == "delete audit log")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deleteAuditLogHandler = new DeleteAuditLogHandler(config);
                        deleteAuditLogHandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }
                else if (useCaseInput.ToLower() == "delete table data")
                {
                    try
                    {
                        while (true)
                        {

                            var allowedTables = config
                           .GetSection("DeleteTableData:AllowedTables")
                           .GetChildren()
                           .Select(x => x.Value!)
                           .ToArray();

                            if (allowedTables.Length == 0)
                            {
                                throw new InvalidOperationException(
                                    "Missing or empty configuration for DeleteTableData:AllowedTables.");
                            }

                            Console.WriteLine($"\nDeleting data for the following tables: {string.Join(", ", allowedTables)}\n");

                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");


                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deleteTableDataHandler = new DeleteTableDataHandler(config);
                        deleteTableDataHandler.Run(useCaseInput);

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }

                else if (useCaseInput.ToLower() == "delete worker lists")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deleteWorkerListHandler = new DeleteWorkerListHandler(config);
                        deleteWorkerListHandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }

                else if (useCaseInput.ToLower() == "after delete table data")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deleteWorkerListHandler = new AfterDeleteTableDataHandler(config);
                        deleteWorkerListHandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }

                else if (useCaseInput.ToLower() == "before delete table data")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a delink operation.");
                            //Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delink operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deleteWorkerListHandler = new BeforeDeleteTableDataHandler(config);
                        deleteWorkerListHandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }

                else if (useCaseInput.ToLower() == "delete backup logs")
                {
                    try
                    {
                        while (true)
                        {
                            Console.WriteLine("\n WARNING: You are about to perform a DELETE operation.");
                            Console.WriteLine("This action may permanently remove data from the database.");

                            Console.Write("\nType 'YES' to confirm or 'NO' to cancel: ");
                            var input = Console.ReadLine()?.Trim();

                            if (string.Equals(input, "YES", StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("\nProceeding with delete operation...");
                                break;
                            }
                            else if (string.Equals(input, "NO", StringComparison.OrdinalIgnoreCase))
                            {
                                goto start;
                            }
                            else
                            {
                                Console.WriteLine("\n Invalid input. Please type 'YES' or 'NO'.");
                            }
                        }
                        var deleteWorkerListHandler = new DeletebackupLogsHandler(config);
                        deleteWorkerListHandler.Run();

                    }
                    catch (Exception e) { Console.WriteLine($"\nUnexpected error: {e.Message}{e.StackTrace}"); }
                }

                else
                {
                    string operation;
                    bool validateOnly = false;
                    while (true)
                    {
                        Console.Write("Enter operation (validation/save): ");
                        operation = Console.ReadLine()?.Trim().ToLower();

                        if (operation == "validation" || operation == "save")
                        {
                            validateOnly = operation == "validation";
                            break;
                        }
                        Console.WriteLine("Invalid operation. Try again.");
                    }



                    try
                    {
                        using var conn = new NpgsqlConnection(connStr);
                        conn.Open();
                        string configusername = config["username"];

                        if (isPrerequisites)
                        {
                            ProcessPrerequisites(config, _errorMap, conn, existingDbCounts, validateOnly, singlePrerequisite);
                        }
                        else if (isworkerorgmappingentites)
                        {
                            ProcessWorkerOrgMappings(config, _errorMap, conn, configusername, existingDbCounts, validateOnly, singleworkerorgmappingentites);
                        }
                        else
                        {
                            switch (useCaseInput.ToLowerInvariant())
                            {
                                case "worker":
                                    ProcessWorker(config, _errorMap, conn, existingDbCounts, validateOnly, encryptionService);
                                    break;
                                case "worker compensation grade":
                                    var workerCompensationGradeHandler = new WorkerCompensationGradeHandler(config, existingDbCounts);
                                    workerCompensationGradeHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "legacy manager worker update":
                                    var legacyManagerWorkerUpdateHandler = new LegacyManagerWorkerUpdateHandler(config);
                                    legacyManagerWorkerUpdateHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "currency export":
                                    var currency = new CurrencyHandler(config, existingDbCounts);
                                    currency.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "state":
                                    var stateHandler1 = new StateHandler(config, existingDbCounts);
                                    stateHandler1.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                //case "worker org mapping":
                                //    var workerOrgMappingHandler = new WorkerOrgMappingHandler(config, existingDbCounts);
                                //    workerOrgMappingHandler.ProcessSheets(validateOnly, configusername, useCaseInput);
                                //    break;
                                case "org manager update":
                                    var managerUpdateHandler = new ManagerUpdateHandler(config, conn, _errorMap, existingDbCounts);
                                    managerUpdateHandler.ProcessSheets(validateOnly);
                                    break;
                                case "company event":
                                    var companyeventHandeler = new CompanyEventHandler(config, existingDbCounts, encryptionService);
                                    companyeventHandeler.ProcessSheets(useCaseInput, validateOnly);
                                    break;
                                case "cost center event":
                                    var costCenterEventHandler = new CostCentEventHandler(config, existingDbCounts, encryptionService);
                                    costCenterEventHandler.ProcessSheets(useCaseInput, validateOnly);
                                    break;
                                case "department event":
                                    var departmentEventHandler = new DepartmentEventHandler(config, existingDbCounts, encryptionService);
                                    departmentEventHandler.ProcessSheets(useCaseInput, validateOnly);
                                    break;
                                case "location event":
                                    var locationEventHandler = new LocationEventHandler(config, existingDbCounts, encryptionService);
                                    locationEventHandler.ProcessSheets(useCaseInput, validateOnly);
                                    break;
                                case "supervisory org event":
                                    var supervisoryOrgEventHandler = new SupervisoryOrgEventHandler(config, existingDbCounts, encryptionService);
                                    supervisoryOrgEventHandler.ProcessSheets(useCaseInput, validateOnly);
                                    break;
                                case "position event":
                                    var positioneventHander = new PositionEventHandler(config, existingDbCounts, encryptionService);
                                    positioneventHander.ProcessSheets(useCaseInput, validateOnly);
                                    break;
                                case "job profile event":
                                    var jobProfileEventHandler = new JobProfileEventHandler(config, existingDbCounts, encryptionService);
                                    jobProfileEventHandler.ProcessSheets(useCaseInput, validateOnly);
                                    break;
                                case "period schedule export event":
                                    var periodScheduleEventHandler = new PeriodScheduleEventHandler(config, existingDbCounts, encryptionService);
                                    periodScheduleEventHandler.ProcessSheets(useCaseInput, validateOnly);
                                    break;
                                case "compensation plan event":
                                    var compensationPlanEvent = new CompensationPlanEvent(config, existingDbCounts, encryptionService);
                                    compensationPlanEvent.ProcessSheets(useCaseInput, validateOnly);
                                    break;
                                case "time calculation tag export event":
                                    var timeCalculationTagEventExportHandler = new TimeCalculationTagEventExportHandler(config, existingDbCounts, encryptionService);
                                    timeCalculationTagEventExportHandler.ProcessSheets(useCaseInput, validateOnly);
                                    break;
                                case "long term leave plan export event":
                                    var longTermLeavePlanEventHandler = new LongTermLeavePlanEventHandler(config, existingDbCounts, encryptionService);
                                    longTermLeavePlanEventHandler.ProcessSheets(useCaseInput, validateOnly);
                                    break;
                                case "worker payment election event":
                                    var workerPaymentElectionEventHandler = new WorkerPaymentElectionEventHandler(config, existingDbCounts, encryptionService);
                                    workerPaymentElectionEventHandler.ProcessSheets(useCaseInput, validateOnly);
                                    break;
                                case "worker bank accounts event":
                                    var bankDetailsEventHandler = new WorkerBankDetailsEventHandler(config, existingDbCounts, encryptionService);
                                    bankDetailsEventHandler.ProcessSheets(useCaseInput, validateOnly);
                                    break;
                                case "worker benefit plans event":
                                    var workerBenefitPlanEventsHandler = new WorkerBenefitPlanEventsHandler(config, existingDbCounts, encryptionService);
                                    workerBenefitPlanEventsHandler.ProcessSheets(useCaseInput, validateOnly);
                                    break;
                                case "worker long term leave plan event":
                                    var workerLongTermLeavePlanEventHandler = new WorkerLongTermLeavePlanEventHandler(config, existingDbCounts, encryptionService);
                                    workerLongTermLeavePlanEventHandler.ProcessSheets(useCaseInput, validateOnly);
                                    break;
                                case "worker event":
                                    var workereventHandler = new WorkerEventHandler(config, existingDbCounts, encryptionService);
                                    workereventHandler.ProcessSheets(useCaseInput, validateOnly);
                                    break;
                                case "benefit provider":
                                    var benefitprovider = new BenefitProviderHandler(config, existingDbCounts);
                                    benefitprovider.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "benefit plan":
                                    var benefitplanhandler = new BenefitPlanHandler(config, existingDbCounts);
                                    benefitplanhandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "currency conversion":
                                    var currencyConversionHandler = new CurrencyConversionHandler(config, existingDbCounts);
                                    currencyConversionHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "worker beneficiary":
                                    var workerbeneficiaryHandler = new WorkerBeneficiaryHandler(config, existingDbCounts, encryptionService, _errorMap);
                                    workerbeneficiaryHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "worker compensation plan":
                                    var workerCompensationPlanHandler = new WorkerCompensationPlanHandler(config, existingDbCounts, encryptionService);
                                    workerCompensationPlanHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "worker benefit plans":
                                    var workerbenefitplansHandler = new WorkerBenefitPlansHandler(config, existingDbCounts, encryptionService, _errorMap);
                                    workerbenefitplansHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "worker dependents":
                                    var workerdependentshandler = new WorkerDependentsHandler(config, existingDbCounts, encryptionService, _errorMap);
                                    workerdependentshandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;

                                case "worker bank account payment election":
                                    var workerBankAccountPaymentElectionHandler = new WorkerBankAccountPaymentElectionHandler(config, existingDbCounts, encryptionService);
                                    workerBankAccountPaymentElectionHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "worker emergency contacts":
                                    var workeremergencycontactshandler = new WorkerEmergencyContactsHandler(config, existingDbCounts, encryptionService, _errorMap);
                                    workeremergencycontactshandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "holiday calendar export":
                                    var holidayCalendarHandler = new HolidayCalenderExportHandler(config, existingDbCounts);
                                    holidayCalendarHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "valid values import":
                                    var validValuesImportHandler = new ValidValuesImportHandler(config);
                                    validValuesImportHandler.ProcessSheets(validateOnly, configusername);
                                    break;
                                case "hcm valid values export":
                                    var hCMValidValuesExportHandler = new HCMValidValuesExportHandler(config);
                                    hCMValidValuesExportHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "hcm localized valid values export":
                                    var hCMLocalizedvalidValuesExport = new HCMLocalizedvalidValuesExport(config);
                                    hCMLocalizedvalidValuesExport.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "time off plan export":
                                    var timeoffplanhandler = new TimeOffPlanHandler(config, existingDbCounts);
                                    timeoffplanhandler.ProcessSheets(useCaseInput, validateOnly, configusername, _errorMap);
                                    break;
                                case "worker time off":
                                    var workertimeoffplanhandler = new WorkerTimeOffHandler(config, existingDbCounts);
                                    workertimeoffplanhandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "balance period export":
                                    var balanceperiodplanhandler = new BalancePeriodHandler(config, existingDbCounts);
                                    balanceperiodplanhandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "worker lists export":
                                    var workerListHandler = new WorkerListHandler(config, existingDbCounts);
                                    workerListHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "custom report export":
                                    var customReportExportHandler = new CustomReportExportHandler(config, existingDbCounts);
                                    customReportExportHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "work schedule calendar export":
                                    var workschedulecalendarhandler = new WorkScheduleCalendarHandler(config, existingDbCounts);
                                    workschedulecalendarhandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "worker long term leave plan":
                                    var workerLongTermLeavePlanHandler = new WorkerLongTermLeavePlanHandler(config, existingDbCounts);
                                    workerLongTermLeavePlanHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "worker work schedule":
                                    var workerWorkSchedule = new WorkerWorkScheduleHandler(config, existingDbCounts);
                                    workerWorkSchedule.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "worker holiday calendar":
                                    var workerHolidayCalendar = new WorkerHolidayCalendarHandler(config, existingDbCounts);
                                    workerHolidayCalendar.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                //case "period export":
                                //    var periodhandler = new PeriodHandler(config, existingDbCounts);
                                //    periodhandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                //    break;
                                case "period schedule export":
                                    var periodschedulehandler = new PeriodScheduleHandler(config, existingDbCounts);
                                    periodschedulehandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "worker time entry template":
                                    var workerTimeEntryTemplatesHandler = new WorkerTimeEntryTemplatesHandler(config, existingDbCounts);
                                    workerTimeEntryTemplatesHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "update worker date of birth":
                                    var UpdateWorkerDateOfBirthHandler = new UpdateWorkerDateOfBirthHandler(config, existingDbCounts, encryptionService);
                                    UpdateWorkerDateOfBirthHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "worker time off request history":
                                    var workerTimeOffRequestHistoryHandler = new WorkerTimeOffRequestHistoryHandler(config, existingDbCounts);
                                    workerTimeOffRequestHistoryHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "worker time off transaction":
                                    var worker_Time_Off_TransactionHandler = new WorkerTimeOffTransactionHandler(config, existingDbCounts);
                                    worker_Time_Off_TransactionHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "worker eligibility rules export":
                                    var workerEligiblityRulesHandler = new WorkerEligiblityRulesHandler(config, existingDbCounts);
                                    workerEligiblityRulesHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "time calculation tag export":
                                    var timeCalculationTagHandler = new TimeCalculationTagHandler(config, existingDbCounts);
                                    timeCalculationTagHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "role export":
                                    var roleExportHandler = new RoleExportHandler(config, existingDbCounts);
                                    roleExportHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "survey response":
                                    var surveyResponseHandler = new SurveyResponseHandler(config, existingDbCounts);
                                    surveyResponseHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "time entry template export":
                                    var timeEntryTemplateHandler = new TimeEntryTemplateHandler(config, existingDbCounts);
                                    timeEntryTemplateHandler.ProcessSheets(useCaseInput, validateOnly, configusername, _errorMap);
                                    break;
                                case "user role mapping export":
                                    var userRoleMappingHandler = new UserRoleMappingHandler(config, existingDbCounts);
                                    userRoleMappingHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "workflow authorized role export":
                                    var workflowAuthorizedRoleHandler = new WorkflowAuthorizedRoleHandler(config, existingDbCounts);
                                    workflowAuthorizedRoleHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "long term leave plan export":
                                    var longTermLeavePlanHandler = new LongTermLeavePlanHandler(config, existingDbCounts);
                                    longTermLeavePlanHandler.ProcessSheets(useCaseInput, validateOnly, configusername, _errorMap);
                                    break;
                                case "user feature permission constraint export":
                                    var featurePermissionUserConstraintHandler = new UserFeaturePermissionConstraintHandler(config, existingDbCounts);
                                    featurePermissionUserConstraintHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "localized valid values import":
                                    var updateLocalizedValidValuesHandler = new UpdateLocalizedValidValuesImportHandler(config);
                                    updateLocalizedValidValuesHandler.ProcessSheets(validateOnly);
                                    break;
                                case "state export":
                                    var stateHandler = new StateExportHandler(config, existingDbCounts);
                                    stateHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "usecase dml":
                                    var usecasedmlhandler = new usecasehandler(config, existingDbCounts);
                                    usecasedmlhandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;
                                case "worker active long term leave transaction":
                                    var workeractiveleavetransctionHandler = new WorkeractiveleavetransctionHandler(config, existingDbCounts);
                                    workeractiveleavetransctionHandler.ProcessSheets(useCaseInput, validateOnly, configusername);
                                    break;

                                default:
                                    Console.WriteLine($"Unknown use case: {useCaseInput}");
                                    break;
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"\nUnexpected error: {ex.Message}");
                    }

                }

                Console.WriteLine($"\nProcessing completed..{useCaseInput}");
                Console.Write("Press ENTER to continue or type 'exit' to quit: ");
                string? next = Console.ReadLine()?.Trim();
                if (next?.Equals("exit", StringComparison.OrdinalIgnoreCase) == true)
                    break;

            }


            static void ProcessPrerequisites(
        IConfiguration config,
        Dictionary<string, string> errorMap,
        NpgsqlConnection conn,
        Dictionary<string, int> existingDbCounts,
        bool validateOnly,
        string? singleEntity = null)
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string configFolder = Path.Combine(baseDir, "config");

                var prerequisiteEntities = new List<string>
    {
        "Company","Supervisory Org","Cost Center","Pay Group","Department","Location",
        "Compensation Element","Compensation Grade","Compensation Grade Profile","Worker Compensation Code","Compensation Plan",
        "Job Family Group","Job Family","Job Profile","Payment Election Rules","Position"
    };

                if (!string.IsNullOrEmpty(singleEntity))
                {
                    prerequisiteEntities = prerequisiteEntities
                        .Where(e => e.Equals(singleEntity, StringComparison.OrdinalIgnoreCase))
                        .ToList();
                }

                string inputFolder = Path.Combine(baseDir, "input\\Prerequisites");
                string fileName = "prerequisites.xlsx";

                foreach (var useCase in prerequisiteEntities)
                {
                    try
                    {
                        string configFilePath = Path.Combine(configFolder, $"{useCase}_config.txt");

                        if (!Directory.Exists(inputFolder))
                            throw new DirectoryNotFoundException($"Input folder not found: {inputFolder}");

                        if (!File.Exists(configFilePath))
                            throw new FileNotFoundException($"Config file not found: {configFilePath}");

                        string jsonText = File.ReadAllText(configFilePath);
                        JObject jsonObj = JObject.Parse(jsonText);

                        string username = config["username"]
                            ?? throw new Exception("Username not found in appsettings file");

                        string excelPath = Path.Combine(inputFolder, fileName);
                        if (!File.Exists(excelPath))
                            throw new FileNotFoundException($"Excel file not found in folder '{inputFolder}': {fileName}");

                        string outputFolder = Path.Combine(inputFolder, useCase);
                        if (!Directory.Exists(outputFolder))
                            Directory.CreateDirectory(outputFolder);

                        string logPath = Path.Combine(outputFolder, $"{useCase}_log.txt");
                        var logger = new Logger(logPath, errorMap);

                        logger.Log($"Usecase: {useCase} | Filename: {fileName} | Operation: {(validateOnly ? "Validation" : "Save")}", level: "header");

                        using var workbook = new XLWorkbook(excelPath);

                        bool isValidFormat = logger.ValidateFileFormat(workbook, jsonObj);

                        if (!isValidFormat)
                        {
                            string errorMsg = $"Excel format invalid for use case: {useCase}";
                            logger.Log(errorMsg);

                            if (!validateOnly)
                            {
                                throw new Exception(errorMsg);
                            }
                            continue;
                        }

                        switch (useCase.ToLower())
                        {
                            case "company": new CompanyHandler(workbook, logger, conn, existingDbCounts).ProcessSheets(jsonObj, validateOnly, username); break;
                            case "supervisory org": new SupervisoryOrgHandler(workbook, logger, conn, existingDbCounts).ProcessSheets(jsonObj, validateOnly, username); break;
                            case "cost center": new CostCenterHandler(workbook, logger, conn, existingDbCounts).ProcessSheets(jsonObj, validateOnly, username); break;
                            case "pay group": new PayGroupHandler(workbook, logger, conn, existingDbCounts).ProcessSheets(jsonObj, validateOnly, username); break;
                            case "department": new DepartmentHandler(workbook, logger, conn, existingDbCounts).ProcessSheets(jsonObj, validateOnly, username); break;
                            case "location": new LocationHandler(workbook, logger, conn, existingDbCounts).ProcessSheets(jsonObj, validateOnly, username); break;
                            case "compensation element": new CompensationElementHandler(workbook, logger, conn, existingDbCounts).ProcessSheets(jsonObj, validateOnly, username); break;
                            case "compensation grade": new CompensationGradeHandler(workbook, logger, conn, existingDbCounts).ProcessSheets(jsonObj, validateOnly, username); break;
                            case "compensation grade profile": new CompensationGradeProfileHandler(workbook, logger, conn, existingDbCounts).ProcessSheets(jsonObj, validateOnly, username); break;
                            case "worker compensation code": new WorkerCompensationCodeHandler(workbook, logger, conn, existingDbCounts).ProcessSheets(jsonObj, validateOnly, username); break;
                            case "compensation plan": new CompensationPlanHandler(workbook, logger, conn, existingDbCounts).ProcessSheets(jsonObj, validateOnly, username); break;
                            case "job family group": new JobFamilyGroupHandler(workbook, logger, conn, existingDbCounts).ProcessSheets(jsonObj, validateOnly, username); break;
                            case "job family": new JobFamilyHandler(workbook, logger, conn, existingDbCounts).ProcessSheets(jsonObj, validateOnly, username); break;
                            case "job profile": new JobProfileHandler(workbook, logger, conn, existingDbCounts).ProcessSheets(jsonObj, validateOnly, username); break;
                            case "payment election rules": new PaymentElectionRulesHandler(workbook, logger, conn, existingDbCounts).ProcessSheets(jsonObj, validateOnly, username); break;
                            case "position": new PositionHandler(workbook, logger, conn, existingDbCounts).ProcessSheets(jsonObj, validateOnly, username); break;
                            default:
                                logger.Log($"No handler found for use case: {useCase}");
                                break;
                        }

                        logger.CreateResultExcel(excelPath, outputFolder, jsonObj, useCase, existingDbCounts);
                        logger.Log($"Completed process for {useCase}", level: "section");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error in {useCase}: {ex.Message}");

                        if (!validateOnly)
                        {
                            throw;
                        }
                    }
                }
            }

            static void ProcessWorker(
       IConfiguration config,
       Dictionary<string, string> errorMap,
       NpgsqlConnection conn,
       Dictionary<string, int> existingDbCounts,
       bool validateOnly,
       AesEncryptionService aesEncryptionService)
            {
                using var tx = conn.BeginTransaction();

                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string configFolder = Path.Combine(baseDir, "config");

                string useCase = "Worker";

                string inputFolder = Path.Combine(baseDir, "input", "Worker");
                string fileName = "Worker.xlsx";
                string excelPath = Path.Combine(inputFolder, fileName);

                string configFilePath = Path.Combine(configFolder, $"{useCase}_config.txt");

                if (!File.Exists(configFilePath))
                {
                    Console.WriteLine($"Config file not found for {useCase}: {configFilePath}");
                    return;
                }


                int? startRow = null;
                int? endRow = null;


                while (true)
                {
                    Console.Write($"Enter row range to {(validateOnly ? "validate" : "save")} (start range from 3) or press Enter for full: ");
                    var input = Console.ReadLine();

                    if (string.IsNullOrWhiteSpace(input))
                    {
                        break;
                    }

                    if (input.Trim() == "0")
                    {
                        Console.WriteLine("Invalid input: 0 is not allowed for range.");
                        continue;
                    }

                    var parts = input.Split('-');

                    if (parts.Length == 2 &&
                        int.TryParse(parts[0], out int s) &&
                         int.TryParse(parts[1], out int e))
                    {
                        if (s <= 0 || e <= 0)
                        {
                            Console.WriteLine("Invalid range: startRow and endRow must be greater than 0.");
                            continue;
                        }

                        if (s < 3)
                        {
                            Console.WriteLine("Invalid range: startRow must be 3 or greater (rows 1 & 2 are headers).");
                            continue;
                        }

                        if (e < s)
                        {
                            Console.WriteLine("Invalid range: endRow cannot be less than startRow.");
                            continue;
                        }

                        startRow = s;
                        endRow = e;
                        break;
                    }

                    else
                    {
                        Console.WriteLine("Invalid input format. Expected format: start-end (e.g., 3-10)");
                    }
                }





                string jsonText = File.ReadAllText(configFilePath);
                JObject jsonObj = JObject.Parse(jsonText);

                string username = config["username"] ?? throw new Exception("Username not found in appsettings file");
                int batchSize = 100;

                if (!int.TryParse(config["BatchSizeForWorkerCommit"], out batchSize))
                {
                    batchSize = 100;
                }


                int validationbatchSize = 100;

                if (!int.TryParse(config["ValidationBatchSize"], out validationbatchSize))
                {
                    validationbatchSize = 100;
                }


                string outputFolder = Path.Combine(inputFolder, useCase);
                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

                string logPath = Path.Combine(outputFolder, $"{useCase}_log.txt");
                var logger = new Logger(logPath, errorMap);

                logger.Log($"Usecase: {useCase} | File: {fileName} | Mode: {(validateOnly ? "Validation" : "Save")}", level: "header");

                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Excel file not found: {excelPath}");
                    return;
                }



                using var workbook = new XLWorkbook(excelPath);

                if (!workbook.TryGetWorksheet("Worker", out var worksheet))
                {
                    Console.WriteLine("Worker sheet not found.");
                    return;
                }

                int excelLastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;

                if (excelLastRow < 3)
                {
                    Console.WriteLine("No data rows found in Worker sheet.");
                    return;
                }

                int finalStart = startRow ?? 3;
                int finalEnd = endRow ?? excelLastRow;

                finalStart = Math.Max(finalStart, 3);
                finalEnd = Math.Min(finalEnd, excelLastRow);

                if (finalStart > finalEnd)
                {
                    Console.WriteLine($"Invalid range {finalStart}-{finalEnd}. Nothing to process.");
                    return;
                }

                //bool isValidFormat = logger.ValidateFileFormat(workbook, jsonObj);

                //if (!isValidFormat)
                //{
                //    logger.Log("Excel format invalid.");
                //    return;
                //}

                bool isValidMandatoryData = logger.validateFileFormatForWorker(workbook, jsonObj, finalStart, finalEnd);

                if (!isValidMandatoryData)
                {
                    logger.Log("Excel format invalid.");
                    return;
                }


                var workerHandler = new WorkerHandler(workbook, logger, conn, tx, existingDbCounts, aesEncryptionService);
                workerHandler.ProcessSheets(jsonObj, validateOnly, username, batchSize, validationbatchSize, finalStart, finalEnd);
                logger.CreateResultExcelWorker(excelPath, outputFolder, jsonObj, useCase, existingDbCounts, startRow, endRow);
            }


            static void ProcessWorkerOrgMappings(
                IConfiguration config,
                Dictionary<string, string> errorMap,
                NpgsqlConnection conn,
                string configusername,
                Dictionary<string, int> existingDbCounts,
                bool validateOnly,
                string? singleEntity = null)
            {
                var workerOrgMappingEntities = new List<string>
        {
                "Worker Company Assignment","Worker Location Assignment","Worker Supervisory Org Assignment","Worker Pay Group Assignment","Worker Cost Center Assignment",
                "Worker Department Assignment"
        };

                if (!string.IsNullOrEmpty(singleEntity))
                {
                    workerOrgMappingEntities = workerOrgMappingEntities
                        .Where(e => e.Equals(singleEntity, StringComparison.OrdinalIgnoreCase))
                        .ToList();
                }

                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string configFolder = Path.Combine(baseDir, "config");

                string inputFolder = Path.Combine(baseDir, "input", "Worker Org Mapping");
                string fileName = "Worker Org Mapping.xlsx";

                foreach (var useCase in workerOrgMappingEntities)
                {
                    try
                    {
                        string configFilePath = Path.Combine(configFolder, $"{useCase}_config.txt");

                        if (!Directory.Exists(inputFolder))
                            throw new DirectoryNotFoundException($"Input folder not found: {inputFolder}");

                        if (!File.Exists(configFilePath))
                            throw new FileNotFoundException($"Config file not found: {configFilePath}");

                        string jsonText = File.ReadAllText(configFilePath);
                        JObject jsonObj = JObject.Parse(jsonText);

                        string username = config["username"]
                            ?? throw new Exception("Username not found in appsettings file");

                        string excelPath = Path.Combine(inputFolder, fileName);

                        if (!File.Exists(excelPath))
                            throw new FileNotFoundException(
                                $"Excel file not found in folder '{inputFolder}': {fileName}");

                        string outputFolder = Path.Combine(inputFolder, useCase);

                        if (!Directory.Exists(outputFolder))
                            Directory.CreateDirectory(outputFolder);

                        string logPath = Path.Combine(outputFolder, $"{useCase}_log.txt");

                        var logger = new Logger(logPath, errorMap);

                        logger.Log(
                            $"Usecase: {useCase} | Filename: {fileName} | Operation: {(validateOnly ? "Validation" : "Save")}",
                            level: "header");

                        using var workbook = new XLWorkbook(excelPath);

                        bool isValidFormat = logger.ValidateFileFormat(workbook, jsonObj);

                        if (!isValidFormat)
                        {
                            string errorMsg = $"Excel format invalid for use case: {useCase}";

                            logger.Log(errorMsg);

                            if (!validateOnly)
                            {
                                throw new Exception(errorMsg);
                            }

                            continue;
                        }

                        switch (useCase.ToLowerInvariant())
                        {
                            case "worker company assignment":
                                var workerCompanyAssignmentHandler = new WorkerCompanyAssignmentHandler(config, existingDbCounts);
                                workerCompanyAssignmentHandler.ProcessSheets(useCase, validateOnly, configusername);
                                break;

                            case "worker location assignment":
                                var workerLocationAssignmentHandler = new WorkerLocationAssignmentHandler(config, existingDbCounts);
                                workerLocationAssignmentHandler.ProcessSheets(useCase, validateOnly, configusername);
                                break;

                            case "worker supervisory org assignment":
                                var workerSupervisoryOrgAssignmentHandler = new WorkerSupervisoryOrgAssign(config, existingDbCounts);
                                workerSupervisoryOrgAssignmentHandler.ProcessSheets(useCase, validateOnly, configusername);

                                break;

                            case "worker pay group assignment":
                                var workerPayGroupAssignmentHandler = new WorkerPaygrouphandler(config, existingDbCounts);
                                workerPayGroupAssignmentHandler.ProcessSheets(useCase, validateOnly, configusername);
                                break;

                            case "worker cost center assignment":
                                var workerCostCenterAssignmentHandler = new WorkerCostCenterHandler(config, existingDbCounts);
                                workerCostCenterAssignmentHandler.ProcessSheets(useCase, validateOnly, configusername);
                                break;

                            case "worker department assignment":
                                var workerDepartmentAssignmentHandler = new WorkerDepartmentHandler(config, existingDbCounts);
                                workerDepartmentAssignmentHandler.ProcessSheets(useCase, validateOnly, configusername);
                                break;



                            default:
                                logger.Log($"No handler found for use case: {useCase}");
                                break;
                        }


                        logger.Log($"Completed process for {useCase}", level: "section");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error in {useCase}: {ex.Message}");

                        if (!validateOnly)
                        {
                            throw;
                        }
                    }
                }
            }

        }
        catch(Exception ex)
        {
            Console.WriteLine($"Error:{ex.Message}");
        }
        Console.ReadLine();
    } 


    private static void OnCancelKeyPress(object? sender, ConsoleCancelEventArgs e)
    {
        e.Cancel = true; 
    }

}

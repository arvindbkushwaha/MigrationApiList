using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using log4net;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.WindowsAzure.Storage.Queue;

namespace MigrationApiDemo
{
    public class MigrationApiDemo
    {
        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        private ListItemCollection itemsCollections;
        private readonly AzureBlob _blobContainingManifestFiles;
        private readonly SharePointMigrationTarget _target;
        private readonly SharePointMigrationSource _source;
        private readonly AzureCloudQueue _migrationApiQueue;
        private readonly TestDataProvider _testDataProvider;
        private ClientContext context = null;
        public MigrationApiDemo()
        {
            Log.Debug("Initiaing SharePoint connection.... ");

            _target = new SharePointMigrationTarget();
            _source = new SharePointMigrationSource();
            Log.Debug("Initiating Storage for test files, manifest en reporting queue");

            _blobContainingManifestFiles = new AzureBlob(
                ConfigurationManager.AppSettings["ManifestBlob.AccountName"],
                ConfigurationManager.AppSettings["ManifestBlob.AccountKey"],
                ConfigurationManager.AppSettings["ManifestBlob.ContainerName"]);

            var testFilesBlob = new AzureBlob(
                ConfigurationManager.AppSettings["SourceFilesBlob.AccountName"],
                ConfigurationManager.AppSettings["SourceFilesBlob.AccountKey"],
                ConfigurationManager.AppSettings["SourceFilesBlob.ContainerName"]);

            _testDataProvider = new TestDataProvider(testFilesBlob);

            _migrationApiQueue = new AzureCloudQueue(
                ConfigurationManager.AppSettings["ReportQueue.AccountName"],
                ConfigurationManager.AppSettings["ReportQueue.AccountKey"],
                ConfigurationManager.AppSettings["ReportQueue.QueueName"]);
        }

        public void ProvisionTestFiles()
        {
            string siteUrl = _source._tenantUrl + _source._siteName;
            context = SPData.GetOnlineContext(siteUrl, _source._username, _source._password);
            itemsCollections = _testDataProvider.ProvisionAndGetFiles(context, _source._listName);
            //getUser();
        }

        public void getUser()
        {
            ClientContext context = SPData.GetOnlineContext("https://cactusglobal.sharepoint.com/sites/medcomdev/", _source._username, _source._password);
            List oList = context.Web.Lists.GetByTitle("ClientInformation");
            ListItem item = oList.GetItemById(15);
            context.Load(item);
            context.ExecuteQuery();
            List<string> users = new List<string>();
            List<User> userIds = new List<User>();
            ListItemVersionCollection versions = item.Versions;
            context.Load(versions);
            context.ExecuteQuery();

            //foreach (FieldUserValue userValue in item["Editor"] as FieldUserValue[])
            //{
            //    //Console.WriteLine(userValue.LookupValue.ToString());
            //    users.Add(userValue.Email);
            //    User user = new User();
            //    user.Id = userValue.LookupId;
            //    user.name = userValue.LookupValue;
            //    user.emailId = userValue.Email;
            //    userIds.Add(user);
            //}
            FieldUserValue userValue = item["Editor"] as FieldUserValue;
            users.Add(userValue.Email);
            User user = new User();
            user.Id = userValue.LookupId;
            user.name = userValue.LookupValue;
            user.emailId = userValue.Email;
            userIds.Add(user);
            var results1 = SPData.getUserInfoUserProperties(context, userIds);
            foreach (var kvp in results1)
            {
                if (kvp.Value.ServerObjectIsNull.HasValue && !kvp.Value.ServerObjectIsNull.Value)
                {
                    Console.WriteLine(kvp.Key);
                    Console.WriteLine("---------------------------------");
                    foreach (var property in kvp.Value.FieldValues)
                    {
                        Console.WriteLine(string.Format("{0}: {1}",
                            property.Key.ToString(), property.Value != null ? property.Value.ToString() : ""));
                    }
                }
                else
                {
                    Console.WriteLine("User not found:" + kvp.Key);
                }
            }
            var results = SPData.GetMultipleUsersProfileProperties(context, userIds, results1);
            // Get the PeopleManager object and then get the target user's properties. 
            foreach (var kvp in results)
            {
                if (kvp.Value.ServerObjectIsNull.HasValue && !kvp.Value.ServerObjectIsNull.Value)
                {
                    Console.WriteLine(kvp.Value.DisplayName);
                    Console.WriteLine("---------------------------------");
                    foreach (var property in kvp.Value.UserProfileProperties)
                    {
                        Console.WriteLine(string.Format("{0}: {1}",
                            property.Key.ToString(), property.Value.ToString()));
                    }
                }
                else
                {
                    Console.WriteLine("User not found:" + kvp.Key);
                }
                Console.WriteLine("------------------------------");
                Console.WriteLine("          ");
            }
            


            Console.ReadLine();
        }
        public void CreateAndUploadMigrationPackage()
        {
            if (itemsCollections.Count > 0)
            {
                var manifestPackage = new ManifestPackage(_target, _source);
                var filesInManifestPackage = manifestPackage.GetManifestPackageFiles(itemsCollections, _source._listName, context);
                var blobContainingManifestFiles = _blobContainingManifestFiles;
                blobContainingManifestFiles.RemoveAllFiles();
                foreach (var migrationPackageFile in filesInManifestPackage)
                {
                    blobContainingManifestFiles.UploadFile(migrationPackageFile.Filename, migrationPackageFile.Contents);
                }
            }
            else
            {
                throw new Exception("No Items for migrate Package for, run ProvisionTestFiles() first!");
            }

        }

        /// <returns>Job Id</returns>
        public Guid StartMigrationJob()
        {
            var sourceFileContainerUrl = _testDataProvider.GetBlobUri();
            var manifestContainerUrl = _blobContainingManifestFiles.GetUri(
                SharedAccessBlobPermissions.Read
                | SharedAccessBlobPermissions.Write
                | SharedAccessBlobPermissions.List);

            var azureQueueReportUrl = _migrationApiQueue.GetUri(
                SharedAccessQueuePermissions.Read
                | SharedAccessQueuePermissions.Add
                | SharedAccessQueuePermissions.Update
                | SharedAccessQueuePermissions.ProcessMessages);

            return _target.StartMigrationJob(sourceFileContainerUrl, manifestContainerUrl, azureQueueReportUrl);
        }

        private void DownloadAndPersistLogFiles(Guid jobId)
        {
            foreach (var filename in _blobContainingManifestFiles.ListFilenames())
            {
                if (filename.StartsWith($"Import-{jobId}"))
                {
                    Log.Debug($"Downloaded logfile {filename}");
                    //File.WriteAllBytes(filename, _blobContainingManifestFiles.DownloadFile(filename));
                }
            }
        }

        public async Task MonitorMigrationApiQueue(Guid jobId)
        {
            while (true)
            {
                var message = await _migrationApiQueue.GetMessageAsync<UpdateMessage>();
                if (message == null)
                {
                    await Task.Delay(TimeSpan.FromSeconds(1));
                    continue;
                }

                switch (message.Event)
                {
                    case "JobEnd":
                        Log.Info($"Migration Job Ended {message.FilesCreated:0.} files created, {message.TotalErrors:0.} errors.!");
                        DownloadAndPersistLogFiles(jobId); // save log files to disk
                        Console.WriteLine("Press ctrl+c to exit");
                        return;
                    case "JobStart":
                        Log.Info("Migration Job Started!");
                        break;
                    case "JobProgress":
                        Log.Debug($"Migration Job in progress, {message.FilesCreated:0.} files created, {message.TotalErrors:0.} errors.");
                        break;
                    case "JobQueued":
                        Log.Info("Migration Job Queued...");
                        break;
                    case "JobWarning":
                        Log.Warn($"Migration Job warning {message.Message}");
                        break;
                    case "JobError":
                        Log.Error($"Migration Job error {message.Message}");
                        break;
                    default:
                        Log.Warn($"Unknown Job Status: {message.Event}, message {message.Message}");
                        break;

                }
            }
        }
    }
}
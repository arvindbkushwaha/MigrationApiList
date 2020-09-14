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

        private ListItemCollection _sourceItemsCollections;
        private List<ListItemCollection> _listDestinationItemsCollections = null;
        private List<ListItemCollection> _listSourceItemsCollections = null;
        private readonly AzureBlob _blobContainingManifestFiles;
        private readonly SharePointMigrationTarget _target;
        private readonly SharePointMigrationSource _source;
        private readonly AzureCloudQueue _migrationApiQueue;
        private readonly TestDataProvider _testDataProvider;
        private ClientContext _sourceContext = null;
        private ClientContext _destinationContext = null;
        private Boolean _isModifiedQueryEnabled = ConfigurationManager.AppSettings["IsModifiedQueryEnabled"] == "Yes" ? true : false;
        public MigrationApiDemo()
        {
            Log.Debug("Initiaing SharePoint connection.... ");

            _target = new SharePointMigrationTarget();
            _source = new SharePointMigrationSource();
            Log.Debug("Initiating Storage for test files, manifest en reporting queue");

            _blobContainingManifestFiles = new AzureBlob(
                ConfigurationManager.AppSettings["ManifestBlob.AccountName"],
                ConfigurationManager.AppSettings["ManifestBlob.AccountKey"],
                ConfigurationManager.AppSettings["ManifestBlob.ContainerName"] + DateTime.Now.ToString("yyyyMMddHHmmss"));

            var testFilesBlob = new AzureBlob(
                ConfigurationManager.AppSettings["SourceFilesBlob.AccountName"],
                ConfigurationManager.AppSettings["SourceFilesBlob.AccountKey"],
                ConfigurationManager.AppSettings["SourceFilesBlob.ContainerName"] + DateTime.Now.ToString("yyyyMMddHHmmss"));

            _testDataProvider = new TestDataProvider(testFilesBlob);

            _migrationApiQueue = new AzureCloudQueue(
                ConfigurationManager.AppSettings["ReportQueue.AccountName"],
                ConfigurationManager.AppSettings["ReportQueue.AccountKey"],
                ConfigurationManager.AppSettings["ReportQueue.QueueName"] + DateTime.Now.ToString("yyyyMMddHHmmss"));
        }

        public void ProvisionTestFiles()
        {
            string siteUrl = _source._tenantUrl + _source._siteName;
            _sourceContext = SPData.GetOnlineContext(siteUrl, _source._username, _source._password);
            if (_source._listName != "ProjectInformationCT")
            {
                _sourceItemsCollections = _testDataProvider.ProvisionAndGetFiles(_sourceContext, _source._listName, true);
            }
            else
            {
                _listSourceItemsCollections = _testDataProvider.GetProjectInformationData(_sourceContext, _source._listName, true);
                //_sourceItemsCollections = _listSourceItemsCollections[0];
            }

            string destionationUrl = _target._tenantUrl + _target.SiteName;
            _destinationContext = SPData.GetOnlineContext(destionationUrl, _target._username, _target._password);
            //Check is fetched data with modified date or Id
            if (!_isModifiedQueryEnabled)
            {
                // fetched destionation data with Id
                _listDestinationItemsCollections = new List<ListItemCollection>();
                ListItemCollection destinationItemsCollections = _testDataProvider.ProvisionAndGetFiles(_destinationContext, _target.ListName, false);
                _destinationContext.Load(destinationItemsCollections);
                if (destinationItemsCollections.Count > 0)
                {
                    _listDestinationItemsCollections.Add(destinationItemsCollections);
                }

            }
            else
            {
                //fetched destionation data based on source Items Id
                _listDestinationItemsCollections = _testDataProvider.GetDestinationFiles(_destinationContext, _target.ListName, _sourceItemsCollections);
            }
            //getUser();
        }

        public void CreateAndUploadMigrationPackage()
        {
            if ((_sourceItemsCollections != null && _sourceItemsCollections.Count > 0) || (_listSourceItemsCollections != null && _listSourceItemsCollections.Count > 0))
            {
                var manifestPackage = new ManifestPackage(_target, _source);
                var filesInManifestPackage = manifestPackage.GetManifestPackageFiles(_sourceItemsCollections, _listSourceItemsCollections, _listDestinationItemsCollections, _source._listName, _sourceContext);
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
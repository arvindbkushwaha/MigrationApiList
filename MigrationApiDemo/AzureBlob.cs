using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using log4net;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Auth;
using Microsoft.WindowsAzure.Storage.Blob;

namespace MigrationApiDemo
{
    public class AzureBlob
    {
        private readonly string _containerName;
        private CloudBlobContainer _containerReference;

        private static readonly ILog Log = LogManager.GetLogger(typeof(AzureBlob));
        /// <summary>
        /// This method is to get the azure blob.
        /// </summary>
        /// <param name="accountName"></param>
        /// <param name="accountKey"></param>
        /// 
        /// /// <param name="containerName"></param>
        public AzureBlob(string accountName, string accountKey, string containerName)
        {
            _containerName = containerName;

            var storageCredentials = new StorageCredentials(accountName, accountKey);
            var cloudStorageAccount = new CloudStorageAccount(storageCredentials, true);

            SetContainerReference(cloudStorageAccount);
        }
        /// <summary>
        /// This method is used to set the container references
        /// </summary>
        /// <param name="cloudStorageAccount"></param>
        private void SetContainerReference(CloudStorageAccount cloudStorageAccount)
        {
            var cloudBlobClient = cloudStorageAccount.CreateCloudBlobClient();
            _containerReference = cloudBlobClient.GetContainerReference(_containerName);
            _containerReference.CreateIfNotExists();
        }
        /// <summary>
        /// This method is used to upload the files.
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="contents"></param>
        public void UploadFile(string filename, byte[] contents)
        {
            var blobReference = _containerReference.GetBlockBlobReference(filename);
            blobReference.UploadFromByteArray(contents, 0, contents.Length);
        }
        /// <summary>
        /// This method is used to remove all the files from azure.
        /// </summary>
        public void RemoveAllFiles()
        {
            var blobs = _containerReference.ListBlobs();
            foreach (var blockBlob in blobs.OfType<CloudBlockBlob>())
            {
                blockBlob.Delete();
            }
        }
        /// <summary>
        /// This method is used to GetUri of SharedAccessBlobPermission.
        /// </summary>
        /// <param name="permissions"></param>
        /// <returns></returns>
        public Uri GetUri(SharedAccessBlobPermissions permissions)
        {
            var policy = new SharedAccessBlobPolicy
            {
                SharedAccessExpiryTime = DateTime.UtcNow.AddDays(31.0),
                Permissions = permissions
            };
            return new Uri(_containerReference.Uri, _containerReference.GetSharedAccessSignature(policy) + "&comp=list&restype=container");
        }
        /// <summary>
        /// This method is used to ge the filenames.
        /// </summary>
        /// <returns></returns>
        public ICollection<string> ListFilenames()
        {
            var blobs = _containerReference.ListBlobs();
            return blobs.OfType<CloudBlockBlob>().Select(x => x.Name).ToList();
        }
        /// <summary>
        /// This method is used to download the file.
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public byte[] DownloadFile(string filename)
        {
            try
            {
                var blobReference = _containerReference.GetBlockBlobReference(filename);

                using (var memoryStream = new MemoryStream())
                {
                    blobReference.DownloadToStream(memoryStream);
                    memoryStream.Position = 0;
                    return memoryStream.ToArray();
                }
            }
            catch (Exception ex)
            {
                Log.Error("Unexpected Exception while downloading file from Azure BLOB", ex);

                throw;
            }
        }
    }
}
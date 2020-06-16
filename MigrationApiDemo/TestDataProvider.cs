using System;
using System.Collections.Generic;
using System.Configuration;
using System.Reflection;
using System.Security;
using System.Text;
using log4net;
using Microsoft.SharePoint.Client;
using Microsoft.WindowsAzure.Storage.Blob;

namespace MigrationApiDemo
{
    public  class TestDataProvider
    {
        private  readonly AzureBlob _azureBlob;
        private  readonly ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        public TestDataProvider(AzureBlob azureBlob)
        {
            _azureBlob = azureBlob;
        }

        public ListItemCollection ProvisionAndGetFiles(ClientContext context, string listName)
        {
            List list = context.Web.Lists.GetByTitle(listName);
            CamlQuery query = new CamlQuery();
            query.ViewXml = @"<View></View>";
            ListItemCollection listItemCollections = list.GetItems(query);
            context.Load(listItemCollections);
            context.ExecuteQuery();
            return listItemCollections;
        }

        public  Uri GetBlobUri()
        {
            return _azureBlob.GetUri(SharedAccessBlobPermissions.Read | SharedAccessBlobPermissions.List);
        }
    }
}
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Security;
using Microsoft.SharePoint.Client;

namespace MigrationApiDemo
{
    public class SharePointMigrationTarget
    {
        public Uri _tenantUrl;
        private readonly string _username;
        private readonly string _password;
        public Dictionary<String, LookupList> lookupListDic = new Dictionary<string, LookupList>();
        public ClientContext _client;
        public FieldCollection _fields;
        public readonly string SiteName;
        public readonly string listName;
        public readonly string ListName;
        public Guid ListId;
        public Guid WebId;
        public Guid RootFolderId;
        public Guid RootFolderParentId;

        public SharePointMigrationTarget() : this(
            new Uri(ConfigurationManager.AppSettings["SharePoint.TenantUrl"]),
            ConfigurationManager.AppSettings["SharePoint.DestinationSiteName"],
            ConfigurationManager.AppSettings["SharePoint.DestinationUsername"],
            ConfigurationManager.AppSettings["SharePoint.DestinationPassword"],
            ConfigurationManager.AppSettings["SharePoint.DestinationListName"])
        {
        }

        public SharePointMigrationTarget(Uri tenantUrl, string siteName, string username, string password, string listName)
        {
            _tenantUrl = tenantUrl;
            SiteName = siteName;
            _username = username;
            _password = password;
            ListName = listName;
            Initialize();
        }

        private void Initialize()
        {
            var securePassword = new SecureString();
            foreach (var c in _password) securePassword.AppendChar(c);

            _client = new ClientContext($"{_tenantUrl}/{SiteName}/");
            _client.Credentials = new SharePointOnlineCredentials(_username, securePassword);

            var _list = _client.Web.Lists.GetByTitle(ListName);
            _client.Load(_list, x => x.RootFolder);
            _client.ExecuteQuery();
            var folder = _list.RootFolder;

            _client.Load(_client.Site, x => x.Id);
            _client.Load(_client.Web, x => x.Id);
            _client.Load(_list, x => x.Id);
            _client.Load(folder, x => x.UniqueId);
            _client.Load(folder, x => x.ParentFolder.UniqueId);
            _client.Load(_list, x => x.Fields);
            _client.ExecuteQuery();
            ListId = _list.Id;
            WebId = _client.Web.Id;
            RootFolderId = folder.UniqueId;
            RootFolderParentId = folder.ParentFolder.UniqueId;
            _fields = _list.Fields;
            foreach (Field field in _fields)
            {
                string fieldType = field.TypeAsString;
                if (!field.Hidden && !field.ReadOnlyField && (fieldType == "Lookup" || fieldType == "LookupMulti"))
                {
                    LookupList lp = new LookupList();
                    var lookupField = _client.CastTo<FieldLookup>(field);
                    _client.Load(lookupField);
                    var lookupListId = new Guid(lookupField.LookupList);
                    lp.listId = lookupListId.ToString();
                    var lookupList = _client.Web.Lists.GetById(lookupListId);
                    _client.Load(lookupList);
                    _client.ExecuteQuery();
                    CamlQuery query = new CamlQuery();
                    query.ViewXml = @"<View></View>";
                    ListItemCollection colls = lookupList.GetItems(query);
                    _client.Load(colls);
                    _client.ExecuteQuery();
                    lp.itemArray = colls;
                    lookupListDic.Add(field.InternalName, lp);

                }
            }
        }
        public Guid StartMigrationJob(Uri sourceFileContainerUrl, Uri manifestContainerUrl, Uri azureQueueReportUrl)
        {
            var result = _client.Site.CreateMigrationJob(WebId, sourceFileContainerUrl.ToString(), manifestContainerUrl.ToString(), azureQueueReportUrl.ToString());
            _client.ExecuteQuery();
            return result.Value;
        }
    }
}
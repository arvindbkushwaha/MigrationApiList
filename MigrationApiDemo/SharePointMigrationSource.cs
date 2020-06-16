using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MigrationApiDemo
{
    public class SharePointMigrationSource
    {
        public Uri _tenantUrl;
        public readonly string _username;
        public readonly string _password;
        public readonly string _siteName;
        public readonly string _listName;
        public readonly string _siteUrl;

        public SharePointMigrationSource() : this(
            new Uri(ConfigurationManager.AppSettings["SharePoint.TenantUrl"]),
            ConfigurationManager.AppSettings["SharePoint.SourceSiteName"],
            ConfigurationManager.AppSettings["SharePoint.SourceUsername"],
            ConfigurationManager.AppSettings["SharePoint.SourcePassword"],
            ConfigurationManager.AppSettings["SharePoint.SourceListName"])
        {
        }
        public SharePointMigrationSource(Uri tenantUrl, string siteName, string username, string password, string listName)
        {
            _tenantUrl = tenantUrl;
            _siteName = siteName;
            _username = username;
            _password = password;
            _listName = listName;
            _siteUrl = _tenantUrl + _siteName;
        }
    }
}

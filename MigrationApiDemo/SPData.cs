using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace MigrationApiDemo
{
    class SPData
    {
        public static ClientContext GetOnlineContext(string siteUrl, string userName, string password)
        {
            SecureString securePassword = GetPassword(password);
            ClientContext context = new ClientContext(siteUrl);
            context.Credentials = new SharePointOnlineCredentials(userName, securePassword);
            return context;
        }

        public static FieldCollection GetFields(ClientContext context, string listName)
        {
            List list = context.Web.Lists.GetByTitle(listName);
            FieldCollection fields = list.Fields;
            context.Load(fields);
            context.ExecuteQuery();
            return fields;
        }
        private static SecureString GetPassword(string password)
        {
            SecureString securePassword = new SecureString();
            foreach (char c in password)
            {
                securePassword.AppendChar(c);
            }
            return securePassword;
        }
        public static Dictionary<string, PersonProperties> GetMultipleUsersProfileProperties(ClientContext context, List<User> Users, Dictionary<int,ListItem> UsersInfo)
        {
            var results = new Dictionary<string, PersonProperties>();
            PeopleManager peopleManager = new PeopleManager(context);
            foreach (var user in Users)
            {
                try
                {
                    ListItem item = UsersInfo[user.Id];
                    if (!(String.IsNullOrEmpty(user.emailId)))
                    {

                        string loginName = item["Name"] != null ? item["Name"].ToString() : string.Empty;  //claim format login name
                        if (!string.IsNullOrEmpty(loginName))
                        {
                            var personProperties = peopleManager.GetPropertiesFor(loginName);
                            context.Load(personProperties, p => p.AccountName, p => p.DisplayName,
                                               p => p.UserProfileProperties);
                            results.Add(user.emailId, personProperties);
                        }

                    }
                }
                catch(Exception e)
                {

                }
            }
            context.ExecuteQuery();
            return results;
        }
        public static PersonProperties GetSingleUsersProfileProperties(ClientContext context, string emailId)
        {
            // Get the PeopleManager object and then get the target user's properties. 
            PeopleManager peopleManager = new PeopleManager(context);
            PersonProperties userProperties = peopleManager.GetPropertiesFor("i:0#.f|membership|" + emailId);

            // This request load the AccountName and user's all other Profile Properties 
            context.Load(userProperties, p => p.AccountName, p => p.UserProfileProperties);
            context.ExecuteQuery();
            return userProperties;
        }
        public static Dictionary<int, ListItem> getUserInfoUserProperties(ClientContext context, List<User> users)
        {
            //var context = SPData.GetOnlineContext("", user)
            var results = new Dictionary<int, ListItem>();
            var userInfoList = context.Web.SiteUserInfoList;
            foreach (var user in users)
            {
                var item = userInfoList.GetItemById(user.Id);
                context.Load(item);
                try
                {
                    context.ExecuteQuery();
                    Console.WriteLine("SID: " + item["Sid"]);
                    results.Add(user.Id, item);
                }
                catch
                {
                    results.Add(user.Id, item);
                }
            }
            return results;
        }

    }
}

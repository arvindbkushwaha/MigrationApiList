using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MigrationApiDemo
{
    public class Common
    {
        public static DateTime ValidXMLDate(string dt)
        {
            var newDT = dt.Split('Z');
            return Convert.ToDateTime(newDT[0]);
        }
        public static string ValidXMLDate(ListItem item, string internalName)
        {
            string itemRegDate = string.Empty;
            if (item[internalName] != null)
            {
                DateTime date = Convert.ToDateTime(item[internalName]);
                itemRegDate = date.ToString("yyyy-MM-ddTHH:mm:ss");
            }
            return itemRegDate;
        }
        public static string ValidXMLDate(Dictionary<string, Object> item, string internalName)
        {
            string itemRegDate = string.Empty;
            if (item[internalName] != null)
            {
                DateTime date = Convert.ToDateTime(item[internalName]);
                itemRegDate = date.ToString("yyyy-MM-ddTHH:mm:ss");
            }
            return itemRegDate;
        }

        public static string GetSingleId(List<User> users, ListItem item, string internalName, Boolean isUserInfoRequired)
        {
            FieldUserValue userValue = item[internalName] as FieldUserValue;
            if (userValue != null && !(users.Any(a => a.Id == userValue.LookupId)))
            {
                User user = new User();
                user.Id = userValue.LookupId;
                user.name = userValue.LookupValue;
                user.emailId = userValue.Email;
                users.Add(user);
            }
            string result = null;
            if (isUserInfoRequired)
            {
                result = userValue != null ? userValue.LookupId.ToString() + ";#;UserInfo" : null;
            }
            else
            {
                result = userValue.LookupId.ToString();
            }
            return result;
        }

        public static string GetMultipleId(List<User> users, ListItem item, string internalName)
        {
            string result = null;
            List<int> Ids = new List<int>();
            if (item.FieldValues.ContainsKey(internalName) && item[internalName] != null)
            {
                foreach (FieldUserValue userValue in item[internalName] as FieldUserValue[])
                {
                    Ids.Add(userValue.LookupId);
                    if (!(users.Any(a => a.Id == userValue.LookupId)))
                    {
                        User user = new User();
                        user.Id = userValue.LookupId;
                        user.name = userValue.LookupValue;
                        user.emailId = userValue.Email;
                        users.Add(user);
                    }
                }
                if (Ids.Count > 0)
                {
                    result += string.Join(";#;#", Ids.ToArray());
                    result += ";#;UserInfo";
                }
            }
            return result;
        }

        public static string GetSingleId(List<User> users, Dictionary<string, Object> item, string internalName, Boolean isUserInfoRequired)
        {
            FieldUserValue userValue = item[internalName] as FieldUserValue;
            if (userValue != null && !(users.Any(a => a.Id == userValue.LookupId)))
            {
                User user = new User();
                user.Id = userValue.LookupId;
                user.name = userValue.LookupValue;
                user.emailId = userValue.Email;
                users.Add(user);
            }
            string result = "";
            if (isUserInfoRequired)
            {
                result = userValue != null ? userValue.LookupId.ToString() + ";UserInfo" : null;
            }
            else
            {
                result = userValue.LookupId.ToString();
            }
            return result;
        }
        public static string GetMultipleId(List<User> users, Dictionary<string, object> item, string internalName)
        {
            string result = null;
            List<int> Ids = new List<int>();
            if (item.ContainsKey(internalName) && item[internalName] != null)
            {

                foreach (FieldUserValue userValue in item[internalName] as FieldUserValue[])
                {
                    Ids.Add(userValue.LookupId);
                    if (!(users.Any(a => a.Id == userValue.LookupId)))
                    {
                        User user = new User();
                        user.Id = userValue.LookupId;
                        user.name = userValue.LookupValue;
                        user.emailId = userValue.Email;
                        users.Add(user);
                    }
                }
                if (Ids.Count > 0)
                {
                    result += string.Join(";#;#", Ids.ToArray());
                    result += ";#;UserInfo";
                }
            }
            return result;
        }
        public static List<ListItem> GetResourceCategorization(ClientContext context)
        {
            var listName = "ResourceCategorization";
            List<ListItem> resourceCat = new List<ListItem>();
            List list = context.Web.Lists.GetByTitle(listName);
            CamlQuery query = new CamlQuery();
            query.ViewXml = @"<View></View>";
            ListItemCollection listItems = list.GetItems(query);
            context.Load(listItems);
            context.ExecuteQuery();
            foreach (var item in listItems)
            {
                resourceCat.Add(item);
            }
            return resourceCat;
        }
        public static ListItem GetActiveItem(List<ListItem> resourceCat, string internalName, int Id)
        {
            ListItem litem = null;
            foreach (var item in resourceCat)
            {
                FieldUserValue userValue = item[internalName] as FieldUserValue;
                if (userValue.LookupId == Id)
                {
                    litem = item;
                    break;
                }
            }
            return litem;
        }
        public static string GetLookUpId(ListItem item, Field field, Dictionary<string, LookupList> lookupDictonary, Boolean isSingleLookup)
        {
            string lookupId = String.Empty;
            string listId = lookupDictonary[field.InternalName].listId;
            if (isSingleLookup)
            {
                FieldLookupValue singleLook = (item[field.InternalName] as FieldLookupValue);
                lookupId = singleLook != null ? singleLook.LookupId + ";#;" + listId : string.Empty;
            }
            else
            {
                var lookupIds = new List<int>();
                var MultipleValues = (item[field.InternalName] as FieldLookupValue[]);
                for (int count = 0; count <= MultipleValues.Length - 1; count++)
                {
                    FieldLookupValue itemValue = MultipleValues[count];
                    if (MultipleValues.Length == 1)
                    {
                        lookupId += itemValue.LookupId + ";#";
                    }
                    else if (MultipleValues.Length > 1)
                    {
                        if (count == MultipleValues.Length - 1)
                        {
                            lookupId += itemValue.LookupId + ";#";
                        }
                        else
                        {
                            lookupId += itemValue.LookupId + ";#;#";
                        }
                    }
                }
                if (MultipleValues.Length > 0)
                {
                    lookupId += ";" + listId;
                }

            }
            return lookupId;
        }

        public static string GetLookUpId(ListItem item, string fieldInternalName, Dictionary<string, LookupList> lookupDictonary, Boolean isSingleLookup)
        {
            string lookupId = String.Empty;
            string listId = lookupDictonary[fieldInternalName].listId;
            if (isSingleLookup)
            {
                FieldLookupValue singleLook = (item[fieldInternalName] as FieldLookupValue);
                lookupId = singleLook != null ? singleLook.LookupId + ";#;" + listId : string.Empty;
            }
            else
            {
                var lookupIds = new List<int>();
                var MultipleValues = (item[fieldInternalName] as FieldLookupValue[]);
                for (int count = 0; count <= MultipleValues.Length - 1; count++)
                {
                    FieldLookupValue itemValue = MultipleValues[count];
                    if (MultipleValues.Length == 1)
                    {
                        lookupId += itemValue.LookupId + ";#";
                    }
                    else if (MultipleValues.Length > 1)
                    {
                        if (count == MultipleValues.Length - 1)
                        {
                            lookupId += itemValue.LookupId + ";#";
                        }
                        else
                        {
                            lookupId += itemValue.LookupId + ";#;#";
                        }
                    }
                }
                if (MultipleValues.Length > 0)
                {
                    lookupId += ";" + listId;
                }

            }
            return lookupId;
        }
        public static string GetLookUpId(Dictionary<string, object> item, Field field, Dictionary<string, LookupList> lookupDictonary, Boolean isSingleLookup)
        {
            string lookupId = String.Empty;
            string listId = lookupDictonary[field.InternalName].listId;
            if (isSingleLookup)
            {
                FieldLookupValue singleLook = (item[field.InternalName] as FieldLookupValue);
                lookupId = singleLook != null ? singleLook.LookupId + ";#;" + listId : string.Empty;
            }
            else
            {
                var lookupIds = new List<int>();
                var MultipleValues = (item[field.InternalName] as FieldLookupValue[]);
                for (int count = 0; count <= MultipleValues.Length - 1; count++)
                {
                    FieldLookupValue itemValue = MultipleValues[count];
                    if (MultipleValues.Length == 1)
                    {
                        lookupId += itemValue.LookupId + ";#";
                    }
                    else if (MultipleValues.Length > 1)
                    {
                        if (count == MultipleValues.Length - 1)
                        {
                            lookupId += itemValue.LookupId + ";#";
                        }
                        else
                        {
                            lookupId += itemValue.LookupId + ";#;#";
                        }
                    }
                }
                if (MultipleValues.Length > 0)
                {
                    lookupId += ";" + listId;
                }

            }
            return lookupId;
        }

    }
}
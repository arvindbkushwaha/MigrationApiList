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
    public class TestDataProvider
    {
        private readonly AzureBlob _azureBlob;
        private readonly ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        private Int32 _fromId = Convert.ToInt32(ConfigurationManager.AppSettings["FromId"]);
        private Int32 _toId = Convert.ToInt32(ConfigurationManager.AppSettings["ToId"]);
        private Boolean _isLimitedItem = ConfigurationManager.AppSettings["IsLimitedItemRequired"] == "Yes" ? true : false;
        private Boolean _isModifiedQueryEnabled = ConfigurationManager.AppSettings["IsModifiedQueryEnabled"] == "Yes" ? true : false;
        private string _fromTime = ConfigurationManager.AppSettings["FromTime"];
        private string _toTime = ConfigurationManager.AppSettings["ToTime"];
        public TestDataProvider(AzureBlob azureBlob)
        {
            _azureBlob = azureBlob;
        }

        public ListItemCollection ProvisionAndGetFiles(ClientContext context, string listName, Boolean isSourceItemsRequired)
        {
            string viewFields = string.Empty;
            List list = context.Web.Lists.GetByTitle(listName);
            CamlQuery query = new CamlQuery();
            string whereQuery = string.Empty;
            if (_isModifiedQueryEnabled)
            {
                whereQuery = "<Query>" +
                                   "<Where>" +
                                      "<And>" +
                                         "<Geq>" +
                                            "<FieldRef Name='Modified' />" +
                                            "<Value IncludeTimeValue='TRUE' Type='DateTime'>" + _fromTime + "</Value>" +
                                         "</Geq>" +
                                         "<Leq>" +
                                            "<FieldRef Name='Modified' />" +
                                            "<Value IncludeTimeValue='TRUE' Type='DateTime'>" + _toTime + "</Value>" +
                                         "</Leq>" +
                                      "</And>" +
                                   "</Where>" +
                                   "<OrderBy>" +
                                    "<FieldRef Name='ID' Ascending='True' />" +
                                   "</OrderBy>" +
                                "</Query>";
            }
            else
            {
                whereQuery = "<Query>" +
                                 "<Where>" +
                                     "<And>" +
                                            "<Geq><FieldRef Name='ID'></FieldRef><Value Type='Counter'>" + _fromId + "</Value></Geq>" +
                                            "<Leq><FieldRef Name='ID'></FieldRef><Value Type='Counter'>" + _toId + "</Value></Leq>" +
                                     "</And>" +
                                  "</Where>" +
                                  "<OrderBy>" +
                                    "<FieldRef Name='ID' Ascending='True' />" +
                                 "</OrderBy>" +
                              "</Query>";
            }

            if (isSourceItemsRequired && listName == "Schedules")
            {
                viewFields = string.Concat(
                     "<FieldRef Name='Id' />",
                     "<FieldRef Name='Title' />",
                     "<FieldRef Name='ActiveCA' />",
                     "<FieldRef Name='Actual_x0020_Start_x0020_Date' />",
                     "<FieldRef Name='Actual_x0020_End_x0020_Date' />",
                     "<FieldRef Name='AllowCompletion' />",
                     "<FieldRef Name='AssignedTo' />",
                     "<FieldRef Name='AssignedToText' />",
                     "<FieldRef Name='CentralAllocationDone' />",
                     "<FieldRef Name='Comments' />",
                     "<FieldRef Name='DisableCascade' />",
                     "<FieldRef Name='DueDate' />",
                     "<FieldRef Name='Entity' />",
                     "<FieldRef Name='ExpectedTime' />",
                     "<FieldRef Name='FinalDocSubmit' />",
                     "<FieldRef Name='IsCentrallyAllocated' />",
                     "<FieldRef Name='IsRated' />",
                     "<FieldRef Name='Milestone' />",
                     "<FieldRef Name='NextTasks' />",
                     "<FieldRef Name='ParentSlot' />",
                     "<FieldRef Name='PreviousAssignedUser' />",
                     "<FieldRef Name='PreviousTaskClosureDate' />",
                     "<FieldRef Name='PrevTasks' />",
                     "<FieldRef Name='ProjectCode' />",
                     "<FieldRef Name='Rated' />",
                     "<FieldRef Name='SkillLevel' />",
                     "<FieldRef Name='StartDate' />",
                     "<FieldRef Name='SubMilestones' />",
                     "<FieldRef Name='Task' />",
                     "<FieldRef Name='Status' />",
                     "<FieldRef Name='TaskComments' />",
                     "<FieldRef Name='TATBusinessDays' />",
                     "<FieldRef Name='TATStatus' />",
                     "<FieldRef Name='TimeSpent' />",
                     "<FieldRef Name='TimeSpentPerDay' />",
                     "<FieldRef Name='TimeZone' />",
                     "<FieldRef Name='Modified' />",
                     "<FieldRef Name='Created' />",
                     "<FieldRef Name='Author' />",
                     "<FieldRef Name='Editor' />",
                     "<FieldRef Name='ContentType' />",
                     "<FieldRef Name='ContentTypeId' />",
                     "<FieldRef Name='_UIVersionString' />",
                     "<FieldRef Name='FileRef' />",
                     "<FieldRef Name='FileLeafRef' />",
                     "<FieldRef Name='FileDirRef' />",
                     "<FieldRef Name='FSObjType' />"
                    );
                query.ViewXml = @"<View Scope='RecursiveAll'><ViewFields>" + viewFields + "</ViewFields>" + whereQuery + "</View>";
                //306591
            }
            else if (isSourceItemsRequired && _isLimitedItem)
            {
                query.ViewXml = @"<View>" + whereQuery + "</View>";
            }
            else if (isSourceItemsRequired)
            {
                query.ViewXml = @"<View></View>";
            }
            else if (!isSourceItemsRequired && listName == "SchedulesCT")
            {
                viewFields = string.Concat(
                    "<FieldRef Name='Id' />",
                    "<FieldRef Name='UniqueId' />");
                query.ViewXml = @"<View Scope='RecursiveAll'><ViewFields>" + viewFields + "</ViewFields>" + whereQuery + "</View>";
            }
            else if (!isSourceItemsRequired && !_isLimitedItem)
            {
                query.ViewXml = @"<View></View>";
            }
            else
            {
                viewFields = string.Concat(
                    "<FieldRef Name='Id' />",
                    "<FieldRef Name='UniqueId' />");
                query.ViewXml = @"<View><ViewFields>" + viewFields + "</ViewFields>" + whereQuery + "</View>";
            }
            ListItemCollection listItemCollections = list.GetItems(query);
            context.Load(listItemCollections);
            context.ExecuteQuery();
            return listItemCollections;
        }

        public List<ListItemCollection> GetDestinationFiles(ClientContext context, string listName, ListItemCollection sourceItemCollections)
        {
            Int32 firstId = 0;
            Int32 lastId = 0;
            Int32 adjustedId = 0;
            Int32 addId = 3999;
            string viewFields = string.Empty;
            List list = context.Web.Lists.GetByTitle(listName);
            CamlQuery query = new CamlQuery();
            string whereQuery = string.Empty;
            List<ListItemCollection> allItems = new List<ListItemCollection>();
            if (sourceItemCollections.Count > 0)
            {
                viewFields = string.Concat(
                    "<FieldRef Name='Id' />",
                    "<FieldRef Name='UniqueId' />");
                firstId = Convert.ToInt32(sourceItemCollections[0]["ID"]);
                lastId = Convert.ToInt32(sourceItemCollections[sourceItemCollections.Count - 1]["ID"]);
                while (lastId - firstId > -1)
                {
                    if (lastId - firstId >= addId)
                    {
                        adjustedId = firstId + addId;
                    }
                    else
                    {
                        adjustedId = lastId;
                    }

                    whereQuery = "<Query>" +
                                      "<Where>" +
                                        "<And>" +
                                               "<Geq><FieldRef Name='ID'></FieldRef><Value Type='Counter'>" + firstId + "</Value></Geq>" +
                                               "<Leq><FieldRef Name='ID'></FieldRef><Value Type='Counter'>" + adjustedId + "</Value></Leq>" +
                                        "</And>" +
                                     "</Where>" +
                                     "<OrderBy>" +
                                       "<FieldRef Name='ID' Ascending='True' />" +
                                    "</OrderBy>" +
                                 "</Query>";
                    if (listName == "SchedulesCT")
                    {
                        query.ViewXml = @"<View Scope='RecursiveAll'><ViewFields>" + viewFields + "</ViewFields>" + whereQuery + "</View>";
                    }
                    else
                    {
                        query.ViewXml = @"<View><ViewFields>" + viewFields + "</ViewFields>" + whereQuery + "</View>";
                    }
                    ListItemCollection listItemCollections = list.GetItems(query);
                    context.Load(listItemCollections);
                    context.ExecuteQuery();
                    if (listItemCollections.Count > 0)
                    {
                        allItems.Add(listItemCollections);
                    }
                    //Increase the firstId by addding adjusted ID.
                    firstId = adjustedId + 1;

                }
            }
            return allItems;
        }
        public Uri GetBlobUri()
        {
            return _azureBlob.GetUri(SharedAccessBlobPermissions.Read | SharedAccessBlobPermissions.List);
        }
    }
}
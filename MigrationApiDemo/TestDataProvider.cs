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

        public List<ListItemCollection> GetProjectInformationData(ClientContext context, string listName, Boolean isSourceItemsRequired)
        {
            string viewFields = string.Empty;
            string viewFields2 = string.Empty;
            List list = context.Web.Lists.GetByTitle(listName);
            CamlQuery query = new CamlQuery();
            CamlQuery query2 = new CamlQuery();
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
            viewFields = string.Concat(
                     "<FieldRef Name='Id' />",
                     "<FieldRef Name='Title' />",
                     "<FieldRef Name='ActualEndDate' />",
                     "<FieldRef Name='ActualStartDate' />",
                     "<FieldRef Name='AnnotationBinder' />",
                     "<FieldRef Name='ArchivalStatus' />",
                     "<FieldRef Name='AuditCheckList' />",
                     "<FieldRef Name='BillingEntity' />",
                     "<FieldRef Name='BusinessVertical' />",
                     "<FieldRef Name='ClientLegalEntity' />",
                     "<FieldRef Name='CMLevel1' />",
                     "<FieldRef Name='CMLevel2' />",
                     "<FieldRef Name='CommentsMT' />",
                     "<FieldRef Name='ConferenceJournal' />",
                     "<FieldRef Name='DeliverableType' />",
                     "<FieldRef Name='DeliveryBackup' />",
                     "<FieldRef Name='DeliveryLevel1' />",
                     "<FieldRef Name='DeliveryLevel2' />",
                     "<FieldRef Name='DescriptionMT' />",
                     "<FieldRef Name='FinanceAuditCheckList' />",
                     "<FieldRef Name='Indication' />",
                     "<FieldRef Name='IsApproved' />",
                     "<FieldRef Name='IsPubSupport' />",
                     "<FieldRef Name='IsStandard' />",
                     "<FieldRef Name='JournalSelectionDate' />",
                     "<FieldRef Name='JournalSelectionURL' />",
                     "<FieldRef Name='LastSubmissionDate' />",
                     "<FieldRef Name='Milestone' />",
                     "<FieldRef Name='Milestones' />",
                     "<FieldRef Name='Molecule' />",
                     "<FieldRef Name='OvernightRequest' />",
                     "<FieldRef Name='PageCount' />",
                     "<FieldRef Name='POC' />",
                     "<FieldRef Name='PrevStatus' />",
                     "<FieldRef Name='PrimaryPOC' />",
                     "<FieldRef Name='PrimaryResMembers' />",
                     "<FieldRef Name='PriorityST' />",
                     "<FieldRef Name='ProjectCode' />",
                     "<FieldRef Name='ProjectFolder' />",
                     "<FieldRef Name='ProjectTask' />",
                     "<FieldRef Name='ProjectType' />",
                     "<FieldRef Name='ProposeClosureDate' />",
                     "<FieldRef Name='ProposedEndDate' />",
                     "<FieldRef Name='ProposedStartDate' />",
                     "<FieldRef Name='PSMembers' />",
                     "<FieldRef Name='PubSupportStatus' />",
                     "<FieldRef Name='QC' />",
                     "<FieldRef Name='QuickProject' />",
                     "<FieldRef Name='Reason' />",
                     "<FieldRef Name='ReasonType' />",
                     "<FieldRef Name='ReferenceCount' />",
                     "<FieldRef Name='RejectionDate' />",
                     "<FieldRef Name='RelatedProject' />",
                     "<FieldRef Name='Reviewers' />",
                     "<FieldRef Name='ServiceLevel' />",
                     "<FieldRef Name='SlideCount' />",
                     "<FieldRef Name='SOWBoxLink' />",
                     "<FieldRef Name='SOWCode' />",
                     "<FieldRef Name='SOWLink' />",
                     "<FieldRef Name='StandardBudgetHrs' />",
                     "<FieldRef Name='StandardService' />",
                     "<FieldRef Name='Status' />",
                     "<FieldRef Name='SubDeliverable' />",
                     "<FieldRef Name='SubDivision' />",
                     "<FieldRef Name='TA' />",
                     "<FieldRef Name='WBJID' />",
                     "<FieldRef Name='Writers' />",
                     "<FieldRef Name='Year' />",
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
            query.ViewXml = @"<View><ViewFields>" + viewFields + "</ViewFields>" + whereQuery + "</View>";
            //View field2 added to restriction in sharepoint lookup column =12

            viewFields2 = string.Concat(
                     "<FieldRef Name='Id' />",
                     "<FieldRef Name='AllDeliveryResources' />",
                     "<FieldRef Name='AllOperationresources' />",
                     "<FieldRef Name='Authors' />",
                     "<FieldRef Name='BD' />",
                     "<FieldRef Name='BDBackup' />",
                     "<FieldRef Name='CMBackup' />",
                     "<FieldRef Name='Editors' />",
                     "<FieldRef Name='GraphicsMembers' />"
                    );
            query2.ViewXml = @"<View><ViewFields>" + viewFields2 + "</ViewFields>" + whereQuery + "</View>";

            ListItemCollection listItemCollections = list.GetItems(query);
            context.Load(listItemCollections);
            context.ExecuteQuery();

            ListItemCollection listItemCollections2 = list.GetItems(query2);
            context.Load(listItemCollections2);
            context.ExecuteQuery();
            List<ListItemCollection> finalItems = new List<ListItemCollection>();
            if (listItemCollections.Count > 0)
            {
                finalItems.Add(listItemCollections);
            }
            if (listItemCollections2.Count > 0)
            {
                finalItems.Add(listItemCollections2);
            }
            return finalItems;
        }
        public Uri GetBlobUri()
        {
            return _azureBlob.GetUri(SharedAccessBlobPermissions.Read | SharedAccessBlobPermissions.List);
        }
    }
}
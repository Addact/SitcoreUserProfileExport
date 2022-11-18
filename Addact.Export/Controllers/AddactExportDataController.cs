using Addact.Export.Models;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Sitecore;
using Sitecore.Analytics.Model;
using Sitecore.DependencyInjection;
using Sitecore.Diagnostics;
using Sitecore.Globalization;
using Sitecore.Marketing.Definitions;
using Sitecore.XConnect;
using Sitecore.XConnect.Client;
using Sitecore.XConnect.Client.Configuration;
using Sitecore.XConnect.Collection.Model;
using Sitecore.XConnect.Collection.Model.Cache;
using Sitecore.XConnect.Operations;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls.Expressions;

namespace Addact.Export.Controllers
{
    public class AddactExportDataController : Controller
    {
     
        public FileContentResult ExportProfile(string startDate, string endDate)
        {
            try
            {

                List<ExperienceProfileDetail> export = new List<ExperienceProfileDetail>();
                string[] paramiters = { WebVisit.DefaultFacetKey, LocaleInfo.DefaultFacetKey, IpInfo.DefaultFacetKey };
                var interactionsobj = new RelatedInteractionsExpandOptions(paramiters);
                DateTime sdate = DateTime.Now;
                DateTime edate = DateTime.Now;

                if (!string.IsNullOrEmpty(startDate) && DateTime.TryParseExact(startDate, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out sdate))
                    interactionsobj.StartDateTime = sdate.AddDays(-1).ToUniversalTime();
                else
                    interactionsobj.StartDateTime = DateTime.UtcNow.AddDays(-30);
                if (!string.IsNullOrEmpty(endDate) && DateTime.TryParseExact(endDate, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out edate))
                    interactionsobj.EndDateTime = edate.AddDays(1).ToUniversalTime();
                else
                    interactionsobj.EndDateTime = DateTime.UtcNow;
                interactionsobj.Limit = int.MaxValue;
                ExportDataResult exportResult;
                using (Sitecore.XConnect.Client.XConnectClient client = SitecoreXConnectClientConfiguration.GetClient())
                {
                    IOrderedAsyncQueryable<Contact> orderedAsyncQueryable;
                    var Allcontacts = client.Contacts;
                    orderedAsyncQueryable = Allcontacts.Where<Contact>((Expression<Func<Contact, bool>>)(x => x.InteractionsCache().InteractionCaches.Any<InteractionCacheEntry>())).OrderByDescending<Contact, DateTime>((Expression<Func<Contact, DateTime>>)(c => c.EngagementMeasures().MostRecentInteractionStartDateTime));
                    List<Contact> contactList = new List<Contact>();
                    bool includeAnonymous = false;
                    var settingvalue = Sitecore.Configuration.Settings.GetSetting("IncludeAnonymous");
                    bool.TryParse(settingvalue, out includeAnonymous);
                    var Contactsid = (IAsyncQueryable<Contact>)orderedAsyncQueryable;
                    if (!includeAnonymous)
                        Contactsid = Contactsid.Where(c => c.Identifiers.Any(t => t.IdentifierType == Sitecore.XConnect.ContactIdentifierType.Known));
                    var Bookmark = (byte[])null;
                    Task<IAsyncEntityBatchEnumerator<Contact>> contactDetaillist = Contactsid.WithExpandOptions<Contact>((ExpandOptions)new ContactExpandOptions(PersonalInformation.DefaultFacetKey)
                    {
                        Interactions = interactionsobj
                    }.Expand<EmailAddressList>().Expand<AddressList>().Expand<PhoneNumberList>()).GetBatchEnumerator<Contact>(Bookmark, int.MaxValue);
                    int num = contactDetaillist.Result.MoveNext<IReadOnlyCollection<Contact>>().Result ? 1 : 0;
                    IReadOnlyCollection<Contact> current = contactDetaillist.Result.Current;
                    Bookmark = contactDetaillist.Result.GetBookmark();
                    var listofcontact = current.Where(x => x.Interactions != null && x.Interactions.Count() > 0).ToList();
                    exportResult = new ExportDataResult()
                    {
                        Content = GenerateFileContent(listofcontact),
                        FileName = GenerateFileName(interactionsobj.StartDateTime.Value, interactionsobj.EndDateTime.Value),
                        MediaType = "application/octet-stream"
                    };
                }

                FileContentResult fileresult;
                if (exportResult != null)
                {
                    fileresult = new FileContentResult(exportResult.Content, exportResult.MediaType);
                    fileresult.FileDownloadName = exportResult.FileName;
                }
                else
                { fileresult = new FileContentResult(new byte[0], "application/octet-stream") { FileDownloadName = "NoData.csv" }; }

                return fileresult;

            }
            catch (Exception ex)
            {

                Log.Error("ERROR IN EXPORT PROFILE GETDATA:", ex.Message);
                return new FileContentResult(new byte[0], "application/octet-stream") { FileDownloadName = "NoData.csv" };
            }
        }




        protected virtual string GenerateFileName(DateTime? startDate, DateTime? endDate)
        {
            string str = string.Empty;
            if (startDate.HasValue)
                str = str + "_from_" + DateUtil.ToIsoDate(startDate.Value);
            if (endDate.HasValue)
                str = str + "_until_" + DateUtil.ToIsoDate(endDate.Value);
            if (string.IsNullOrEmpty(str))
                str = "-" + DateUtil.IsoNow;
            return FormattableString.Invariant(FormattableStringFactory.Create("Profile-Data{0}.csv", (object)str));
        }
        public Sitecore.Security.Accounts.User GetUserbyUserName(string userName)
        {
            var UserName = string.Empty;
            var domain = Context.Domain;
            if (domain != null)
            {
                UserName = domain.GetFullName(userName);
            }
            if (Sitecore.Security.Accounts.User.Exists(UserName))
                return Sitecore.Security.Accounts.User.FromName(UserName, true);
            return null;
        }
        protected byte[] GenerateFileContent(List<Contact> contacts)
        {
            StringBuilder stringBuilder = new StringBuilder();
            string[] fieldColumnsList = {
                "FirstName",
                "MiddleName",
                "LastName",
                "Email",
                "Phone Number",
                "Address",
                "Company Name",
                "Event Type",
                "Page Url",
                "Page View Date",
                "Duration"
                };

            stringBuilder.AppendLine(string.Join(";", fieldColumnsList));
            var goalDefinitionManager = ServiceLocator.ServiceProvider.GetDefinitionManagerFactory().GetDefinitionManager<Sitecore.Marketing.Definitions.Goals.IGoalDefinition>();
            foreach (var contact in contacts)
            {
                try
                {
                    var contecatbehavior = contact.ContactBehaviorProfile();
                    if (contact != null)
                    {
                        string ContactEmail = "";
                        string ContactPhone = "";
                        string Address = "";
                        string Websitename = "";
                        if (contact != null)
                        {
                            EmailAddressList emailsFacetData = contact.GetFacet<EmailAddressList>();
                            if (emailsFacetData != null)
                            {
                                EmailAddress preferred = emailsFacetData.PreferredEmail;
                                ContactEmail = emailsFacetData.PreferredEmail.SmtpAddress;
                            }
                            var phoneNumber = contact.GetFacet<PhoneNumberList>();
                            if (phoneNumber != null)
                            {
                                PhoneNumber pn = phoneNumber.PreferredPhoneNumber;
                                ContactPhone = pn.Extension + " " + pn.Number;
                            }
                            var addresslist = contact.GetFacet<AddressList>();
                            if (addresslist != null)
                            {
                                Address add = addresslist.PreferredAddress;

                                Address = add.AddressLine1;
                                Address += !string.IsNullOrEmpty(Address) ? ", " + add.AddressLine2 : "";
                                Address += !string.IsNullOrEmpty(Address) ? ", " + add.AddressLine3 : "";
                                Address += !string.IsNullOrEmpty(Address) ? ", " + add.AddressLine4 : "";
                                Address += !string.IsNullOrEmpty(Address) ? ", " + add.City : "";
                                Address += !string.IsNullOrEmpty(Address) ? ", " + add.CountryCode : "";
                                Address += !string.IsNullOrEmpty(Address) ? "," + add.StateOrProvince : "";
                                Address += !string.IsNullOrEmpty(Address) ? ", " + add.PostalCode : "";
                            }
                        }

                        var webview = contact.Interactions;
                        foreach (var interaction in webview)
                        {
                            var ipinfo = interaction.IpInfo();
                            var intWebvisit = interaction.WebVisit();
                            if (intWebvisit != null)
                                Websitename = intWebvisit.SiteName;
                            if (interaction != null && interaction.Events != null && interaction.Events.Count() > 0)
                            {
                                foreach (var pevent in interaction.Events)
                                {
                                    string[] strArray = new string[18];
                                    var persion = contact.Personal();
                                    if (persion != null)
                                    {
                                        strArray[0] = contact.Personal() != null && !string.IsNullOrEmpty(persion.FirstName) ? persion.FirstName : "";
                                        strArray[1] = contact.Personal() != null && !string.IsNullOrEmpty(persion.MiddleName) ? persion.MiddleName : "";
                                        strArray[2] = contact.Personal() != null && !string.IsNullOrEmpty(persion.LastName) ? persion.LastName : "";
                                        strArray[3] = !string.IsNullOrEmpty(ContactEmail) ? ContactEmail : "";
                                        strArray[4] = !string.IsNullOrEmpty(ContactPhone) ? ContactPhone : "";
                                        strArray[5] = !string.IsNullOrEmpty(Address.ToString()) ? Address.ToString() : "";
                                        strArray[6] = contact.Personal() != null && !string.IsNullOrEmpty(persion.Title) ? persion.Title : "";
                                        if (contact.Personal() != null && !string.IsNullOrEmpty(ContactEmail) && string.IsNullOrEmpty(persion.Title))
                                        {
                                            var fullName = "extranet" + "\\" + ContactEmail;
                                            Sitecore.Security.Accounts.User user = GetUserbyUserName(fullName);
                                            if (user != null && user.Profile != null)
                                            {
                                                var companyname = user.Profile.GetCustomProperty("CRMCompanyName");
                                                if (!string.IsNullOrEmpty(companyname))
                                                {
                                                    strArray[6] = companyname;
                                                }
                                                
                                            }
                                            else
                                            {
                                                try
                                                {
                                                    using (OrganizationService service = new OrganizationService("BGH"))
                                                    {

                                                        Microsoft.Xrm.Sdk.Query.QueryExpression query = new Microsoft.Xrm.Sdk.Query.QueryExpression
                                                        {
                                                            EntityName = "contact",
                                                            ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet("parentcustomerid")

                                                        };
                                                        query.Criteria = new Microsoft.Xrm.Sdk.Query.FilterExpression();
                                                        query.Criteria.AddCondition("emailaddress1", ConditionOperator.Equal, ContactEmail);
                                                        var ContactRecord = service.RetrieveMultiple(query);

                                                        if (ContactRecord != null && ContactRecord.Entities != null && ContactRecord.Entities.Count() > 0)
                                                        {
                                                            if (ContactRecord.Entities.FirstOrDefault().Contains("parentcustomerid") && ContactRecord.Entities.FirstOrDefault()["parentcustomerid"] != null)
                                                            {
                                                                EntityReference parentcustomerid = (EntityReference)ContactRecord.Entities.FirstOrDefault()["parentcustomerid"];
                                                                if (parentcustomerid.Name != null)
                                                                {
                                                                    strArray[6] = parentcustomerid.Name;
                                                                }
                                                            }

                                                        }
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    var error = ex.Message;
                                                }
                                            }
                                        }
                                        if (pevent.GetType().Name == "PageViewEvent")
                                        {
                                            var pageview = (PageViewEvent)pevent;
                                            if (!string.IsNullOrEmpty(pageview.Url) && !pageview.Url.Contains("AddactExportData"))
                                            {
                                                if (pageview.Url.Contains("ClientSaveResourceTracker.ashx"))
                                                {
                                                    strArray[7] = "Saved Resource";
                                                    string itemId = string.Empty;
                                                    MatchCollection guids = Regex.Matches(pageview.Url, @"(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}"); //Match all substrings in findGuid
                                                    for (int i = 0; i < guids.Count; i++)
                                                    {
                                                        itemId = guids[i].Value; //Set Match to the value from the match                                                    
                                                    }

                                                    if (!string.IsNullOrEmpty(itemId))
                                                    {
                                                        var item = Sitecore.Context.Database.GetItem(new Sitecore.Data.ID(itemId));
                                                        strArray[8] = item != null && !string.IsNullOrEmpty(item.Name) ? item.Name : "";
                                                    }
                                                    if (pageview.Timestamp.Date != null)
                                                    {
                                                        strArray[9] = pageview.Timestamp.Date.ToString("MM-dd-yyyy");
                                                    }
                                                    else
                                                    {
                                                        strArray[9] = "";
                                                    }
                                                    strArray[10] = pageview.Duration.ToString();
                                                }
                                                else if (pageview.Url.Contains("ClientDownloadPDFTracker.ashx"))
                                                {
                                                    if (pageview.Url.Contains("|"))
                                                    {
                                                        strArray[7] = "Download PDF";
                                                        var name = System.Web.HttpUtility.UrlDecode(pageview.Url.Split('|')[1]);
                                                        name = name.Replace(name.Substring(name.IndexOf("EventName") - 3), "");

                                                        name = name.Replace("\"", "");
                                                        strArray[8] = !string.IsNullOrEmpty(name) ? name : pageview.Url;

                                                        if (pageview.Timestamp.Date != null)
                                                        {
                                                            strArray[9] = pageview.Timestamp.Date.ToString("MM-dd-yyyy");
                                                        }
                                                        else
                                                        {
                                                            strArray[9] = "";
                                                        }
                                                        strArray[10] = pageview.Duration.ToString();
                                                    }
                                                }
                                                else if (pageview.Url.Contains("/PageNotFound") || pageview.Url.Contains("/en/PageNotFound"))
                                                {
                                                    continue;
                                                }
                                                else if (pageview.Url.Contains("/fieldtracking/register") || pageview.Url.Contains("/formbuilder?fxb.FormItemId"))
                                                {
                                                    continue;
                                                }
                                                else
                                                {
                                                    strArray[7] = "Page View";
                                                    strArray[8] = pageview.Url;
                                                    //strArray[9] = pageview.Timestamp.ToShortDateString();
                                                    if (pageview.Timestamp.Date != null)
                                                    {
                                                        strArray[9] = pageview.Timestamp.Date.ToString("MM-dd-yyyy");
                                                    }
                                                    else
                                                    {
                                                        strArray[9] = "";
                                                    }
                                                    strArray[10] = pageview.Duration.ToString();
                                                }
                                            }
                                        }
                                        if (pevent.GetType().Name == "Goal")
                                        {
                                            var pagegoal = (Sitecore.XConnect.Goal)pevent;
                                            Guid goalId = pagegoal.DefinitionId;
                                            Sitecore.Marketing.Definitions.Goals.IGoalDefinition goalOne = goalDefinitionManager.Get(goalId, new CultureInfo("en"), true);
                                            strArray[7] = "Goal";
                                            strArray[8] = goalOne != null ? goalOne.Name : "";
                                            //strArray[9] = pagegoal.Timestamp.ToShortDateString();
                                            if (pagegoal.Timestamp.Date != null)
                                            {
                                                strArray[9] = pagegoal.Timestamp.Date.ToString("MM-dd-yyyy");
                                            }
                                            else
                                            {
                                                strArray[9] = "";
                                            }
                                            strArray[10] = pagegoal.Duration.ToString();
                                        }
                                    }
                                    if (!string.IsNullOrEmpty(strArray[7]))
                                    {
                                        stringBuilder.AppendLine(string.Join(";", strArray));
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("ERROR IN EXPORT PROFILE CONTACT:", ex.Message);
                }
            }
            byte[] buffer = Encoding.ASCII.GetBytes(stringBuilder.ToString());

            return buffer;
        }

    }
}
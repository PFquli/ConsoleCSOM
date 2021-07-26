using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Collections.Generic;
using Microsoft.SharePoint.Client.UserProfiles;

using SP = Microsoft.SharePoint.Client;

using System.Text;

namespace ConsoleCSOM
{
    internal class SharepointInfo
    {
        public string SiteUrl { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
    }

    internal class Program
    {
        private static string ListNameConst = "CSOM Test List";

        private static string ContentTypeIdConst = "0x0100" + new Guid("33847D27-C289-47F3-AEE7-AFED960DF770").ToString("N"); // parent is Item type

        private static string ContentTypeNameConst = "ContentTypeNameConst";

        private static string DocumentLibNameConst = "Document Test";

        private static string UserEmailConst = "admin@mystartpoint.onmicrosoft.com";

        private static async Task Main(string[] args)
        {
            try
            {
                using (var clientContextHelper = new ClientContextHelper())
                {
                    ClientContext ctx = GetContext(clientContextHelper);
                    ctx.Load(ctx.Web);
                    await ctx.ExecuteQueryAsync();

                    Console.WriteLine($"Site {ctx.Web.Title}");

                    //await CreateCSOMTestList(ctx);
                    //await SimpleCamlQueryAsync(ctx);
                    //await CsomTermSetAsync(ctx);
                    //await CreateTermSetInDevTenant(ctx);
                    //await CreateNewTerms(ctx);
                    //await CreateSiteFields(ctx);
                    //await CreateContentType(ctx);
                    //await AddFieldToContentType(ctx);
                    //await AddContentTypeToList(ctx);
                    //await SetDefaultContentTypeForList(ctx);
                    //await CreateNewListItems(ctx);
                    //await UpdateDefaultValueForAboutField(ctx);
                    //await AddNewListItemsAfterUpdatingAboutDefault(ctx);
                    //await UpdateDefaultValueForCityField(ctx);
                    //await AddNewListItemsAfterUpdatingCityDefault(ctx);
                    //await QueryListItemNotAboutDefault(ctx);
                    //await CreateListViewWithFilters(ctx);
                    //await UpdateMultipleAboutDefaultField(ctx);
                    //await CreateAuthorFieldInList(ctx);
                    //await MigrateAllListItemsAndSetAdmin(ctx);
                    //await CreateMultiTaxonomyField(ctx);
                    //await AddFieldToContentTypeAndMakeAvailableInList(ctx);
                    //await AddListItemsWithTaxonomyMultiValue(ctx);
                    //await CreateDocumentLibrary(ctx);
                    //await AddContentTypeToDocumentLibrary(ctx);
                    //await CreateFolderAndSubFolder(ctx);
                    //await CreateListItemsInSubFolder(ctx);
                    //await StockholmItemsInSubFolder(ctx);
                    //await UploadDocumentToDocumentLibrary(ctx);
                    //await CreateFolderViewAndSetDefaultView(ctx);
                    //var user = LoadUserFromEmailOrUserName(ctx, "admin@mystartpoint.onmicrosoft.com");
                    //CreateTestLevelPermissionLevel(ctx);
                    //DeleteGroupFromSite(ctx);
                    //CreateTestGroupWithTestLevelAndAddUser(ctx);
                    //GetInheritedGroupFromSubsite(ctx);

                    //UpdateTestBoolPropertyInCurrentUserProfile(ctx, "true");
                    //UpdateTestTextPropertyInCurrentUserProfile(ctx, "Updated text");
                    //UpdateTestDatePropertyInCurrentUserProfile(ctx, "2021-7-26");
                    //UpdateTestIntegerPropertyInCurrentUserProfile(ctx, "123");
                    //UpdateTestEmailPropertyInCurrentUserProfile(ctx, "testing@testmail.com");
                    //UpdateTestPersonPropertyInCurrentUserProfile(ctx, "User Unknown");
                    //UpdateTestSingleTaxonomyPropertyInCurrentUserProfile(ctx, "Rome");
                    //UpdateTestMultipleTaxonomyPropertyInCurrentUserProfile(ctx, new List<string>() { "Ho Chi Minh", "Stockholm" });
                    UserProfileConsoleApp(ctx);
                }

                Console.WriteLine($"Press Any Key To Stop!");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
            }
        }

        #region CSOM & CAML Training

        #region 1/1

        private static async Task CreateCSOMTestList(ClientContext ctx)
        {
            Web web = ctx.Web;
            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.TemplateType = (int)ListTemplateType.GenericList;
            creationInfo.Description = web.Description;
            creationInfo.Title = ListNameConst;
            try
            {
                SP.List newList = web.Lists.Add(creationInfo);

                await ctx.ExecuteQueryAsync();
            }
            catch (Exception ex)
            { }
        }

        #endregion 1/1

        #region 1/2

        private static async Task CreateTermSetInDevTenant(ClientContext ctx)
        {
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            // Get the term store by name
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            const string termGroupName = "Training";
            try
            {
                TermGroup termGroup = termStore.CreateGroup(termGroupName, new Guid("88145EE9-7A0C-4445-A40B-51E6A97C8DB5"));
                const string termSetName = "city-Quoc";
                const int lcid = 1033;
                TermSet termSet = termGroup.CreateTermSet(termSetName, new Guid("98CBDBED-53AE-42E9-AD08-42DED45922D0"), lcid);
                var terms = termSet.GetAllTerms();

                ctx.Load(terms);
                await ctx.ExecuteQueryAsync();
            }
            catch (Exception e)
            {
            }
        }

        #endregion 1/2

        #region 1/3

        private async static Task CreateNewTerms(ClientContext ctx)
        {
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            // Get the term store by name
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get the term group by Name
            TermGroup termGroup = termStore.Groups.GetByName("Training");
            // Get the term set by Name
            TermSet termSet = termGroup.TermSets.GetByName("city-Quoc");
            // Create new terms
            int lcid = 1033;
            try
            {
                Term term0 = termSet.CreateTerm("Ho Chi Minh", lcid, new Guid("2FCA8C5F-0DEF-442F-8386-FEB21568109B"));
                Term term1 = termSet.CreateTerm("Stockholm", lcid, new Guid("65F5B6AF-3FD0-4790-966E-2F34EF5C5504"));
                ctx.Load(term0);
                ctx.Load(term1);
                await ctx.ExecuteQueryAsync();
            }
            catch (Exception e)
            {
            }
        }

        #endregion 1/3

        #region 1/4

        private async static Task CreateSiteFields(ClientContext ctx)
        {
            try
            {
                Web web = ctx.Web;
                FieldCollection fields = web.Fields;
                fields.AddFieldAsXml(
                    @"<Field                Type='Text'
                                            Name='about'
                                            Required='FALSE'
                                            DisplayName='about'
                                            Description=''
                                            Group='Custom Columns'/>", true, AddFieldOptions.DefaultValue);
                var field = fields.AddFieldAsXml(
                    @"<Field                Type='TaxonomyFieldType'
                                            Name='cityCSOM'
                                            Required='FALSE'
                                            DisplayName='cityCSOM'
                                            Description=''
                                            Hidden='FALSE'
                                            Group='Custom Columns'/>", true, AddFieldOptions.DefaultValue);
                TaxonomyField taxonomyField = ctx.CastTo<TaxonomyField>(field);
                TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);
                TermStore termStore = session.GetDefaultSiteCollectionTermStore();

                // Get the term group by Name
                TermGroup termGroup = termStore.Groups.GetByName("Training");
                // Get the term set by Name
                TermSet termSet = termGroup.TermSets.GetByName("city-Quoc");

                ctx.Load(termSet, tst => tst.Id);
                ctx.Load(termStore, ts => ts.Id);
                ctx.ExecuteQuery();

                var termStoreId = termStore.Id;
                var termSetId = termSet.Id;
                taxonomyField.SspId = termStoreId;
                taxonomyField.TermSetId = termSetId;
                taxonomyField.TargetTemplate = String.Empty;
                taxonomyField.AnchorId = Guid.Empty;
                taxonomyField.Update();
                await ctx.ExecuteQueryAsync();
            }
            catch (Exception e)
            {
            }
        }

        #endregion 1/4

        #region 1/5

        private async static Task CreateContentType(ClientContext ctx)
        {
            ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
            ctx.Load(contentTypes);
            ctx.ExecuteQuery();

            foreach (var item in contentTypes)
            {
                if (item.StringId == ContentTypeIdConst)
                    return;
            }

            // Create a Content Type Information object.
            ContentTypeCreationInformation newCt = new ContentTypeCreationInformation();
            // Set the name for the content type.
            newCt.Name = "ContentTypeNameConst";
            // Inherit from oob document - 0x0101 and assign.
            newCt.Id = ContentTypeIdConst;
            // Set content type to be available from specific group.
            newCt.Group = "Custom Content Types";
            // Create the content type.
            ContentType testContentType = contentTypes.Add(newCt);
            await ctx.ExecuteQueryAsync();
        }

        private async static Task AddFieldToContentType(ClientContext ctx)
        {
            ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
            var query = (from contentType in contentTypes where contentType.Name == "ContentTypeNameConst" select contentType);
            var results = ctx.LoadQuery(query);
            ctx.ExecuteQuery();
            ContentType testContentType = (ContentType)results.FirstOrDefault();

            // Add site fields about and city to content type
            Field targetField0 = ctx.Web.AvailableFields.GetByInternalNameOrTitle("about");
            Field targetField1 = ctx.Web.AvailableFields.GetByInternalNameOrTitle("cityCSOM");

            ctx.Load(targetField0);
            ctx.Load(targetField1);
            ctx.ExecuteQuery();

            // Workaround: check for duplicate field, delete it and try adding again
            bool success = false;
            while (!success)
            {
                try
                {
                    FieldLinkCreationInformation fldLink0 = new FieldLinkCreationInformation();
                    FieldLinkCreationInformation fldLink1 = new FieldLinkCreationInformation();
                    fldLink0.Field = targetField0;
                    fldLink1.Field = targetField1;

                    fldLink0.Field.Required = false;
                    fldLink1.Field.Required = false;

                    testContentType.FieldLinks.Add(fldLink0);
                    testContentType.FieldLinks.Add(fldLink1);
                    testContentType.Update(false);

                    await ctx.ExecuteQueryAsync();
                    success = true;
                }
                catch (ServerException ex)
                {
                    if (ex.Message.Contains("A duplicate field name"))
                    {
                        var splitMessage = ex.Message.Split(new[] { '\"' }, StringSplitOptions.RemoveEmptyEntries);
                        var duplicateName = splitMessage[1];
                        DeleteSiteColumn(duplicateName, ctx);
                    }
                    else
                    {
                        throw;
                    }
                }
            }
        }

        private static void DeleteSiteColumn(string name, ClientContext ctx)
        {
            var siteColumn = ctx.Web.Fields.GetByInternalNameOrTitle(name);
            if (siteColumn == null) return;
            siteColumn.DeleteObject();
            ctx.ExecuteQuery();
        }

        private async static Task AddContentTypeToList(ClientContext ctx)
        {
            ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
            var query = (from contentType in contentTypes where contentType.Name == "ContentTypeNameConst" select contentType);
            var results = ctx.LoadQuery(query);
            ctx.ExecuteQuery();
            ContentType testContentType = (ContentType)results.FirstOrDefault();
            // Get list
            List testList = ctx.Web.Lists.GetByTitle(ListNameConst);
            // Add content type to list and update
            testList.ContentTypes.AddExistingContentType(testContentType);
            testList.Update();
            ctx.Web.Update();
            await ctx.ExecuteQueryAsync();
        }

        #endregion 1/5

        #region 1/6

        private static async Task SetDefaultContentTypeForList(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(ListNameConst);
            ContentTypeCollection ctCol = list.ContentTypes;
            ctx.Load(ctCol, coll => coll.Include(
                ct => ct.Name,
                        ct => ct.Id));
            ctx.ExecuteQuery();
            List<ContentTypeId> reverseOrder = (from ct in ctCol where ct.Name.Equals("ContentTypeNameConst", StringComparison.OrdinalIgnoreCase) select ct.Id).ToList();
            list.RootFolder.UniqueContentTypeOrder = reverseOrder;
            list.RootFolder.Update();
            list.Update();
            await ctx.ExecuteQueryAsync();
        }

        #endregion 1/6

        #region 1/7

        private async static Task CreateNewListItems(ClientContext ctx)
        {
            SP.List oList = ctx.Web.Lists.GetByTitle(ListNameConst);

            Field field = oList.Fields.GetByTitle("cityCSOM");

            ctx.Load(field);

            ctx.ExecuteQuery();

            TaxonomyField taxField = ctx.CastTo<TaxonomyField>(field);

            ctx.Load(oList);
            for (var i = 0; i < 5; i++)
            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = oList.AddItem(itemCreateInfo);
                oListItem["Title"] = "Title" + i;
                oListItem["about"] = "about" + i;
                oListItem["ContentTypeId"] = ContentTypeIdConst;

                taxField.SetFieldValueByValue(oListItem, new TaxonomyFieldValue()
                {
                    WssId = -1, // alway let it -1
                    Label = "Ho Chi Minh",
                    TermGuid = "2fca8c5f-0def-442f-8386-feb21568109b"
                });
                oListItem.Update();
            }
            await ctx.ExecuteQueryAsync();
        }

        #endregion 1/7

        #region 1/8

        private async static Task UpdateDefaultValueForAboutField(ClientContext ctx)
        {
            List targetList = ctx.Web.Lists.GetByTitle(ListNameConst);

            Field oField = ctx.Web.Fields.GetByTitle("about");

            // Need to load field to get default value of it
            //ctx.Load(oField);

            oField.DefaultValue = "about default";
            oField.UpdateAndPushChanges(true);
            await ctx.ExecuteQueryAsync();
        }

        private async static Task AddNewListItemsAfterUpdatingAboutDefault(ClientContext ctx)
        {
            List oList = ctx.Web.Lists.GetByTitle(ListNameConst);

            Field field = oList.Fields.GetByTitle("cityCSOM");

            ctx.Load(field);

            ctx.ExecuteQuery();

            TaxonomyField taxField = ctx.CastTo<TaxonomyField>(field);

            ctx.Load(oList);
            for (var i = 10; i < 12; i++)
            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = oList.AddItem(itemCreateInfo);
                oListItem["Title"] = "Title" + i;
                oListItem["ContentTypeId"] = ContentTypeIdConst;

                taxField.SetFieldValueByValue(oListItem, new TaxonomyFieldValue()
                {
                    WssId = -1, // alway let it -1
                    Label = "Stockholm",
                    TermGuid = "65f5b6af-3fd0-4790-966e-2f34ef5c5504"
                });
                oListItem.Update();
            }
            await ctx.ExecuteQueryAsync();
        }

        #endregion 1/8

        #region 1/9

        private async static Task UpdateDefaultValueForCityField(ClientContext ctx)
        {
            Field field = ctx.Web.Fields.GetByTitle("cityCSOM");

            ctx.Load(field);

            ctx.ExecuteQuery();

            TaxonomyField taxField = ctx.CastTo<TaxonomyField>(field);

            TaxonomyFieldValue defaultValue = new TaxonomyFieldValue();
            defaultValue.WssId = -1;
            defaultValue.Label = "Ho Chi Minh";
            // GUID should be stored lowercase, otherwise it will not work in Office 2010
            defaultValue.TermGuid = "2fca8c5f-0def-442f-8386-feb21568109b";

            // Get the Validated String for the taxonomy value
            var validatedValue = taxField.GetValidatedString(defaultValue);
            ctx.ExecuteQuery();

            // Set the selected default value for the site column
            taxField.DefaultValue = validatedValue.Value;
            taxField.UserCreated = false;
            taxField.UpdateAndPushChanges(true);
            await ctx.ExecuteQueryAsync();
        }

        // Todo: fix the "default value of city not showing" bug
        private async static Task AddNewListItemsAfterUpdatingCityDefault(ClientContext ctx)
        {
            List oList = ctx.Web.Lists.GetByTitle(ListNameConst);

            Field field = oList.Fields.GetByTitle("cityCSOM");

            ctx.Load(field);

            ctx.ExecuteQuery();

            TaxonomyField taxField = ctx.CastTo<TaxonomyField>(field);

            ctx.Load(oList);
            for (var i = 20; i < 22; i++)
            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = oList.AddItem(itemCreateInfo);
                oListItem["Title"] = "Title" + i;
                oListItem["about"] = "about" + i;
                oListItem["ContentTypeId"] = ContentTypeIdConst;
                oListItem.Update();
            }
            await ctx.ExecuteQueryAsync();
        }

        #endregion 1/9

        #region 2/1

        private static async Task QueryListItemNotAboutDefault(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(ListNameConst);
            var items = list.GetItems(new CamlQuery()
            {
                ViewXml = @"<View>
                                <Query>
                                    <OrderBy><FieldRef Name='ID' Ascending='False'/></OrderBy>
                                    <Where>
                                      <Neq>
                                        <FieldRef Name='about'></FieldRef>
                                        <Value Type='Text'>about default</Value>
                                      </Neq>
                                    </Where>
                                </Query>
                                <RowLimit>20</RowLimit>
                            </View>"
                //FolderServerRelativeUrl = "/sites/PrecioFishbone/CSOM Test List/"
            });
            ctx.Load(items);
            await ctx.ExecuteQueryAsync();
        }

        #endregion 2/1

        #region 2/2

        private static async Task CreateListViewWithFilters(ClientContext ctx)
        {
            // clientcontext.Web.Lists.GetById - This option also can be used to get the list using List GUID
            // This value is NOT List internal name
            List targetList = ctx.Web.Lists.GetByTitle(ListNameConst);

            ViewCollection viewCollection = targetList.Views;

            ViewCreationInformation viewCreationInformation = new ViewCreationInformation();
            viewCreationInformation.Title = "HCM Newest";

            // Specify type of the view. Below are the options

            // 1. none - The type of the list view is not specified

            // 2. html - Sspecifies an HTML list view type

            // 3. grid - Specifies a datasheet list view type

            // 4. calendar- Specifies a calendar list view type

            // 5. recurrence - Specifies a list view type that displays recurring events

            // 6. chart - Specifies a chart list view type

            // 7. gantt - Specifies a Gantt chart list view type

            viewCreationInformation.ViewTypeKind = ViewType.Html;

            // You can optionally specify row limit for the view
            viewCreationInformation.RowLimit = 10;

            // You can optionally specify a query as mentioned below.
            // Create one CAML query to filter list view and mention that query below
            viewCreationInformation.Query = "<Where><Eq><FieldRef Name = 'cityCSOM' /><Value Type = 'TaxonomyFieldType'>Ho Chi Minh</Value></Eq></Where><OrderBy><FieldRef Name='Modified' Ascending='False'/></OrderBy>";

            // Add all the fields over here with comma separated value as mentioned below
            // You can mention display name or internal name of the column
            string CommaSeparateColumnNames = "ID,Title,cityCSOM,about";
            viewCreationInformation.ViewFields = CommaSeparateColumnNames.Split(',');

            View listView = viewCollection.Add(viewCreationInformation);
            ctx.ExecuteQuery();

            // Code to update the display name for the view.
            listView.Title = "HCM Newest";

            // You can optionally specify Aggregation: Field references for totals columns or calculated columns
            //listView.Aggregations = "<FieldRef Name='Title' Type='COUNT'/>";

            listView.Update();
            await ctx.ExecuteQueryAsync();
        }

        #endregion 2/2

        #region 2/3

        private static async Task UpdateMultipleAboutDefaultField(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(ListNameConst);
            var items = list.GetItems(new CamlQuery()
            {
                ViewXml = @"<View>
                                <Query>
                                    <OrderBy><FieldRef Name='ID' Ascending='False'/></OrderBy>
                                    <Where>
                                      <Eq>
                                        <FieldRef Name='about'></FieldRef>
                                        <Value Type='Text'>about default</Value>
                                      </Eq>
                                    </Where>
                                </Query>
                            </View>"
                //FolderServerRelativeUrl = "/sites/PrecioFishbone/CSOM Test List/"
            });
            ctx.Load(items);
            ctx.ExecuteQuery();
            int updateTracker = 0;
            foreach (ListItem item in items)
            {
                item["about"] = "Update script";
                item.Update();
                updateTracker++;
                if (updateTracker > 1)
                {
                    await ctx.ExecuteQueryAsync();
                    updateTracker = 0;
                }
            }
            if (updateTracker > 0)
                await ctx.ExecuteQueryAsync();
        }

        #endregion 2/3

        #region 2/4

        private static async Task CreateAuthorFieldInList(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(ListNameConst);
            FieldCollection fields = list.Fields;
            string field = @"<Field Name='CSOMTestAuthor' DisplayName='CSOM Test Author' Type='User' Group='Custom Columns' />";
            fields.AddFieldAsXml(field, true, AddFieldOptions.DefaultValue);

            await ctx.ExecuteQueryAsync();
        }

        private static async Task MigrateAllListItemsAndSetAdmin(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(ListNameConst);
            CamlQuery query = CamlQuery.CreateAllItemsQuery();
            ListItemCollection listItems = list.GetItems(query);
            ctx.Load(listItems);
            //List<UserEntity> admins = ctx.Site.RootWeb.GetAdministrators();
            //UserEntity admin = admins[0];
            var currentUser = ctx.Web.EnsureUser(UserEmailConst);
            ctx.Load(currentUser, cu => cu.Id);
            ctx.ExecuteQuery();
            int userId = currentUser.Id;
            foreach (ListItem listItem in listItems)
            {
                FieldUserValue uservalue = new FieldUserValue();
                uservalue.LookupId = userId;
                listItem["CSOM_x0020_Test_x0020_Author"] = uservalue;
                listItem.Update();
            }
            await ctx.ExecuteQueryAsync();
        }

        #endregion 2/4

        #region 3/1

        private static async Task CreateMultiTaxonomyField(ClientContext ctx)
        {
            FieldCollection fields = ctx.Web.Fields;
            var field = fields.AddFieldAsXml(
                                        @"<Field
                                            Type='TaxonomyFieldTypeMulti'
                                            Name='cities'
                                            Required='FALSE'
                                            DisplayName='cities'
                                            Description=''
                                            Hidden='FALSE'
                                            EnforceUniqueValues='FALSE'
                                            Group ='Custom Columns'/>", true, AddFieldOptions.DefaultValue);
            TaxonomyField taxonomyField = ctx.CastTo<TaxonomyField>(field);
            TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = session.GetDefaultSiteCollectionTermStore();

            // Get the term group by Name
            TermGroup termGroup = termStore.Groups.GetByName("Training");
            // Get the term set by Name
            TermSet termSet = termGroup.TermSets.GetByName("city-Quoc");

            ctx.Load(termSet, tst => tst.Id);
            ctx.Load(termStore, ts => ts.Id);
            ctx.ExecuteQuery();

            var termStoreId = termStore.Id;
            var termSetId = termSet.Id;
            taxonomyField.SspId = termStoreId;
            taxonomyField.TermSetId = termSetId;
            taxonomyField.TargetTemplate = String.Empty;
            taxonomyField.AnchorId = Guid.Empty;
            taxonomyField.AllowMultipleValues = true;
            taxonomyField.Update();
            await ctx.ExecuteQueryAsync();
        }

        #endregion 3/1

        #region 3/2

        private static async Task AddFieldToContentTypeAndMakeAvailableInList(ClientContext ctx)
        {
            ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
            var query = (from contentType in contentTypes where contentType.Name == "ContentTypeNameConst" select contentType);
            var results = ctx.LoadQuery(query);
            ctx.ExecuteQuery();
            ContentType testContentType = (ContentType)results.FirstOrDefault();
            // Add site fields about and city to content type
            Field field = ctx.Web.AvailableFields.GetByInternalNameOrTitle("cities");

            ctx.Load(field);
            ctx.ExecuteQuery();

            // Workaround: check for duplicate field, delete it and try adding again
            bool success = false;
            while (!success)
            {
                try
                {
                    FieldLinkCreationInformation fldLink = new FieldLinkCreationInformation();
                    fldLink.Field = field;
                    fldLink.Field.Required = false;

                    testContentType.FieldLinks.Add(fldLink);
                    testContentType.Update(true);

                    await ctx.ExecuteQueryAsync();
                    success = true;
                }
                catch (ServerException ex)
                {
                    if (ex.Message.Contains("A duplicate field name"))
                    {
                        var splitMessage = ex.Message.Split(new[] { '\"' }, StringSplitOptions.RemoveEmptyEntries);
                        var duplicateName = splitMessage[1];
                        DeleteSiteColumn(duplicateName, ctx);
                    }
                    else
                    {
                        throw;
                    }
                }
            }
        }

        #endregion 3/2

        #region 3/3

        private static async Task AddListItemsWithTaxonomyMultiValue(ClientContext ctx)
        {
            List oList = ctx.Web.Lists.GetByTitle(ListNameConst);

            Field field = oList.Fields.GetByTitle("cities");

            ctx.Load(field);

            ctx.ExecuteQuery();

            TaxonomyField taxField = ctx.CastTo<TaxonomyField>(field);

            ctx.Load(oList);
            for (var i = 30; i < 33; i++)
            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = oList.AddItem(itemCreateInfo);
                oListItem["Title"] = "Title" + i;
                oListItem["ContentTypeId"] = ContentTypeIdConst;
                TaxonomyFieldValueCollection values = new TaxonomyFieldValueCollection(ctx, string.Empty, field);
                // PopulateFromLabelGuidPairs string: label|Guid. All WssId's will be set to -1
                values.PopulateFromLabelGuidPairs(@"Stockholm|65f5b6af-3fd0-4790-966e-2f34ef5c5504");
                values.PopulateFromLabelGuidPairs(@"Ho Chi Minh|2fca8c5f-0def-442f-8386-feb21568109b");
                taxField.SetFieldValueByValueCollection(oListItem, values);
                oListItem.Update();
            }
            await ctx.ExecuteQueryAsync();
        }

        #endregion 3/3

        #region 3/4

        private static async Task CreateDocumentLibrary(ClientContext ctx)
        {
            Web web = ctx.Web;
            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.TemplateType = (int)ListTemplateType.DocumentLibrary;
            creationInfo.Description = web.Description;
            creationInfo.Title = DocumentLibNameConst;
            try
            {
                SP.List newList = web.Lists.Add(creationInfo);

                await ctx.ExecuteQueryAsync();
            }
            catch (Exception ex)
            { }
        }

        private static async Task AddContentTypeToDocumentLibrary(ClientContext ctx)
        {
            ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
            var query = (from contentType in contentTypes where contentType.Name == "ContentTypeNameConst" select contentType);
            var results = ctx.LoadQuery(query);
            ctx.ExecuteQuery();
            ContentType testContentType = (ContentType)results.FirstOrDefault();
            // Get list
            List testDoc = ctx.Web.Lists.GetByTitle(DocumentLibNameConst);
            // Add content type to list and update
            testDoc.ContentTypes.AddExistingContentType(testContentType);
            testDoc.Update();
            ctx.Web.Update();
            await ctx.ExecuteQueryAsync();
        }

        #endregion 3/4

        #region 3/5

        private static async Task CreateFolderAndSubFolder(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(DocumentLibNameConst);
            //Enable Folder creation for the list
            list.EnableFolderCreation = true;
            FolderCollection folders = list.RootFolder.Folders;

            ctx.Load(folders);
            list.Update();
            ctx.ExecuteQuery();

            Folder newFolder = folders.Add("Folder 1");

            newFolder.Update();

            ctx.ExecuteQuery();

            newFolder.Folders.Add("Folder 2");

            newFolder.Update();

            await ctx.ExecuteQueryAsync();
        }

        private static async Task CreateListItemsInSubFolder(ClientContext ctx)
        {
            List oList = ctx.Web.Lists.GetByTitle(DocumentLibNameConst);
            ctx.Load(oList.RootFolder, f => f.ServerRelativeUrl);
            ctx.ExecuteQuery();
            string targetFolderPath = "Folder 1/Folder 2";
            string targetUrl = $"{oList.RootFolder.ServerRelativeUrl}/{targetFolderPath}";

            for (var i = 40; i < 43; i++)
            {
                var fileCreationInfo = new FileCreationInformation
                {
                    Content = Encoding.ASCII.GetBytes("test"),
                    Url = $"{targetUrl}/test{i}.txt"
                };
                File file = oList.RootFolder.Files.Add(fileCreationInfo);
                ctx.ExecuteQuery();
                ListItem oListItem = file.ListItemAllFields;
                oListItem["Title"] = "Title" + i;
                oListItem["ContentTypeId"] = ContentTypeIdConst;
                oListItem["about"] = "Folder test";
                oListItem.Update();
            }
            Field field = oList.Fields.GetByTitle("cities");

            ctx.Load(field);

            ctx.ExecuteQuery();

            TaxonomyField taxField = ctx.CastTo<TaxonomyField>(field);

            for (var i = 43; i < 45; i++)
            {
                var fileCreationInfo = new FileCreationInformation
                {
                    Content = Encoding.ASCII.GetBytes("test"),
                    Url = $"{targetUrl}/test{i}.txt"
                };
                File file = oList.RootFolder.Files.Add(fileCreationInfo);
                ctx.ExecuteQuery();
                ListItem oListItem = file.ListItemAllFields;
                oListItem["Title"] = "Title" + i;
                oListItem["ContentTypeId"] = ContentTypeIdConst;
                taxField.SetFieldValueByValue(oListItem, new TaxonomyFieldValue()
                {
                    WssId = -1, // alway let it -1
                    Label = "Stockholm",
                    TermGuid = "65f5b6af-3fd0-4790-966e-2f34ef5c5504"
                });
                oListItem.Update();
            }
            await ctx.ExecuteQueryAsync();
        }

        #endregion 3/5

        #region 3/6

        private static async Task StockholmItemsInSubFolder(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(DocumentLibNameConst);
            ctx.Load(list.RootFolder, f => f.ServerRelativeUrl);
            ctx.ExecuteQuery();
            string targetFolderPath = "Folder 1/Folder 2";
            string targetUrl = $"{list.RootFolder.ServerRelativeUrl}/{targetFolderPath}";
            var items = list.GetItems(new CamlQuery()
            {
                ViewXml = @"<View>
                                <Query>
                                    <OrderBy><FieldRef Name='Modified' Ascending='False'/></OrderBy>
                                    <Where>
                                        <Eq>
                                            <FieldRef Name = 'cities' />
                                            <Value Type = 'TaxonomyFieldTypeMulti'>Stockholm</Value>
                                        </Eq>
                                    </Where>
                                </Query>
                                <RowLimit>20</RowLimit>
                            </View>",
                FolderServerRelativeUrl = targetUrl
            });

            ctx.Load(items);
            await ctx.ExecuteQueryAsync();
        }

        #endregion 3/6

        #region 3/7

        private static async Task UploadDocumentToDocumentLibrary(ClientContext ctx)
        {
            string filePath = @"C:\Document.docx";
            List list = ctx.Web.Lists.GetByTitle(DocumentLibNameConst);
            ctx.Load(list.RootFolder, f => f.ServerRelativeUrl);
            ctx.ExecuteQuery();
            var fileCreationInfo = new FileCreationInformation();
            byte[] FileContent = System.IO.File.ReadAllBytes(filePath);
            fileCreationInfo.ContentStream = new System.IO.MemoryStream(FileContent);
            fileCreationInfo.Overwrite = true;
            fileCreationInfo.Url = $"{list.RootFolder.ServerRelativeUrl}/Document.docx";
            SP.File file = list.RootFolder.Files.Add(fileCreationInfo);
            ctx.ExecuteQuery();
            ListItem listItem = file.ListItemAllFields;
            listItem["Title"] = "Test";
            listItem["ContentTypeId"] = ContentTypeIdConst;
            listItem.Update();
            await ctx.ExecuteQueryAsync();
        }

        #endregion 3/7

        #region 4/1

        private static async Task CreateFolderViewAndSetDefaultView(ClientContext ctx)
        {
            List targetList = ctx.Web.Lists.GetByTitle(DocumentLibNameConst);
            ViewCollection viewCollection = targetList.Views;

            ViewCreationInformation viewCreationInformation = new ViewCreationInformation();
            viewCreationInformation.Title = "Folders";

            // Specify type of the view. Below are the options

            // 1. none - The type of the list view is not specified

            // 2. html - Sspecifies an HTML list view type

            // 3. grid - Specifies a datasheet list view type

            // 4. calendar- Specifies a calendar list view type

            // 5. recurrence - Specifies a list view type that displays recurring events

            // 6. chart - Specifies a chart list view type

            // 7. gantt - Specifies a Gantt chart list view type

            viewCreationInformation.ViewTypeKind = ViewType.Html;

            // You can optionally specify row limit for the view
            viewCreationInformation.RowLimit = 20;

            // You can optionally specify a query as mentioned below.
            // Create one CAML query to filter list view and mention that query below
            viewCreationInformation.Query = "<Where><Eq><FieldRef Name = 'FSObjType' /><Value Type = 'Int'>1</Value></Eq></Where><OrderBy><FieldRef Name='Modified' Ascending='False'/></OrderBy>";

            // Add all the fields over here with comma separated value as mentioned below
            // You can mention display name or internal name of the column
            string CommaSeparateColumnNames = "Type,Name,Modified,Modified By,about,cities,cityCSOM";
            viewCreationInformation.ViewFields = CommaSeparateColumnNames.Split(',');

            View listView = viewCollection.Add(viewCreationInformation);
            ctx.ExecuteQuery();

            // Code to update the display name for the view.
            listView.Title = "Folders";

            // Set default view
            listView.DefaultView = true;

            // You can optionally specify Aggregation: Field references for totals columns or calculated columns
            //listView.Aggregations = "<FieldRef Name='Title' Type='COUNT'/>";

            listView.Update();
            await ctx.ExecuteQueryAsync();
        }

        #endregion 4/1

        #region 4/2

        private static User LoadUserFromEmailOrUserName(ClientContext ctx, string username)
        {
            User result = ctx.Web.EnsureUser(username);
            ctx.Load(result);
            ctx.ExecuteQuery();
            return result;
        }

        #endregion 4/2

        #endregion CSOM & CAML Training

        #region Permission Training

        #region Bb1

        private static Web LoadSubsite(ClientContext ctx)
        {
            Web web = ctx.Web;
            WebCollection webs = web.Webs;
            var query = from subweb in webs where subweb.Title == "Finance and Accounting" select subweb;
            var result = ctx.LoadQuery(query);
            ctx.ExecuteQuery();
            return result.FirstOrDefault();
        }

        private static void BreakRoleInheritanceForSite(ClientContext ctx)
        {
            Web web = LoadSubsite(ctx);
            web.BreakRoleInheritance(true, false);
            ctx.Load(web);
            ctx.ExecuteQuery();
        }

        /// <summary>
        /// A.K.A restore inheritance
        /// </summary>
        /// <param name="ctx">
        /// ClientContext
        /// </param>
        private static void DeleteUniquePermissions(ClientContext ctx)
        {
            Web web = LoadSubsite(ctx);
            web.ResetRoleInheritance();
            ctx.Load(web);
            ctx.ExecuteQuery();
        }

        private static void GetGroupDefaultWhenAddingNewUser(ClientContext ctx)
        {
            Web web = ctx.Web;
            ctx.Load(web, w => w.AssociatedMemberGroup, w => w.Title);

            ctx.ExecuteQuery();

            Console.WriteLine("AssociatedMemberGroup is the default group for new users \"" + web.Title + "\"");

            Console.WriteLine("*************************************************");

            Console.WriteLine("AssociatedMemberGroup in this site is: " + web.AssociatedMemberGroup.Title);
        }

        #endregion Bb1

        #region Bb2

        private static void CreateTestLevelPermissionLevel(ClientContext ctx)
        {
            Web web = ctx.Web;
            BasePermissions basePermissions = new BasePermissions();
            // requested permissions
            basePermissions.Set(PermissionKind.ManageLists);
            basePermissions.Set(PermissionKind.CreateAlerts);
            // associated permissions
            basePermissions.Set(PermissionKind.ViewPages);
            basePermissions.Set(PermissionKind.ViewListItems);
            basePermissions.Set(PermissionKind.Open);
            RoleDefinitionCreationInformation roleDefinitionCreationInfo = new RoleDefinitionCreationInformation();
            roleDefinitionCreationInfo.BasePermissions = basePermissions;
            roleDefinitionCreationInfo.Name = "Test Level";
            roleDefinitionCreationInfo.Description = "Testing permission level with manage lists and create alerts.";
            RoleDefinition roleDefinition = ctx.Web.RoleDefinitions.Add(roleDefinitionCreationInfo);
            ctx.ExecuteQuery();
        }

        #endregion Bb2

        #region Bb3

        // This task is just examine how to add a new user to a existing group on UI

        #endregion Bb3

        #region Bb4

        private static Group CreateTestGroupAndAddUser(ClientContext ctx)
        {
            Web web = ctx.Web;
            GroupCreationInformation groupCI = new GroupCreationInformation();
            groupCI.Title = "Test Group";
            groupCI.Description = "Testing group for Test Level permission level";
            Group group = web.SiteGroups.Add(groupCI);
            group.Owner = web.EnsureUser("admin@mystartpoint.onmicrosoft.com");
            group.Users.AddUser(web.EnsureUser("quoc.lien.hiep@preciofishbone.se"));
            group.Update();
            ctx.ExecuteQuery();
            return group;
        }

        private static void CreateTestGroupWithTestLevelAndAddUser(ClientContext ctx)
        {
            Web web = ctx.Web;
            ctx.Load(web, w => w.RoleDefinitions);
            ctx.ExecuteQuery();
            RoleDefinitionCollection roleDefinitions = web.RoleDefinitions;
            RoleDefinition testLevel = roleDefinitions.GetByName("Test Level");
            ctx.Load(testLevel);
            ctx.ExecuteQuery();
            web.RoleAssignments.Add(CreateTestGroupAndAddUser(ctx), new RoleDefinitionBindingCollection(ctx) { testLevel });
            ctx.ExecuteQuery();
        }

        private static void DeleteGroupFromSite(ClientContext ctx)
        {
            try
            {
                Group oGroup = ctx.Web.SiteGroups.GetByName("Test Group");

                // Load the group
                ctx.Load(oGroup);
                ctx.ExecuteQuery();

                // Remove group
                ctx.Web.SiteGroups.Remove(oGroup);
                ctx.ExecuteQuery();
            }
            catch (Exception e)
            {
            }
        }

        #endregion Bb4

        #region Bb5

        private static void GetInheritedGroupFromSubsite(ClientContext ctx)
        {
            Web subsite = LoadSubsite(ctx);
            GroupCollection groups = subsite.SiteGroups;
            var query = from g in groups where g.Title == "Test Group" select g;
            var result = ctx.LoadQuery(query);
            ctx.ExecuteQuery();
            Group group = result.FirstOrDefault();
            if (group != null)
            {
                RoleAssignmentCollection roleAssignments = subsite.RoleAssignments;
                ctx.Load(subsite, w => w.RoleAssignments.Where(ra => ra.Member.LoginName == group.LoginName));
                ctx.ExecuteQuery();
                String level = "";
                foreach (var ra in roleAssignments)
                {
                    ctx.Load(ra.Member);
                    ctx.Load(ra.RoleDefinitionBindings);
                    ctx.ExecuteQuery();
                    RoleDefinition? definition = ra.RoleDefinitionBindings.FirstOrDefault(rd => rd.Name == "Test Level");
                    if (definition != null)
                    {
                        level = definition.Name;
                    }
                }
                if (level != "")
                {
                    Console.WriteLine($"Found a inherited group named {group.Title} with permission level {level}");
                }
            }
        }

        #endregion Bb5

        #endregion Permission Training

        #region User Profile Training

        private static void UpdateTestTextPropertyInCurrentUserProfile(ClientContext ctx, string newValue)
        {
            PeopleManager peopleManager = new PeopleManager(ctx);
            PersonProperties personProperties = peopleManager.GetMyProperties();
            ctx.Load(personProperties, p => p.AccountName);
            ctx.ExecuteQuery();

            peopleManager.SetSingleValueProfileProperty(personProperties.AccountName, "TestText", newValue);
            ctx.ExecuteQuery();
        }

        private static void UpdateTestBoolPropertyInCurrentUserProfile(ClientContext ctx, string newValue)
        {
            PeopleManager peopleManager = new PeopleManager(ctx);
            PersonProperties personProperties = peopleManager.GetMyProperties();
            ctx.Load(personProperties, p => p.AccountName);
            ctx.ExecuteQuery();

            peopleManager.SetSingleValueProfileProperty(personProperties.AccountName, "TestBool", newValue);
            ctx.ExecuteQuery();
        }

        private static void UpdateTestEmailPropertyInCurrentUserProfile(ClientContext ctx, string newValue)
        {
            PeopleManager peopleManager = new PeopleManager(ctx);
            PersonProperties personProperties = peopleManager.GetMyProperties();
            ctx.Load(personProperties, p => p.AccountName);
            ctx.ExecuteQuery();

            peopleManager.SetSingleValueProfileProperty(personProperties.AccountName, "TestEmail", newValue);
            ctx.ExecuteQuery();
        }

        private static void UpdateTestDatePropertyInCurrentUserProfile(ClientContext ctx, string newValue)
        {
            PeopleManager peopleManager = new PeopleManager(ctx);
            PersonProperties personProperties = peopleManager.GetMyProperties();
            ctx.Load(personProperties, p => p.AccountName);
            ctx.ExecuteQuery();

            peopleManager.SetSingleValueProfileProperty(personProperties.AccountName, "TestDate", newValue);
            ctx.ExecuteQuery();
        }

        private static void UpdateTestIntegerPropertyInCurrentUserProfile(ClientContext ctx, string newValue)
        {
            PeopleManager peopleManager = new PeopleManager(ctx);
            PersonProperties personProperties = peopleManager.GetMyProperties();
            ctx.Load(personProperties, p => p.AccountName);
            ctx.ExecuteQuery();

            peopleManager.SetSingleValueProfileProperty(personProperties.AccountName, "TestInteger", newValue);
            ctx.ExecuteQuery();
        }

        private static void UpdateTestPersonPropertyInCurrentUserProfile(ClientContext ctx, string newValue)
        {
            User user = ctx.Web.EnsureUser(newValue);
            PeopleManager peopleManager = new PeopleManager(ctx);
            PersonProperties personProperties = peopleManager.GetMyProperties();
            ctx.Load(user, u => u.LoginName);
            ctx.ExecuteQuery();
            PersonProperties personProperties1 = peopleManager.GetPropertiesFor(user.LoginName);
            ctx.Load(personProperties, p => p.AccountName);
            ctx.Load(personProperties1, p => p.AccountName);
            ctx.ExecuteQuery();

            peopleManager.SetSingleValueProfileProperty(personProperties.AccountName, "TestPerson", personProperties1.AccountName);
            ctx.ExecuteQuery();
        }

        private static void UpdateTestSingleTaxonomyPropertyInCurrentUserProfile(ClientContext ctx, string newValue)
        {
            PeopleManager peopleManager = new PeopleManager(ctx);
            PersonProperties personProperties = peopleManager.GetMyProperties();
            ctx.Load(personProperties, p => p.AccountName);
            ctx.ExecuteQuery();

            peopleManager.SetSingleValueProfileProperty(personProperties.AccountName, "TestSingleTaxonomy", newValue);
            ctx.ExecuteQuery();
        }

        private static void UpdateTestMultipleTaxonomyPropertyInCurrentUserProfile(ClientContext ctx, List<string> newValue)
        {
            PeopleManager peopleManager = new PeopleManager(ctx);
            PersonProperties personProperties = peopleManager.GetMyProperties();
            ctx.Load(personProperties, p => p.AccountName);
            ctx.ExecuteQuery();

            peopleManager.SetMultiValuedProfileProperty(personProperties.AccountName, "TestMultipleTaxonomy", newValue);
            ctx.ExecuteQuery();
        }

        private static void UserProfileConsoleApp(ClientContext ctx)
        {
            int ans0;
            while (true)
            {
                Console.Clear();
                Console.WriteLine("Choose the user profile property you want to update");
                Console.WriteLine("TestText : 1");
                Console.WriteLine("TestEmail : 2");
                Console.WriteLine("TestDate : 3");
                Console.WriteLine("TestBool : 4");
                Console.WriteLine("TestInteger : 5");
                Console.WriteLine("TestPerson : 6");
                Console.WriteLine("TestSingleTaxonomy : 7");
                Console.WriteLine("TestMultipleTaxonomy : 8");
                string temp = Console.ReadLine();
                if (int.TryParse(temp, out ans0))
                {
                    break;
                }
            }
            switch (ans0)
            {
                case 1:
                    while (true)
                    {
                        try
                        {
                            Console.WriteLine("New value for TestText. Notice there's a limit in string length.");
                            string temp = Console.ReadLine();
                            UpdateTestTextPropertyInCurrentUserProfile(ctx, temp);
                            break;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Please try again.");
                            Console.WriteLine("=====================================");
                        }
                    }
                    break;

                case 2:
                    while (true)
                    {
                        try
                        {
                            Console.WriteLine("New value for TestEmail. Notice the email format.");
                            string temp = Console.ReadLine();
                            UpdateTestEmailPropertyInCurrentUserProfile(ctx, temp);
                            break;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Please try again.");
                            Console.WriteLine("=====================================");
                        }
                    }
                    break;

                case 3:
                    while (true)
                    {
                        try
                        {
                            Console.WriteLine("New value for TestDate. The format is YYYY-MM-DD or MM-DD-YYYY.");
                            string temp = Console.ReadLine();
                            UpdateTestDatePropertyInCurrentUserProfile(ctx, temp);
                            break;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Please try again.");
                            Console.WriteLine("=====================================");
                        }
                    }
                    break;

                case 4:
                    while (true)
                    {
                        try
                        {
                            Console.WriteLine("New value for TestBool(true/false).");
                            string temp = Console.ReadLine();
                            UpdateTestBoolPropertyInCurrentUserProfile(ctx, temp);
                            break;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Please try again.");
                            Console.WriteLine("=====================================");
                        }
                    }
                    break;

                case 5:
                    while (true)
                    {
                        try
                        {
                            Console.WriteLine("New value for TestInteger.");
                            string temp = Console.ReadLine();
                            UpdateTestIntegerPropertyInCurrentUserProfile(ctx, temp);
                            break;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Please try again.");
                            Console.WriteLine("=====================================");
                        }
                    }
                    break;

                case 6:
                    while (true)
                    {
                        try
                        {
                            Console.WriteLine("New value for TestPerson. Currently available: Quoc Lien and User Unknown");
                            string temp = Console.ReadLine();
                            UpdateTestPersonPropertyInCurrentUserProfile(ctx, temp);
                            break;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Please try again.");
                            Console.WriteLine("=====================================");
                        }
                    }
                    break;

                case 7:
                    while (true)
                    {
                        try
                        {
                            Console.WriteLine("New value for TestSingleTaxonomy. Use term name for this language.");
                            string temp = Console.ReadLine();
                            UpdateTestSingleTaxonomyPropertyInCurrentUserProfile(ctx, temp);
                            break;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Please try again.");
                            Console.WriteLine("=====================================");
                        }
                    }
                    break;

                case 8:
                    while (true)
                    {
                        try
                        {
                            Console.WriteLine("New value for TestMultipleTaxonomy. Use term names for this language. Separate names by comma.");
                            string temp = Console.ReadLine();
                            List<string> result = temp.Split(",").ToList();
                            UpdateTestMultipleTaxonomyPropertyInCurrentUserProfile(ctx, result);
                            break;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Please try again.");
                            Console.WriteLine("=====================================");
                        }
                    }
                    break;
            }
            Console.WriteLine("Update successfully. Please check use profile admin");
        }

        #endregion User Profile Training

        #region Search Training

        private static void Search(ClientContext ctx)
        {
            QueryHandler queryHandler = new QueryHandler(ctx);
            List<int> propIndex = new List<int>();
            List<string> chaining = new List<string>();
            List<string> filter = new List<string>();
            bool isEnd = false;
            bool fullText = false;
            var fullTextQuery = "";
            while (true)
            {
                try
                {
                    Console.WriteLine("===============================");
                    Console.WriteLine("Full-text search or property search?");
                    Console.WriteLine("Full-text: 0");
                    Console.WriteLine("Property: 1");
                    fullText = Console.ReadLine() == "0";
                    break;
                }
                catch (Exception e)
                {
                    Console.WriteLine("Wrong input. Please try again!");
                }
            }
            if (fullText)
            {
                Console.WriteLine("Input the value you want to search:");
                fullTextQuery = Console.ReadLine();
            }
            else
            {
                while (!isEnd)
                {
                    while (true)
                    {
                        Console.WriteLine("=============================");
                        queryHandler.ShowAllPropertiesAndTheirIndexes();
                        Console.WriteLine("=============================");
                        Console.WriteLine("Choose a number associate with the property you want to search.");
                        string prop = Console.ReadLine();
                        if (int.TryParse(prop, out _))
                        {
                            propIndex.Add(int.Parse(prop));
                            break;
                        }
                        else
                        {
                            Console.WriteLine("Number only. Please try again!");
                        }
                    }
                    Console.WriteLine($"Searching {queryHandler.GetDisplayNameByIndex(propIndex.ElementAt(propIndex.Count - 1))} with value:");
                    string value = Console.ReadLine();
                    queryHandler.SetPropertyValueByIndex(propIndex.ElementAt(propIndex.Count - 1), value);
                    while (true)
                    {
                        Console.WriteLine($"Include other property?");
                        Console.WriteLine($"No : 0");
                        Console.WriteLine($"AND : 1");
                        Console.WriteLine($"OR : 2");
                        var input = Console.ReadLine();
                        if (int.TryParse(input, out _))
                        {
                            int temp = int.Parse(input);
                            switch (temp)
                            {
                                case 1:
                                    chaining.Add("AND");
                                    break;

                                case 2:
                                    chaining.Add("OR");
                                    break;

                                default:
                                    isEnd = true;
                                    break;
                            }
                            break;
                        }
                        else
                        {
                            Console.WriteLine("Number only. Please try again!");
                            Console.WriteLine("===================================");
                        }
                    }
                }
            }
            while (true)
            {
                Console.WriteLine("===================================");
                Console.WriteLine("Add a time filter?");
                Console.WriteLine($"No : 0");
                Console.WriteLine($"A specific date : 1");
                Console.WriteLine($"A time range : 2");
                var range = Console.ReadLine();
                if (int.TryParse(range, out _))
                {
                    int temp = int.Parse(range);
                    switch (temp)
                    {
                        case 1:
                            bool success = false;
                            while (!success)
                            {
                                try
                                {
                                    Console.WriteLine("Day:");
                                    int day = int.Parse(Console.ReadLine());
                                    Console.WriteLine("Month:");
                                    int month = int.Parse(Console.ReadLine());
                                    Console.WriteLine("Year:");
                                    int year = int.Parse(Console.ReadLine());
                                    filter.Add($"{year}-{month}-{day}");
                                    success = true;
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine("Number only. Please try again!");
                                    Console.WriteLine("===================================");
                                    success = false;
                                }
                            }
                            break;

                        case 2:
                            while (true)
                            {
                                try
                                {
                                    Console.WriteLine("From day:");
                                    int fromDay = int.Parse(Console.ReadLine());
                                    Console.WriteLine("From month:");
                                    int fromMonth = int.Parse(Console.ReadLine());
                                    Console.WriteLine("From year:");
                                    int fromYear = int.Parse(Console.ReadLine());
                                    filter.Add($"{fromYear}-{fromMonth}-{fromDay}");
                                    Console.WriteLine("To day:");
                                    int toDay = int.Parse(Console.ReadLine());
                                    Console.WriteLine("To month:");
                                    int toMonth = int.Parse(Console.ReadLine());
                                    Console.WriteLine("To year:");
                                    int toYear = int.Parse(Console.ReadLine());
                                    filter.Add($"{toYear}-{toMonth}-{toDay}");
                                    break;
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine("Number only. Please try again!");
                                    Console.WriteLine("===================================");
                                }
                            }
                            break;

                        default:
                            break;
                    }
                    break;
                }
                else
                {
                    Console.WriteLine("Number only. Please try again!");
                    Console.WriteLine("===================================");
                }
            }
            if (fullText)
            {
                queryHandler.PerformingFullTextSearch(fullTextQuery, filter);
            }
            else
            {
                queryHandler.PerformPropertySearch(propIndex, chaining, filter);
            }
        }

        private static void SearchConsoleApp(ClientContext ctx)
        {
            while (true)
            {
                Console.Clear();
                Search(ctx);
                bool restart = true;
                while (true)
                {
                    Console.WriteLine("==============================");
                    Console.WriteLine("Do you want to try again?");
                    Console.WriteLine("Yes : 0");
                    Console.WriteLine("No : 1");
                    var temp = Console.ReadLine();
                    if (int.TryParse(temp, out _))
                    {
                        var ans = int.Parse(temp);
                        if (ans == 1)
                        {
                            restart = false;
                        }
                        break;
                    }
                    else
                    {
                        Console.WriteLine("Number only. Please try again!");
                    }
                }
                if (!restart)
                {
                    break;
                }
            }
        }

        #endregion Search Training

        private static ClientContext GetContext(ClientContextHelper clientContextHelper)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();
            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }

        private static async Task GetFieldTermValue(ClientContext Ctx, string termId)
        {
            //load term by id
            TaxonomySession session = TaxonomySession.GetTaxonomySession(Ctx);
            Term taxonomyTerm = session.GetTerm(new Guid(termId));
            Ctx.Load(taxonomyTerm, t => t.Labels,
                                   t => t.Name,
                                   t => t.Id);
            await Ctx.ExecuteQueryAsync();
        }

        private static async Task ExampleSetTaxonomyFieldValue(ListItem item, ClientContext ctx)
        {
            var field = ctx.Web.Fields.GetByTitle("fieldname");

            ctx.Load(field);
            await ctx.ExecuteQueryAsync();

            var taxField = ctx.CastTo<TaxonomyField>(field);

            taxField.SetFieldValueByValue(item, new TaxonomyFieldValue()
            {
                WssId = -1, // alway let it -1
                Label = "correct label here",
                TermGuid = "term id"
            });
            item.Update();
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CsomTermSetAsync(ClientContext ctx)
        {
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            // Get the term store by name
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get the term group by Name
            TermGroup termGroup = termStore.Groups.GetByName("Test");
            // Get the term set by Name
            TermSet termSet = termGroup.TermSets.GetByName("Test Term Set");

            var terms = termSet.GetAllTerms();

            ctx.Load(terms);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CsomLinqAsync(ClientContext ctx)
        {
            var fieldsQuery = from f in ctx.Web.Fields
                              where f.InternalName == "Test" ||
                                    f.TypeAsString == "TaxonomyFieldTypeMulti" ||
                                    f.TypeAsString == "TaxonomyFieldType"
                              select f;

            var fields = ctx.LoadQuery(fieldsQuery);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task SimpleCamlQueryAsync(ClientContext ctx)
        {
            var list = ctx.Web.Lists.GetByTitle("Documents");

            var allItemsQuery = CamlQuery.CreateAllItemsQuery();
            var allFoldersQuery = CamlQuery.CreateAllFoldersQuery();

            var items = list.GetItems(new CamlQuery()
            {
                ViewXml = @"<View>
                                <Query>
                                    <OrderBy><FieldRef Name='ID' Ascending='False'/></OrderBy>
                                </Query>
                                <RowLimit>20</RowLimit>
                            </View>",
                FolderServerRelativeUrl = "/sites/test-site-duc-11111/Shared%20Documents/2"
                //example for site: https://omniapreprod.sharepoint.com/sites/test-site-duc-11111/
            });

            ctx.Load(items);
            await ctx.ExecuteQueryAsync();
        }
    }
}
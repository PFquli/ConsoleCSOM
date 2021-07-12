using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Collections.Generic;

using SP = Microsoft.SharePoint.Client;

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
                    await AddFieldToContentType(ctx);
                    await AddContentTypeToList(ctx);
                    await CreateNewListItems(ctx);
                }

                Console.WriteLine($"Press Any Key To Stop!");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
            }
        }

        private static async Task CreateCSOMTestList(ClientContext ctx)
        {
            Web web = ctx.Web;
            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.TemplateType = (int)ListTemplateType.GenericList;
            creationInfo.Description = web.Description;
            creationInfo.Title = "CSOM Test";
            try
            {
                SP.List newList = web.Lists.Add(creationInfo);

                await ctx.ExecuteQueryAsync();
            }
            catch (Exception ex)
            { }
        }

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

        private async static Task CreateSiteFields(ClientContext ctx)
        {
            try
            {
                Web web = ctx.Web;
                FieldCollection fields = web.Fields;
                fields.AddFieldAsXml(
                    @"<Field ID='139B1AB0-0EDC-4D4B-8B35-632CED9F3DCD' Type='Text'
                                            Name='about'
                                            Required='FALSE'
                                            DisplayName='about'
                                            Description=''
                                            Group='Custom Columns'/>", true, AddFieldOptions.DefaultValue);
                fields.AddFieldAsXml(
                    @"<Field ID='2E660010-96AE-4CA1-895B-DE92AB67451F' Type='TaxonomyFieldType'
                                            Name='cityCSOM'
                                            Required='FALSE'
                                            DisplayName='cityCSOM'
                                            Description=''
                                            Hidden='FALSE'
                                            Group='Custom Columns'/>", true, AddFieldOptions.DefaultValue);
                await ctx.ExecuteQueryAsync();
            }
            catch (Exception e)
            {
            }
        }

        private async static Task CreateContentType(ClientContext ctx)
        {
            ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
            ctx.Load(contentTypes);
            ctx.ExecuteQuery();
            string id = "0x0100" + new Guid("7CDA06D1-66B4-4450-8FC3-28CD64FB2C3C").ToString("N"); // parent is Item type

            foreach (var item in contentTypes)
            {
                if (item.StringId == id)
                    return;
            }

            // Create a Content Type Information object.
            ContentTypeCreationInformation newCt = new ContentTypeCreationInformation();
            // Set the name for the content type.
            newCt.Name = "CSOM Test content type";
            // Inherit from oob document - 0x0101 and assign.
            newCt.Id = id;
            // Set content type to be available from specific group.
            newCt.Group = "Custom Content Types";
            // Create the content type.
            ContentType testContentType = contentTypes.Add(newCt);
            await ctx.ExecuteQueryAsync();
        }

        private async static Task AddFieldToContentType(ClientContext ctx)
        {
            ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
            ctx.Load(contentTypes);
            ctx.ExecuteQuery();
            // Give content type name over here
            ContentType testContentType = (from contentType in contentTypes where contentType.Name == "CSOM Test content type" select contentType).FirstOrDefault();

            ctx.Load(testContentType);
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
            ctx.Load(contentTypes);
            ctx.ExecuteQuery();
            // Give content type name over here
            ContentType testContentType = (from contentType in contentTypes where contentType.Name == "CSOM Test content type" select contentType).FirstOrDefault();

            ctx.Load(testContentType);
            // Get list
            List testList = ctx.Web.Lists.GetByTitle("CSOM Test");
            // Add content type to list and update
            testList.ContentTypes.AddExistingContentType(testContentType);
            testList.Update();
            ctx.Web.Update();
            await ctx.ExecuteQueryAsync();
        }

        internal class ListItemModel
        {
            public ListItemModel()
            {
            }

            public string Title { get; set; }
            public string About { get; set; }
            public string City { get; set; }
        }

        private async static Task CreateNewListItems(ClientContext ctx)
        {
            List<ListItemModel> listItems = new List<ListItemModel>
            {
                new ListItemModel{Title="ListItem1", About="City1", City="Ho Chi Minh"},
                new ListItemModel{Title="ListItem2", About="City2", City="Paris"},
                new ListItemModel{Title="ListItem3", About="City3", City="London"},
                new ListItemModel{Title="ListItem4", About="City4", City="Can Tho"},
                new ListItemModel{Title="ListItem5", About="City5", City="New York"}
            };
            SP.List oList = ctx.Web.Lists.GetByTitle("CSOM Test");
            ctx.Load(oList);
            ctx.ExecuteQuery();
            foreach (ListItemModel model in listItems)
            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = oList.AddItem(itemCreateInfo);
                oListItem["Title"] = model.Title;
                oListItem["about"] = model.About;
                oListItem["cityCSOM"] = model.City;
                oListItem.Update();
            }
            await ctx.ExecuteQueryAsync();
        }

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
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;

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

                    await CreateCSOMTestList(ctx);
                    //await SimpleCamlQueryAsync(ctx);
                    //await CsomTermSetAsync(ctx);
                    await CreateTermSetInDevTenant(ctx);
                    await CreateNewTerms(ctx);
                    await CreateSiteFields(ctx);
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
                                            Name='city CSOM test'
                                            Required='FALSE'
                                            DisplayName='citycsomtest'
                                            Description=''
                                            Hidden='FALSE'
                                            Group='Custom Columns'/>", true, AddFieldOptions.DefaultValue);
                await ctx.ExecuteQueryAsync();
            }
            catch (Exception e)
            {
            }
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
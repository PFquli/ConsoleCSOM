using ConsoleCSOM.Models;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleCSOM
{
    internal class QueryHandler
    {
        private ManagedProperty[] arrayOfManagedProperties { get; set; }
        private readonly Microsoft.SharePoint.Client.ClientContext ctx;

        public QueryHandler(ClientContext ctx)
        {
            arrayOfManagedProperties = new ManagedProperty[]
            {
            new ManagedProperty("First Name", "RefinableString00"),
            new ManagedProperty("Book Genre", "RefinableString01"),
            new ManagedProperty("Book Category", "RefinableString01"),
            new ManagedProperty("Is Member", "RefinableString02"),
            new ManagedProperty("Group Leaders", "RefinableString03"),
            new ManagedProperty("Group Monitors", "RefinableString03"),
            new ManagedProperty("Return Date", "RefinableDate00"),
            new ManagedProperty("Borrow End Date", "RefinableDate00"),
            new ManagedProperty("Borrowed Book Quantity", "RefinableInt00"),

            new ManagedProperty("User Custom Text", "RefinableString04"),
            new ManagedProperty("User Custom Email", "RefinableString05"),
            new ManagedProperty("User Custom Boolean", "RefinableString06"),
            new ManagedProperty("User Custom Person", "RefinableString07"),
            new ManagedProperty("User Custom Single Taxonomy", "RefinableString08"),
            new ManagedProperty("User Custom Multiple Taxonomy", "RefinableString10")
            };
            this.ctx = ctx;
        }

        public void ShowAllPropertiesAndTheirIndexes()
        {
            for (int i = 0; i < arrayOfManagedProperties.Length; i++)
            {
                ManagedProperty currentProperty = arrayOfManagedProperties[i];
                Console.WriteLine($"{currentProperty.DisplayName} : {i}");
            }
        }

        public string GetDisplayNameByIndex(int i)
        {
            return arrayOfManagedProperties[i].DisplayName;
        }

        public void SetPropertyValueByIndex(int i, string value)
        {
            arrayOfManagedProperties[i].Value = value;
        }

        private void ShowResult(IDictionary<String, Object> resultRow)
        {
            Console.WriteLine("====================================");
            Console.WriteLine($"Title: {resultRow["Title"]} ");
            Console.WriteLine($"Author: {resultRow["Author"]} ");
            Console.WriteLine($"SiteName: {resultRow["SiteName"]}");
            Console.WriteLine($"Path: {resultRow["Path"]}");
        }

        public void PerformSingleSearch(int index)
        {
            KeywordQuery keywordQuery = new KeywordQuery(ctx);
            SearchExecutor searchExecutor = new SearchExecutor(ctx);
            keywordQuery.QueryText = $"{arrayOfManagedProperties[index].ManagedPropertyName}:{arrayOfManagedProperties[index].Value}";
            keywordQuery.EnableSorting = true;
            keywordQuery.RowsPerPage = 10;
            keywordQuery.RowLimit = 100;
            keywordQuery.StartRow = 0;
            ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(keywordQuery);
            ctx.ExecuteQuery();
            int trows = results.Value[0].TotalRows;
            if (trows == 0)
            {
                Console.WriteLine("No result found. Please try again.");
                return;
            }
            var resultRows = results.Value[0].ResultRows;
            Console.WriteLine($"Found total {trows} row(s)");
            foreach (var resultRow in resultRows)
            {
                ShowResult(resultRow);
            }
        }

        public void PerformChainingSearch(List<int> propIndex, List<string> chaining)
        {
            StringBuilder query = new StringBuilder();
            for (int i = 0; i < propIndex.Count; i++)
            {
                if (i > 0)
                {
                    query.Append($" {chaining[i - 1]} ");
                }
                query.Append($"{arrayOfManagedProperties[propIndex[i]].ManagedPropertyName}:{arrayOfManagedProperties[propIndex[i]].Value}");
            }
            KeywordQuery keywordQuery = new KeywordQuery(ctx);
            SearchExecutor searchExecutor = new SearchExecutor(ctx);
            Console.WriteLine($"Chaining search query: {query.ToString()}");
            keywordQuery.QueryText = query.ToString();
            keywordQuery.EnableSorting = true;
            keywordQuery.RowsPerPage = 10;
            keywordQuery.RowLimit = 100;
            keywordQuery.StartRow = 0;
            ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(keywordQuery);
            ctx.ExecuteQuery();
            int trows = results.Value[0].TotalRows;
            if (trows == 0)
            {
                Console.WriteLine("No result found. Please try again.");
                return;
            }
            var resultRows = results.Value[0].ResultRows;
            Console.WriteLine($"Found total {trows} row(s)");
            foreach (var resultRow in resultRows)
            {
                ShowResult(resultRow);
            }
        }

        public void PerformSearch(List<int> propIndex, List<string> chaining)
        {
            if (propIndex.Count < 2)
            {
                PerformSingleSearch(propIndex.ElementAt(0));
            }
            else
            {
                PerformChainingSearch(propIndex, chaining);
            }
        }
    }
}
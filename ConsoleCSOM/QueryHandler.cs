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

        public void PerformSingleSearch(int index, List<string> filter)
        {
            KeywordQuery keywordQuery = new KeywordQuery(ctx);
            SearchExecutor searchExecutor = new SearchExecutor(ctx);
            Console.WriteLine("Performing single search for:");
            ManagedProperty searchProp = arrayOfManagedProperties[index];
            if (filter.Count == 0)
            {
                Console.WriteLine($"{searchProp.DisplayName} with value of {searchProp.Value}");
                keywordQuery.QueryText = $"{searchProp.ManagedPropertyName}:{searchProp.Value}";
            }
            else
            {
                string dateManagedPropertyName = arrayOfManagedProperties[6].ManagedPropertyName;
                string dateDisplayName0 = arrayOfManagedProperties[6].DisplayName;
                string dateDisplayName1 = arrayOfManagedProperties[7].DisplayName;
                if (filter.Count > 1)
                {
                    Console.WriteLine($"{searchProp.DisplayName} with value of {searchProp.Value} AND {dateDisplayName0}/{dateDisplayName1} is between {filter.ElementAt(0)} and {filter.ElementAt(1)}");
                    keywordQuery.QueryText = $"{searchProp.ManagedPropertyName}:{searchProp.Value} AND {dateManagedPropertyName}>={filter.ElementAt(0)} AND {dateManagedPropertyName}<={filter.ElementAt(1)}";
                }
                else
                {
                    Console.WriteLine($"{searchProp.DisplayName} with value of {searchProp.Value} AND {dateDisplayName0}/{dateDisplayName1} is {filter.ElementAt(0)}");
                    keywordQuery.QueryText = $"{searchProp.ManagedPropertyName}:{searchProp.Value} AND {dateManagedPropertyName}={filter.ElementAt(0)}";
                }
            }
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

        public void PerformChainingSearch(List<int> propIndex, List<string> chaining, List<string> filter)
        {
            StringBuilder query = new StringBuilder();
            Console.WriteLine("Performing chaining search for:");
            for (int i = 0; i < propIndex.Count; i++)
            {
                ManagedProperty searchProp = arrayOfManagedProperties[propIndex[i]];
                if (i > 0)
                {
                    query.Append($" {chaining[i - 1]} ");
                    Console.Write($" {chaining[i - 1]} ");
                }
                query.Append($"{searchProp.ManagedPropertyName}:{searchProp.Value}");
                Console.Write($"{searchProp.DisplayName} with a value of {searchProp.Value}");
            }
            if (filter.Count != 0)
            {
                string dateManagedPropertyName = arrayOfManagedProperties[6].ManagedPropertyName;
                string dateDisplayName0 = arrayOfManagedProperties[6].DisplayName;
                string dateDisplayName1 = arrayOfManagedProperties[7].DisplayName;
                if (filter.Count > 1)
                {
                    Console.WriteLine($" AND {dateDisplayName0}/{dateDisplayName1} is between {filter.ElementAt(0)} and {filter.ElementAt(1)}");
                    query.Append($" AND {dateManagedPropertyName}>={filter.ElementAt(0)} AND {dateManagedPropertyName}<={filter.ElementAt(1)}");
                }
                else
                {
                    Console.WriteLine($" AND {dateDisplayName0}/{dateDisplayName1} is {filter.ElementAt(0)}");
                    query.Append($" AND {dateManagedPropertyName}={filter.ElementAt(0)}");
                }
            }
            KeywordQuery keywordQuery = new KeywordQuery(ctx);
            SearchExecutor searchExecutor = new SearchExecutor(ctx);
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

        public void PerformingFullTextSearch(string query, string filter)
        {
            KeywordQuery keywordQuery = new KeywordQuery(ctx);
            SearchExecutor searchExecutor = new SearchExecutor(ctx);
            Console.WriteLine("Performing full-text search for:");
            ManagedProperty searchProp = arrayOfManagedProperties[0];
            if (filter == "")
            {
                Console.WriteLine($"{searchProp.DisplayName} with value of {searchProp.Value}");
                keywordQuery.QueryText = $"{searchProp.ManagedPropertyName}:{searchProp.Value}";
            }
            else
            {
                Console.WriteLine($"{searchProp.DisplayName} with value of {searchProp.Value} AND last modified time is {filter}");
                keywordQuery.QueryText = $"{searchProp.ManagedPropertyName}:{searchProp.Value} AND LastModifiedTime={filter}";
            }
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

        public void PerformSearch(List<int> propIndex, List<string> chaining, List<string> filter)
        {
            if (propIndex.Count < 2)
            {
                PerformSingleSearch(propIndex.ElementAt(0), filter);
            }
            else
            {
                PerformChainingSearch(propIndex, chaining, filter);
            }
        }
    }
}
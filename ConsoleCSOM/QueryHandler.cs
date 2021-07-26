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
            Console.WriteLine($"Site Name: {resultRow["SiteName"]}");
            Console.WriteLine($"Path: {resultRow["Path"]}");
            Console.WriteLine($"First Name: {resultRow["RefinableString00"]}");
            Console.WriteLine($"Book Genre/Book Category: {resultRow["RefinableString01"]}");
            Console.WriteLine($"Is Member: {resultRow["RefinableString02"]}");
            Console.WriteLine($"Group Leaders/Monitors: {resultRow["RefinableString03"]}");
            Console.WriteLine($"Return Date/Borrow End Date: {resultRow["RefinableDate00"]}");
            Console.WriteLine($"Borrowed Book Quantity: {resultRow["RefinableInt00"]}");
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
            keywordQuery.SelectProperties.Add("RefinableDate00");
            keywordQuery.SelectProperties.Add("RefinableInt00");
            keywordQuery.SelectProperties.Add("RefinableString00");
            keywordQuery.SelectProperties.Add("RefinableString01");
            keywordQuery.SelectProperties.Add("RefinableString02");
            keywordQuery.SelectProperties.Add("RefinableString03");
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
                Console.WriteLine($"{searchProp.DisplayName} with a value of {searchProp.Value}");
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
            keywordQuery.SelectProperties.Add("RefinableDate00");
            keywordQuery.SelectProperties.Add("RefinableInt00");
            keywordQuery.SelectProperties.Add("RefinableString00");
            keywordQuery.SelectProperties.Add("RefinableString01");
            keywordQuery.SelectProperties.Add("RefinableString02");
            keywordQuery.SelectProperties.Add("RefinableString03");
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

        public void PerformingFullTextSearch(string query, List<string> filter)
        {
            KeywordQuery keywordQuery = new KeywordQuery(ctx);
            SearchExecutor searchExecutor = new SearchExecutor(ctx);
            Console.WriteLine("Performing full-text search for:");
            if (filter.Count == 0)
            {
                Console.WriteLine($"{query}");
                keywordQuery.QueryText = $"{query}";
            }
            else
            {
                string dateManagedPropertyName = arrayOfManagedProperties[6].ManagedPropertyName;
                string dateDisplayName0 = arrayOfManagedProperties[6].DisplayName;
                string dateDisplayName1 = arrayOfManagedProperties[7].DisplayName;
                if (filter.Count > 1)
                {
                    Console.WriteLine($"{query} AND {dateDisplayName0}/{dateDisplayName1} is between {filter.ElementAt(0)} and {filter.ElementAt(1)}");
                    keywordQuery.QueryText = $"{query} AND {dateManagedPropertyName}>={filter.ElementAt(0)} AND {dateManagedPropertyName}<={filter.ElementAt(1)}";
                }
                else
                {
                    Console.WriteLine($"{query} AND {dateDisplayName0}/{dateDisplayName1} is {filter.ElementAt(0)}");
                    keywordQuery.QueryText = $"{query} AND {dateManagedPropertyName}={filter.ElementAt(0)}";
                }
            }
            keywordQuery.EnableSorting = true;
            keywordQuery.RowsPerPage = 10;
            keywordQuery.RowLimit = 100;
            keywordQuery.StartRow = 0;
            keywordQuery.SelectProperties.Add("RefinableDate00");
            keywordQuery.SelectProperties.Add("RefinableInt00");
            keywordQuery.SelectProperties.Add("RefinableString00");
            keywordQuery.SelectProperties.Add("RefinableString01");
            keywordQuery.SelectProperties.Add("RefinableString02");
            keywordQuery.SelectProperties.Add("RefinableString03");
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

        public void PerformPropertySearch(List<int> propIndex, List<string> chaining, List<string> filter)
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
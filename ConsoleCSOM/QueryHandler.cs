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
        private ManagedProperty firstName, bookGenre, bookCategory, isMember, groupLeaders, groupMonitors, returnDate, borrowEndDate, borrowedBookQuantity;
        private ManagedProperty userText, userEmail, userBool, userPerson, userSingleTaxonomy, userMultipleTaxonomy;
        private readonly Microsoft.SharePoint.Client.ClientContext ctx;

        public QueryHandler(ClientContext ctx)
        {
            firstName = new ManagedProperty("First Name", "RefinableString00");
            bookGenre = new ManagedProperty("Book Genre", "RefinableString01");
            bookCategory = new ManagedProperty("Book Category", "RefinableString01");
            isMember = new ManagedProperty("Is Member", "RefinableString02");
            groupLeaders = new ManagedProperty("Group Leaders", "RefinableString03");
            groupMonitors = new ManagedProperty("Group Monitors", "RefinableString03");
            returnDate = new ManagedProperty("Return Date", "RefinableDate00");
            borrowEndDate = new ManagedProperty("Borrow End Date", "RefinableDate00");
            borrowedBookQuantity = new ManagedProperty("Borrowed Book Quantity", "RefinableInt00");

            userText = new ManagedProperty("User Custom Text", "RefinableString04");
            userEmail = new ManagedProperty("User Custom Email", "RefinableString05");
            userBool = new ManagedProperty("User Custom Boolean", "RefinableString06");
            userPerson = new ManagedProperty("User Custom Person", "RefinableString07");
            userSingleTaxonomy = new ManagedProperty("User Custom Single Taxonomy", "RefinableString08");
            userMultipleTaxonomy = new ManagedProperty("User Custom Multiple Taxonomy", "RefinableString10");
            this.ctx = ctx;
        }

        private void ShowResult(IDictionary<String, Object> resultRow)
        {
            Console.WriteLine("====================================");
            Console.WriteLine($"Title: {resultRow["Title"]} ");
            Console.WriteLine($"Author: {resultRow["Author"]} ");
            Console.WriteLine($"SiteName: {resultRow["SiteName"]}");
            Console.WriteLine($"Path: {resultRow["Path"]}");
        }

        public void PerformSingleSearch(string keyword)
        {
            KeywordQuery keywordQuery = new KeywordQuery(ctx);
            SearchExecutor searchExecutor = new SearchExecutor(ctx);
            keywordQuery.QueryText = keyword;
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
            foreach (var resultRow in resultRows)
            {
                ShowResult(resultRow);
            }
        }
    }
}
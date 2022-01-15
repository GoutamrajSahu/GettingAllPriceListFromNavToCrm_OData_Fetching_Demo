using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace GettingAllPriceListFromNavToCrm
{
    class Program
    {

        public class ProductMasterFields
        {
            public String Item_Subcategory_Code { get; set; }
            public String Price_State_Code { get; set; }
            public String Price_Code { get; set; }
            public String Item_No { get; set; }
            public String Unit_Price { get; set; }
            public String Price_Change_Percent  { get; set; }
            public String New_Unit_Price { get; set; }
            public String Start_Date { get; set; }
            public String End_Date { get; set; }
            public String Status { get; set; }
            public String Description { get; set; }
            public String Unit_of_Measure_Code { get; set; }
            public String Item_Wt { get; set; }
            public String Item_Category_Code { get; set; }
            public String Send_Mail { get; set; }
            public String Delete_Reason { get; set; }
            public String Reopen_Reason { get; set; }
            public String Reject_Reason { get; set; }
            public String Product_Group_Code { get; set; }
            public String ETag { get; set; }
        }


        static void Main(string[] args)
        {
            CrmServiceClient service = connect();
            var url = ConfigurationManager.AppSettings["uri"];
            // Create a request for the URL.
            WebRequest request = WebRequest.Create(url);
            var NavUserName = ConfigurationManager.AppSettings["NavUsername"];
            var NavPassword = ConfigurationManager.AppSettings["NavPassword"];
            request.Credentials = new NetworkCredential(NavUserName, NavPassword);
            // Get the response.
            WebResponse response = request.GetResponse();
            using (Stream dataStream = response.GetResponseStream())
            {
                // Open the stream using a StreamReader for easy access.
                StreamReader reader = new StreamReader(dataStream);
                // Read the content.
                string responseFromServer = reader.ReadToEnd();

                var Json = JsonConvert.DeserializeObject<Dictionary<string, object>>(responseFromServer);
                var pp = Json["value"];

                var Json1 = JsonConvert.DeserializeObject<List<ProductMasterFields>>(pp.ToString());
                int x = Json1.Count();
                for (int y = 0; x > y; y++)
                {


                    Entity newPriceListItem = new Entity("zx_pricelistitem");

                    /*<-----------------------------------------------------------Function to Get and Set Lookup Start-------------------------------------------------->*/
                    void findAndSetLookup(string lookFieldLogicalName, string targetEntityLogicalName, string targetEntityAttributeToMatch, string jsonDataToFind)
                    {
                        QueryExpression Query = new QueryExpression()
                        {
                            EntityName = targetEntityLogicalName,
                            ColumnSet = new ColumnSet(true),
                            Criteria = new FilterExpression
                            {
                                FilterOperator = LogicalOperator.And,
                                Conditions =
                            {
                                new ConditionExpression
                                {
                                    AttributeName = targetEntityAttributeToMatch,
                                    Operator = ConditionOperator.Equal,
                                    Values = { jsonDataToFind }
                                }
                            }
                            }
                        };
                        EntityCollection EntityCollection = service.RetrieveMultiple(Query);
                        int countOfEc = EntityCollection.Entities.Count;
                        if (countOfEc != 0)
                        {
                            Guid brandGroup_Id = EntityCollection[0].Id;
                            newPriceListItem.Attributes[lookFieldLogicalName] = new EntityReference(targetEntityLogicalName, brandGroup_Id);
                        }
                    }
                    /*<-----------------------------------------------------------Function to Get and Set Lookup End--------------------------------------------------->*/
                   
                    findAndSetLookup("zx_state", "zx_salesterritory", "zx_territoryname", Json1[y].Price_State_Code);//Json1[y].Price_State_Code; <--zx_state from price list items
                    newPriceListItem.Attributes["zx_pricelistitemcode"] = Json1[y].Price_Code; //Price_Code <--zx_pricelistitemcode from price list items

                    if (!String.Equals(Json1[y].Unit_Price, ""))
                    {
                        decimal amount = decimal.Parse(Json1[y].Unit_Price);
                        newPriceListItem.Attributes["zx_amount"] = amount; //Unit_Price <--zx_amount from price list items
                    }

                    if (!String.Equals(Json1[y].Start_Date, ""))
                    {
                        newPriceListItem.Attributes["zx_startdate"] = DateTime.Parse(Json1[y].Start_Date); //Start_Date <--zx_startdate from price list items
                    }

                    if (!String.Equals(Json1[y].End_Date, ""))
                    {
                        newPriceListItem.Attributes["zx_enddate"] = DateTime.Parse(Json1[y].End_Date); //End_Date <--zx_enddate from price list items/ price list
                    }

                    findAndSetLookup("zx_product", "zx_productmaster", "zx_productname", Json1[y].Description);//Description <--zx_product from price list items

                    findAndSetLookup("zx_uom", "zx_unitofmeasure", "zx_unitofmeasurecode", Json1[y].Unit_of_Measure_Code);//Unit_of_Measure_Code <--zx_uom from price list items

                    try
                    {
                        service.Create(newPriceListItem);
                        Console.WriteLine($"Price List Item Code:{Json1[y].Price_Code}, Product Name: {Json1[y].Description}, Status: Created Successfully in CRM !!!");
                        //Console.WriteLine("Success, Got It");
                    }

                    catch (Exception ex)
                    {
                        Console.WriteLine("Error");
                        Console.WriteLine(ex);
                        Console.WriteLine("Error, Got It");
                        //  Console.WriteLine(responseFromServer);
                    }
                }
                Console.WriteLine("All product from NAV Created Successfully in CRM !!!");
            }

        }

        public static CrmServiceClient connect()
        {
            var url = ConfigurationManager.AppSettings["url"];
            var userName = ConfigurationManager.AppSettings["username"];
            var password = ConfigurationManager.AppSettings["password"];

            string conn = $@"  Url = {url}; AuthType = OAuth;
            UserName = {userName};
            Password = {password};
            AppId = 51f81489-12ee-4a9e-aaae-a2591f45987d;
            RedirectUri = app://58145B91-0C36-4500-8554-080854F2AC97;
            LoginPrompt=Auto;
            RequireNewInstance = True";

            var svc = new CrmServiceClient(conn);
            return svc;
        }
    }
}

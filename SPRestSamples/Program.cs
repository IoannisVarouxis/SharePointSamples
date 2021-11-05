/* SPDX-License-Identifier: Apache-2.0 */
/* Copyright © 2021 - Ioannis Varouxis */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPRestSamples
{
    class Program
    {
        static void Main(string[] args)
        {
            string url = "http://sp2013";
            string listName = "Εργασίες";
            //string listName = "Tasks";
            string content = null;
            string listItemTitle = null;
            string listItemType = null;
            int listItemId = -1;

            Random rnd = new Random(DateTime.Now.Millisecond);
            bool exit = false;

            while (!exit)
            {
                Console.WriteLine("1. Get list items");
                Console.WriteLine("2. Get list item type");
                Console.WriteLine("3. Create new list item");
                Console.WriteLine("4. Update newly created list item");
                Console.WriteLine("5. Execute CAML query");
                Console.WriteLine("0. Exit");
                Console.WriteLine();
                Console.Write("Select action [1 - 5]: ");
                string command = Console.ReadLine();

                Console.WriteLine();

                switch (command)
                {
                    case "1":
                        // Get list items
                        content = SPRestGetItems(url, listName).Result;
                        Console.WriteLine(content);

                        break;
                    case "2":
                        // Get list item type (used for adding new items to list)
                        // For this to work the list must not be empty
                        listItemType = GetItemType(url, listName).Result;
                        Console.WriteLine("List Item Type: " + listItemType);

                        break;
                    case "3":
                        if (!string.IsNullOrWhiteSpace(listItemType))
                        {
                            // Create new list item
                            Tuple<string, object>[] data = new Tuple<string, object>[] {
                                new Tuple<string, object>("Title", "Test - " + rnd.Next(int.MaxValue)),
                                new Tuple<string, object>("AssignedToId", new int[] { 15, 16 }), // Multiuser field, these must be valid user ids in SP site
                                new Tuple<string, object>("StartDate", DateTime.Today.ToString("yyyy-MM-dd")),
                                new Tuple<string, object>("DueDate", DateTime.Today.AddDays(15).ToString("yyyy-MM-dd"))
                            };
                            listItemTitle = data[0].Item2.ToString();
                            listItemId = SPRestCreateItem(url, listName, data, listItemType).Result;
                            Console.WriteLine("New List Item Id: " + listItemId);
                        }
                        else
                            Console.WriteLine("INVALID: Get the List Item Type first.");

                        break;
                    case "4":
                        if (listItemId > 0)
                        {
                            string newtitle = listItemTitle + " - UPDATED";
                            // Update newly created list item
                            Tuple<string, object>[] dataForUpdate = new Tuple<string, object>[] {
                                new Tuple<string, object>("Title", newtitle)
                            };
                            bool updated = SPRestUpdateItem(url, listName, listItemId, dataForUpdate, listItemType).Result;
                            Console.WriteLine(string.Format("Update List Item Result: {4}{0}    Id: {1}{0}    Previous Title: '{2}'{0}    New Title: '{3}'", Environment.NewLine, listItemId, listItemTitle, newtitle, updated));
                        }
                        else
                            Console.WriteLine("INVALID: Create a new list item first.");

                        break;
                    case "5":
                        // Execute CAML query
                        string query = "<View><Query></Query><RowLimit>1</RowLimit><ViewFields><FieldRef Name='ID' /><FieldRef Name='Title' /><FieldRef Name='AssignedTo' /></ViewFields></View>";
                        content = SPRestExecuteCamlQuery(url, listName, query).Result;
                        Console.WriteLine(string.Format("CAML query : {0}    {1}{0}", Environment.NewLine, query));
                        Console.WriteLine(content);

                        break;
                    case "0":
                        exit = true;

                        break;
                    default:
                        Console.WriteLine(Environment.NewLine + "Invalid command." + Environment.NewLine + "Select action [1 - 5]: ");
                        break;
                }
                Console.WriteLine();
            }
        }

        private static string GetSPFormDigest(string url)
        {
            string digest = "";

            System.Net.Http.HttpClientHandler clientHandler = new System.Net.Http.HttpClientHandler()
            {
                UseDefaultCredentials = true
            };
            // Or define network credentials
            //clientHandler.Credentials = new System.Net.NetworkCredential("username", "pwd", "DOMAIN");

            System.Net.Http.HttpClient client = new System.Net.Http.HttpClient(clientHandler);

            client.BaseAddress = new System.Uri(url);
            string cmd = "_api/contextinfo";
            client.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");
            client.DefaultRequestHeaders.Add("ContentType", "application/json");
            client.DefaultRequestHeaders.Add("ContentLength", "0");
            System.Net.Http.StringContent httpContent = new System.Net.Http.StringContent("");
            var response = client.PostAsync(cmd, httpContent).Result;
            if (response.IsSuccessStatusCode)
            {
                string content = response.Content.ReadAsStringAsync().Result;
                SPRestProxyClasses.RootObject obj = Newtonsoft.Json.JsonConvert.DeserializeObject<SPRestProxyClasses.RootObject>(content);
                if (obj != null)
                {
                    digest = obj.d.GetContextWebInformation.FormDigestValue;
                }

                // Alternate method, no proxy classes required
                //Newtonsoft.Json.Linq.JToken t = Newtonsoft.Json.Linq.JToken.Parse(content);
                //digest = t["d"]["GetContextWebInformation"]["FormDigestValue"].ToString();
            }

            return digest;
        }

        private static async Task<Newtonsoft.Json.Linq.JToken> SPRestGetItemsJson(string url, string listName)
        {
            try
            {
                System.Net.Http.HttpClientHandler clientHandler = new System.Net.Http.HttpClientHandler()
                {
                    UseDefaultCredentials = true
                };
                // Or define network credentials
                //clientHandler.Credentials = new System.Net.NetworkCredential("username", "pwd", "DOMAIN");

                System.Net.Http.HttpClient client = new System.Net.Http.HttpClient(clientHandler);

                string digest = GetSPFormDigest(url);

                client.BaseAddress = new System.Uri(url);
                client.DefaultRequestHeaders.Clear();
                client.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");
                client.DefaultRequestHeaders.Add("X-RequestDigest", digest);

                // UrlEncode for plain english characters is not required, but it is for for non-english characters and various symbols
                string encodedListName = System.Web.HttpUtility.UrlEncode(listName);

                System.Net.Http.HttpResponseMessage response = await client.GetAsync("_api/web/lists/GetByTitle('" + encodedListName + "')/items");
                string spRequestId = response.Headers.GetValues("request-id").First(); // Correlation ID in ULS logs

                if (response.IsSuccessStatusCode)
                {
                    string text = await response.Content.ReadAsStringAsync();

                    Newtonsoft.Json.Linq.JToken parsedJson = Newtonsoft.Json.Linq.JToken.Parse(text);

                    return parsedJson;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            return null;
        }

        private static async Task<string> SPRestGetItems(string url, string listName)
        {
            Newtonsoft.Json.Linq.JToken json = await SPRestGetItemsJson(url, listName);
            if (json != null)
            {
                return json.ToString(Newtonsoft.Json.Formatting.Indented);
            }
            else return null;
        }

        private static async Task<string> GetItemType(string url, string listName)
        {
            Newtonsoft.Json.Linq.JToken json = await SPRestGetItemsJson(url, listName);
            if (json != null && json["d"]["results"].HasValues)
            {
                string itemType = json["d"]["results"][0]["__metadata"]["type"].ToString();
                return itemType;
            }
            else return null;
        }

        private static async Task<int> SPRestCreateItem(string url, string listName, Tuple<string, object>[] data, string listItemType = "SP.Data.ListListItem")
        {
            int id = -1;
            try
            {
                System.Net.Http.HttpClientHandler clientHandler = new System.Net.Http.HttpClientHandler()
                {
                    UseDefaultCredentials = true
                };
                // Or define network credentials
                //clientHandler.Credentials = new System.Net.NetworkCredential("username", "pwd", "DOMAIN");

                System.Net.Http.HttpClient client = new System.Net.Http.HttpClient(clientHandler);

                string digest = GetSPFormDigest(url);

                client.BaseAddress = new System.Uri(url);
                client.DefaultRequestHeaders.Clear();
                client.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");
                client.DefaultRequestHeaders.Add("X-RequestDigest", digest);
                client.DefaultRequestHeaders.Add("X-HTTP-Method", "POST");

                System.Net.Http.StringContent content = new System.Net.Http.StringContent(BuildPostedItemData(listItemType, data));
                //System.Net.Http.StringContent content = new System.Net.Http.StringContent(" { '__metadata': { 'type': '" + listItemType + "' }, 'Title': 'Test: " + DateTime.Now.Ticks.ToString() + "', 'AssignedToId': 16, 'DueDate': '" + DateTime.Today.ToString("yyyy-MM-dd") + "' }");
                // If AssignedTo is multi-value lookup field:
                //System.Net.Http.StringContent content = new System.Net.Http.StringContent(" { '__metadata': { 'type': 'SP.Data.TasksListItem' }, 'Title': 'Test: " + DateTime.Now.Ticks.ToString() + "', 'AssignedToId': { 'results': [16] }, 'DueDate': '" + DateTime.Today.ToString("yyyy-MM-dd") + "' }");

                content.Headers.ContentType = System.Net.Http.Headers.MediaTypeHeaderValue.Parse("application/json;odata=verbose");

                // UrlEncode for plain english characters is not required, but it is for for non-english characters and various symbols
                string encodedListName = System.Web.HttpUtility.UrlEncode(listName);

                System.Net.Http.HttpResponseMessage response = await client.PostAsync("_api/web/lists/GetByTitle('" + encodedListName + "')/items", content);
                string spRequestId = response.Headers.GetValues("request-id").First(); // Correlation ID in ULS logs

                if (response.IsSuccessStatusCode)
                {
                    string text = await response.Content.ReadAsStringAsync();

                    Newtonsoft.Json.Linq.JToken parsedJson = Newtonsoft.Json.Linq.JToken.Parse(text);
                    string id_text = parsedJson["d"]["Id"].ToString();
                    id = int.Parse(id_text);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            return id;
        }

        private static async Task<bool> SPRestUpdateItem(string url, string listName, int id, Tuple<string, object>[] data, string listItemType = "SP.Data.ListListItem")
        {
            bool ret = false;
            try
            {
                System.Net.Http.HttpClientHandler clientHandler = new System.Net.Http.HttpClientHandler()
                {
                    UseDefaultCredentials = true
                };
                // Or define network credentials
                //clientHandler.Credentials = new System.Net.NetworkCredential("username", "pwd", "DOMAIN");

                System.Net.Http.HttpClient client = new System.Net.Http.HttpClient(clientHandler);

                string digest = GetSPFormDigest(url);

                client.BaseAddress = new System.Uri(url);
                client.DefaultRequestHeaders.Clear();
                client.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");
                client.DefaultRequestHeaders.Add("X-RequestDigest", digest);
                client.DefaultRequestHeaders.Add("IF-MATCH", "*");
                client.DefaultRequestHeaders.Add("X-HTTP-Method", "MERGE");

                System.Net.Http.StringContent content = new System.Net.Http.StringContent(BuildPostedItemData(listItemType, data));

                content.Headers.ContentType = System.Net.Http.Headers.MediaTypeHeaderValue.Parse("application/json;odata=verbose");

                // UrlEncode for plain english characters is not required, but it is for for non-english characters and various symbols
                string encodedListName = System.Web.HttpUtility.UrlEncode(listName);

                System.Net.Http.HttpResponseMessage response = await client.PostAsync("_api/web/lists/GetByTitle('" + encodedListName + "')/items(" + id.ToString() + ")", content);
                string spRequestId = response.Headers.GetValues("request-id").First(); // Correlation ID in ULS logs

                if (response.IsSuccessStatusCode)
                {
                    string text = await response.Content.ReadAsStringAsync();

                    ret = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            return ret;
        }

        private static string BuildPostedItemData(string listItemType, Tuple<string, object>[] data)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("{ \"__metadata\": { \"type\": \"" + listItemType + "\" }");

            for (int i = 0; i < data.Length; i++)
            {
                string v = Newtonsoft.Json.JsonConvert.SerializeObject(data[i].Item2);
                if (data[i].Item2 is Array)
                {
                    sb.AppendFormat(", \"{0}\":  {{ 'results': {1} }}", data[i].Item1, v);
                }
                else
                    sb.AppendFormat(", \"{0}\": {1}", data[i].Item1, v);
            }
            sb.Append(" }");

            return sb.ToString();
        }

        private static async Task<string> SPRestExecuteCamlQuery(string url, string listName, string query = "<View><Query></Query><RowLimit>1</RowLimit></View>")
        {
            try
            {
                System.Net.Http.HttpClientHandler clientHandler = new System.Net.Http.HttpClientHandler()
                {
                    UseDefaultCredentials = true
                };
                // Or define network credentials
                //clientHandler.Credentials = new System.Net.NetworkCredential("username", "pwd", "DOMAIN");

                System.Net.Http.HttpClient client = new System.Net.Http.HttpClient(clientHandler);

                string digest = GetSPFormDigest(url);

                client.BaseAddress = new System.Uri(url);
                client.DefaultRequestHeaders.Clear();
                client.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");
                client.DefaultRequestHeaders.Add("X-RequestDigest", digest);
                client.DefaultRequestHeaders.Add("X-HTTP-Method", "POST");

                string contentPost = " { \"query\" : { \"__metadata\": { \"type\": \"SP.CamlQuery\" }, \"ViewXml\": \"" + query + "\" } }";

                System.Net.Http.StringContent content = new System.Net.Http.StringContent(contentPost);

                content.Headers.ContentType = System.Net.Http.Headers.MediaTypeHeaderValue.Parse("application/json;odata=verbose");

                // UrlEncode for plain english characters is not required, but it is for for non-english characters and various symbols
                string encodedListName = System.Web.HttpUtility.UrlEncode(listName);

                System.Net.Http.HttpResponseMessage response = await client.PostAsync("_api/web/lists/GetByTitle('" + encodedListName + "')/GetItems", content);
                string spRequestId = response.Headers.GetValues("request-id").First(); // Correlation ID in ULS logs

                if (response.IsSuccessStatusCode)
                {
                    string text = await response.Content.ReadAsStringAsync();

                    Newtonsoft.Json.Linq.JToken parsedJson = Newtonsoft.Json.Linq.JToken.Parse(text);

                    return parsedJson.ToString(Newtonsoft.Json.Formatting.Indented);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            return null;
        }
    }
}

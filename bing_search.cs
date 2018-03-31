using System;
using System.Net.Http.Headers;
using System.Text;
using System.Net.Http;
using System.Web;
using System.Collections.Generic;
using System.Net;
using System.IO;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Linq;
 
namespace BingAPI
{
    class Program
    {
        // **********************************************
        // *** Update or verify the following values. ***
        // **********************************************
 
        // Replace the accessKey string value with your valid access key.
        const string accessKey = "27688d58abf44f8a800f6db10371e9c7";
 
        // Verify the endpoint URI.  At this writing, only one endpoint is used for Bing
        // search APIs.  In the future, regional endpoints may be available.  If you
        // encounter unexpected authorization errors, double-check this value against
        // the endpoint for your Bing Web search instance in your Azure dashboard.
        const string uriBase = "https://api.cognitive.microsoft.com/bing/v7.0/search";
 
        const string searchTerm = "Spotify";
 
        // Used to return search results including relevant headers
        struct SearchResult
        {
            public String jsonResult;
            public Dictionary<String, String> relevantHeaders;
        }
 
        [STAThread]
        static void Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            if (accessKey.Length == 32)
            {
                Console.WriteLine("Select a file");
                OpenFileDialog fl = new OpenFileDialog();
                fl.ShowDialog();
                string FileName = fl.FileName;
                OleDbConnection conn = new OleDbConnection
                       ("Provider=Microsoft.Jet.OleDb.4.0; Data Source = " +
                         Path.GetDirectoryName(FileName) +
                         "; Extended Properties = \"Text;HDR=YES;FMT=Delimited\"");
 
                conn.Open();
 
                OleDbDataAdapter adapter = new OleDbDataAdapter
                       ("SELECT * FROM " + Path.GetFileName(FileName), conn);
 
                DataSet ds = new DataSet("Temp");
                adapter.Fill(ds);
 
                conn.Close();
 
                DataTable outputdt = ds.Tables[0].Copy();
                outputdt.Columns.Add("URL");
                outputdt.Columns.Add("rawsearchoutput");
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    try
                    {
                        if (ds.Tables[0].Rows[i][0].ToString().Trim() == "") continue;
                        Console.WriteLine("Searching the Web for: " + ds.Tables[0].Rows[i][0].ToString());
 
                        SearchResult result = BingWebSearch(ds.Tables[0].Rows[i][0].ToString());
 
                        //Console.WriteLine("\nRelevant HTTP Headers:\n");
                        //foreach (var header in result.relevantHeaders) Console.WriteLine(header.Key + ": " + header.Value);
 
                        //Console.WriteLine("\nJSON Response:\n");
                        //Console.WriteLine(JsonPrettyPrint(result.jsonResult));
 
                        JObject json = JObject.Parse(result.jsonResult);
                        JArray URLs = (JArray)json["webPages"]["value"];
                        outputdt.Rows[i]["rawsearchoutput"] = result.jsonResult;
                        outputdt.Rows[i]["URL"] = "";
                        for (int j = 0; j < URLs.Count; j++)
                        {
                            outputdt.Rows[i]["URL"] += URLs[j]["url"].ToString() + "|";
                        }
                        System.Threading.Thread.Sleep(300);
 
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }
                StringBuilder sb = new StringBuilder();
 
                IEnumerable<string> columnNames = outputdt.Columns.Cast<DataColumn>().
                                                  Select(column => column.ColumnName);
                sb.AppendLine(string.Join(",", columnNames));
                foreach (DataRow row in outputdt.Rows)
                {
                    IEnumerable<string> fields = row.ItemArray.Select(field =>
                      string.Concat("\"", field.ToString().Replace("\"", "\"\""), "\""));
                    sb.AppendLine(string.Join(",", fields));
                }
                File.WriteAllText("test.csv", sb.ToString());
            }
            else
            {
                Console.WriteLine("Invalid Bing Search API subscription key!");
                Console.WriteLine("Please paste yours into the source code.");
            }
 
            Console.Write("\nPress Enter to exit ");
            Console.ReadLine();
        }
 
        /// <summary>
        /// Performs a Bing Web search and return the results as a SearchResult.
        /// </summary>
        static SearchResult BingWebSearch(string searchQuery)
        {
            // Construct the URI of the search request
            var uriQuery = uriBase + "?q=" + Uri.EscapeDataString(searchQuery);
 
            // Perform the Web request and get the response
            WebRequest request = HttpWebRequest.Create(uriQuery);
            request.Headers["Ocp-Apim-Subscription-Key"] = accessKey;
            HttpWebResponse response = (HttpWebResponse)request.GetResponseAsync().Result;
            string json = new StreamReader(response.GetResponseStream()).ReadToEnd();
 
            // Create result object for return
            var searchResult = new SearchResult()
            {
                jsonResult = json,
                relevantHeaders = new Dictionary<String, String>()
            };
 
            // Extract Bing HTTP headers
            foreach (String header in response.Headers)
            {
                if (header.StartsWith("BingAPIs-") || header.StartsWith("X-MSEdge-"))
                    searchResult.relevantHeaders[header] = response.Headers[header];
            }
 
            return searchResult;
        }
 
        /// <summary>
        /// Formats the given JSON string by adding line breaks and indents.
        /// </summary>
        /// <param name="json">The raw JSON string to format.</param>
        /// <returns>The formatted JSON string.</returns>
        static string JsonPrettyPrint(string json)
        {
            if (string.IsNullOrEmpty(json))
                return string.Empty;
 
            json = json.Replace(Environment.NewLine, "").Replace("\t", "");
 
            StringBuilder sb = new StringBuilder();
            bool quote = false;
            bool ignore = false;
            char last = ' ';
            int offset = 0;
            int indentLength = 2;
 
            foreach (char ch in json)
            {
                switch (ch)
                {
                    case '"':
                        if (!ignore) quote = !quote;
                        break;
                    case '\\':
                        if (quote && last != '\\') ignore = true;
                        break;
                }
 
                if (quote)
                {
                    sb.Append(ch);
                    if (last == '\\' && ignore) ignore = false;
                }
                else
                {
                    switch (ch)
                    {
                        case '{':
                        case '[':
                            sb.Append(ch);
                            sb.Append(Environment.NewLine);
                            sb.Append(new string(' ', ++offset * indentLength));
                            break;
                        case '}':
                        case ']':
                            sb.Append(Environment.NewLine);
                            sb.Append(new string(' ', --offset * indentLength));
                            sb.Append(ch);
                            break;
                        case ',':
                            sb.Append(ch);
                            sb.Append(Environment.NewLine);
                            sb.Append(new string(' ', offset * indentLength));
                            break;
                        case ':':
                            sb.Append(ch);
                            sb.Append(' ');
                            break;
                        default:
                            if (quote || ch != ' ') sb.Append(ch);
                            break;
                    }
                }
                last = ch;
            }
 
            return sb.ToString().Trim();
        }
 
    }
}

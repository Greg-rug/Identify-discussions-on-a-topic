using System.ComponentModel;
using System.Collections;
using System.Security.Cryptography;
using System.Security.AccessControl;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Linq;
using System.Web;
using Newtonsoft.Json;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace ConsoleApplication3
{
    class Program
    {
        /// The client information used to get the OAuth Access Token from the server.
        static string clientId     = "h6Zo1OcV8Oo-nxkB-6qUnlpt6xkgMk4P";
        static string clientSecret = "etoCDc38q2QCZgni9THXakFwWUEfKB0xsiwhQHZVlkDXb0xCxUBfi09J1AKM8Ub3";

 
        // The server base address
        static string baseUrl      = "https://api.codeproject.com/";
 
        // this will hold the Access Token returned from the server.
        static string accessToken  = null;
		
        static void Main(string[] args)
        {
            Console.WriteLine("Starting ...");
            DoIt().Wait();
            Console.ReadLine();
        }

        private static readonly JsonSerializerSettings _options = new() { NullValueHandling = NullValueHandling.Ignore };

        /// <summary>
        /// This method takes an object variable, which it then writes as a JSON string on the file designated
        /// with the value stored in the filename string
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="fileName"></param>//  
        public static void SimpleWrite(object obj, string fileName)
        {
             var jsonString = JsonConvert.SerializeObject(obj, _options);
             File.AppendAllText(fileName, jsonString); 
             
        }
 
        /// <summary>
        /// This method does all the work to get an Access Token and read the first page of
        /// Articles from the server.
        /// </summary>
        /// <returns></returns>
        private static async Task<int> DoIt()
        {
            // Get the Access Token.
            accessToken  = await GetAccessToken();
            Console.WriteLine( accessToken != null ? "Got Token" : "No Token found");
			
			// Get the Forum Messages
            Console.WriteLine();
            Console.WriteLine("------ New Thread Messages ------");
            int i = 1;
            dynamic[] responses = new dynamic[102];
            //This while loop goes through the 102 pages of messages from the forum and stores them in the 
            //dynamic object response at the index i-1, where i starts at 1
            while(i<=102){
                dynamic response = await GetThreadMessages(3304,"Messages",i);
                if (response.items != null){
                    responses[i-1]=response;
                    var messages = (dynamic)response.items;
                    foreach(dynamic message in messages)
                    Console.WriteLine("Title: {0}", message.title);      
                 }
                i++; 
            }
            // This writes the stored responses to the Message.json file
            SimpleWrite(responses,"Messages.json");
            // Create a Workbook object
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Read JSON File
            string jsonInput = File.ReadAllText("Messages.json");            

            // Set JsonLayoutOptions
            JsonLayoutOptions options = new JsonLayoutOptions(); 
            options.ArrayAsTable = true;

            // Import JSON Data
            JsonUtility.ImportData(jsonInput, worksheet.Cells, 0, 0, options);

            // Save Excel file
            // Remove the first row of the Excel sheet since it causes issues with the model training 
            workbook.Save("JsonData.xlsx");
        
            return 0;
        }
 
        /// <summary>
        /// This method uses the OAuth Client Credentials Flow to get an Access Token to provide
        /// Authorization to the APIs.
        /// </summary>
        /// <returns></returns>
        private static async Task<string> GetAccessToken()
        {
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(baseUrl);
 
                // We want the response to be JSON.
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
 
                // Build up the data to POST.
                List<KeyValuePair<string, string>> postData = new List<KeyValuePair<string, string>>();
                postData.Add(new KeyValuePair<string, string>("grant_type",    "client_credentials"));
                postData.Add(new KeyValuePair<string, string>("client_id",     clientId));
                postData.Add(new KeyValuePair<string, string>("client_secret", clientSecret));
 
                FormUrlEncodedContent content = new FormUrlEncodedContent(postData);
 
                // Post to the Server and parse the response.
                HttpResponseMessage response = await client.PostAsync("Token", content);
                string jsonString            = await response.Content.ReadAsStringAsync();
                object responseData          = JsonConvert.DeserializeObject(jsonString);
 
                // return the Access Token.
                return ((dynamic)responseData).access_token;
            }
        }

		/// <summary>
        /// Gets the thread messages on a page.
        /// </summary>
        /// <param name="page">The page to get.</param>
        /// <param name="threadId">The identifiaction tag of the thread.</param>
        /// <returns>The messages on a given thread.</returns>
        private static async Task<dynamic> GetThreadMessages(int forumId, string mode, int page)
        {
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(baseUrl);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
 
                // Add the Authorization header with the AccessToken.
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);
 
                // create the URL string.
                string url = string.Format("v1/Forum/{0}/{1}?page={2}",forumId,mode,page);
 
                // make the request
                HttpResponseMessage response = await client.GetAsync(url);
 
                // parse the response and return the data.
                string jsonString = await response.Content.ReadAsStringAsync();
                object responseData = JsonConvert.DeserializeObject(jsonString);
                return (dynamic)responseData;
            }
        }
    }
}
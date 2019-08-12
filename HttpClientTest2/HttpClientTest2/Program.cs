using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace HttpClientTest2
{
    class Program
    {
        static void Main(string[] args)
        {
            CallWebAPIAsync().Wait();

        }
        static async Task CallWebAPIAsync()
        {
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri("https://jsonplaceholder.typicode.com/todos/1");
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                //GET Method  
                HttpResponseMessage response = await client.GetAsync("https://jsonplaceholder.typicode.com/todos/1");

                if (response.IsSuccessStatusCode)
                {
                    //
                    string responseBody = await response.Content.ReadAsStringAsync();
                    Console.WriteLine("Success" + responseBody);
                    JObject parsed = JObject.Parse(responseBody);
                    foreach (var pair in parsed)
                    {
                        Console.WriteLine("{0}: {1}", pair.Key, pair.Value);
                    }
                    Console.ReadKey();
                }
                else
                {
                    Console.WriteLine("Internal server Error");
                    Console.ReadKey();
                }
            }
        }
    }
}

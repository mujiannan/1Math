using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Net.Http;
using System.Collections;
namespace _1Math
{
    class Text
    {
        private string _content;
        public string Content
        {
            set
            {
                _content = value;
            }
        }
        public async Task<string> ToEnglishAsync()
        {
            string host = "https://api.cognitive.microsofttranslator.com";
            string route = "/translate?api-version=3.0&to=en";
            string subscriptionKey = "b72b53dd8edd435db758795dc00894d2";
            System.Object[] body = new System.Object[] { new { Text =_content } };
            string requestBody = JsonConvert.SerializeObject(body);
            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage
            {
                Method = HttpMethod.Post,
                RequestUri = new Uri(host + route),
                Content = new StringContent(requestBody, Encoding.UTF8, "application/json")
            };
            request.Headers.Add("Ocp-Apim-Subscription-Key", subscriptionKey);
            var response = await client.SendAsync(request);
            string responseBody = await response.Content.ReadAsStringAsync();
            Newtonsoft.Json.Linq.JArray result = JsonConvert.DeserializeObject<Newtonsoft.Json.Linq.JArray>(responseBody);
            string translation = result[0]["translations"][0]["text"].ToString();
            client.Dispose();
            request.Dispose();
            return translation;
        }
    }
}

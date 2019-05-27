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
        private List<object> _content=new List<object>();
        private List<string> _result = new List<string>();
        public void AddContent(string newContent)
        {
            _content.Add(new {Text = newContent});
        }
        public async Task<List<string>> ToEnglishAsync()
        {
            string host = "https://api.cognitive.microsofttranslator.com";
            string route = "/translate?api-version=3.0&to=en";
            string subscriptionKey = Properties.Resources.AzureCognitiveKey;
            string requestBody = JsonConvert.SerializeObject(_content);
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
            try
            {
                Newtonsoft.Json.Linq.JArray result = JsonConvert.DeserializeObject<Newtonsoft.Json.Linq.JArray>(responseBody);
                for (int i = 0; i < result.Count; i++)
                {
                    _result.Add(result[i]["translations"][0]["text"].ToString());
                }
            }
            catch (Exception)
            {
                throw;
            }
            client.Dispose();
            request.Dispose();
            return _result;
        }
    }
}

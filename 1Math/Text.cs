using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net.Http;
using System.Collections;
namespace _1Math
{
    class Translator
    {
        private static string _host = "https://api.cognitive.microsofttranslator.com";
        private static string _subscriptionKey = Properties.Resources.AzureCognitiveKey;
        private List<object> _contents=new List<object>();
        private List<string> _result = new List<string>();
        public void AddContent(string newContent)
        {
            _contents.Add(new {Text = newContent});
        }
        public List<object> Contents
        {
            set
            {
                _contents = value;
            }
        }
        private static string _acceptLanguages=string.Empty;//不要在AcceptLanguages之外使用这个字段
        private static string AcceptLanguages
        {
            get
            {
                if (_acceptLanguages==string.Empty)
                {
                    string route = "/languages?api-version=3.0";
                    using (HttpClient httpClient = new HttpClient())
                    {
                        using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Get, new Uri(_host + route)))
                        {
                            httpRequestMessage.Headers.Add("Accept-Language", "en");
                            _acceptLanguages=httpClient.SendAsync(httpRequestMessage).Result.Content.ReadAsStringAsync().Result;
                        }
                    }
                    
                }
                return _acceptLanguages;
            }
        }
        private static Dictionary<string, Dictionary<string, string>> _translatableLanguages=null;
        public static Dictionary<string, Dictionary<string, string>> TranslatableLanguages//Dictionary<code, Dictionary<{name,nativeName,dir}, UnkownString>>//其中dir表示文字排列方向
        {
            get
            {
                if (_translatableLanguages==null)
                {
                    string translation = JsonConvert.DeserializeObject<JObject>(AcceptLanguages)["translation"].ToString();
                    _translatableLanguages = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, string>>>(translation);//如果字段为null，则先获取再设置，最后总是返回该字段
                }
                return _translatableLanguages;
            }
        }
        public static string Test
        {
            get
            {
                return TranslatableLanguages["zh-Hans"]["nativeName"];
            }
        }
        public async Task<List<string>> TranslateAsync(string toLanguageCode)
        {
            if (!TranslatableLanguages.ContainsKey(toLanguageCode))
            {
                throw new Exception("UnexpectedLanguageCode");
            }
            string route = "/translate?api-version=3.0&to="+ toLanguageCode;
            string requestBody = JsonConvert.SerializeObject(_contents);
            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage
            {
                Method = HttpMethod.Post,
                RequestUri = new Uri(_host + route),
                Content = new StringContent(requestBody, Encoding.UTF8, "application/json")
            };
            request.Headers.Add("Ocp-Apim-Subscription-Key", _subscriptionKey);
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

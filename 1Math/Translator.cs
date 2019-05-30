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
    public class Translator
    {
        private string _host;
        private string _subscriptionKey;
        public Translator(string baseUrl, string key)//baseUrl should be string like "api.cognitive.microsofttranslator.com", determined by the Region of your AzureCognitiveService
        {
            _host = "https://" + baseUrl;
            _subscriptionKey = key;
        }
        private List<List<object>> _perfectContentForTranslation=new List<List<object>> { new List<object>()};//this whole List<List<string>> must be limited to 5000 characters

        private int _CharactersCount = 0;

        //SetThreadsCount
        private int _limitedThreadsCount=2;
        public int LimitedThreadsCount
        {
            get => _limitedThreadsCount;
            set
            {
                if (value>0)
                {
                    _limitedThreadsCount = value;
                }
                else
                {
                    throw new Exception("InvalidLimitedThreadsCount");
                }
            }
        }

        //SetContensForTranslation
        private const int _limitedCharactersCount = 5000;
        private const int _limitedItemsCount = 100;
        public void AddContent(string newContent)
        {
            if (newContent.Length > _limitedCharactersCount)
            {
                throw new Exception("OutOfCharactersCount");
            }
            _perfectContentForTranslation[_perfectContentForTranslation.Count - 1].Add(new { Text=newContent });//先验证是否超过100毫无意义，因为就算没超过100，也可能在添加后导致长度达到5000……
            _CharactersCount += newContent.Length;
            if (_perfectContentForTranslation[_perfectContentForTranslation.Count - 1].Count>_limitedItemsCount|| _CharactersCount > _limitedCharactersCount)//each List<string> must be limited to one hundred item
            {
                _perfectContentForTranslation[_perfectContentForTranslation.Count - 1].RemoveAt(_perfectContentForTranslation[_perfectContentForTranslation.Count - 1].Count - 1);//that's complex, but only one line
                _perfectContentForTranslation.Add(new List<object>());
                _perfectContentForTranslation[_perfectContentForTranslation.Count - 1].Add(new{ Text = newContent});
                _CharactersCount = newContent.Length;//a new List<string>, a new _CharactersCount
            }
        }
        public List<string> Contents//user naturally set the contents, ignore all limits
        {
            set
            {
                for (int i = 0; i < value.Count; i++)
                {
                    AddContent(value[i]);
                }
            }
        }

        //jsonAcceptLanguage
        private string _acceptLanguages=string.Empty;//不要在AcceptLanguages之外使用这个字段
        private string AcceptLanguages
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

        //GetTranslableLanguage
        public struct Language
        {
            public string name, nativeName, dir;
        }
        private static Dictionary<string, Language> _translatableLanguages=null;
        public Dictionary<string, Language> TranslatableLanguages//code as key, struct Language as value
        {
            get
            {
                if (_translatableLanguages==null)
                {
                    string translation = JsonConvert.DeserializeObject<JObject>(AcceptLanguages)["translation"].ToString();//extract "translation" from jsonAcceptLanguage
                    _translatableLanguages = JsonConvert.DeserializeObject<Dictionary<string, Language>>(translation);//如果字段为null，则先获取再设置，最后总是返回该字段
                }
                return _translatableLanguages;
            }
        }
        
        //This is a useless TranslatingProgressReporter(If you translate more than one million characters at one time, maybe it's useful)
        public class TranslatingEventArgs : EventArgs
        {
            public TranslatingEventArgs(double newProgress)
            {
                if (newProgress >= 0 & newProgress <= 1)
                {
                    NewProgress = newProgress;
                }
                else
                {
                    throw new Exception("InvalidNewProgress");
                }
            }
            public double NewProgress { get; }
        }
        public delegate void DTranslating(object sender, TranslatingEventArgs e);
        public event DTranslating ProgressChange;
        private void ChangeProgress(double newProgress)
        {
            DTranslating dTranslating = ProgressChange;
            if (dTranslating!=null)
            {
                ProgressChange(this, new TranslatingEventArgs(newProgress));
            }
        }

        //Main method for translation
        public async Task<List<string>> TranslateAsync(string toLanguageCode)
        {
            Task<List<string>>[] tasks = new Task<List<string>>[_perfectContentForTranslation.Count];
            for (int i = 0; i < _perfectContentForTranslation.Count; i++)
            {
                tasks[i] =TranslateAsync(_perfectContentForTranslation[i], toLanguageCode);
            }
            List<string> results = new List<string>();
            for (int i = 0; i < tasks.Length; i++)
            {
                results.AddRange(await tasks[i]);
                ChangeProgress((double)(i+1) / tasks.Length);
            }
            return results;
        }
        //TranslateAsync will call this method
        private async Task<List<string>> TranslateAsync(List<object> contentsForTranslation,string toLanguageCode)
        {
            if (!TranslatableLanguages.ContainsKey(toLanguageCode))
            {
                throw new Exception("UnexpectedLanguageCode");
            }
            string route = "/translate?api-version=3.0&to="+ toLanguageCode;
            string requestBody = JsonConvert.SerializeObject(contentsForTranslation);
            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage
            {
                Method = HttpMethod.Post,
                RequestUri = new Uri(_host + route),
                Content = new StringContent(requestBody, Encoding.UTF8, "application/json")
            };
            request.Headers.Add("Ocp-Apim-Subscription-Key", _subscriptionKey);
            HttpResponseMessage response = await client.SendAsync(request);
            string responseBody = await response.Content.ReadAsStringAsync();
            List<string> results = new List<string>();
            try
            {
                Newtonsoft.Json.Linq.JArray result = JsonConvert.DeserializeObject<Newtonsoft.Json.Linq.JArray>(responseBody);
                for (int i = 0; i < result.Count; i++)
                {
                    results.Add(result[i]["translations"][0]["text"].ToString());
                }
            }
            catch (Exception)
            {
                throw;
            }
            client.Dispose();
            request.Dispose();
            return results;
        }
    }
}

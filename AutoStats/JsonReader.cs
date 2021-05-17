using System.IO;
using System.Net.Http;
using Newtonsoft.Json;

namespace AutoStats
{
    public class JsonReader
    {
        public dynamic Content { get; }

        public JsonReader(HttpResponseMessage httpResponseMessage)
        {
            var task = httpResponseMessage.Content.ReadAsStringAsync();
            task.Wait();
            var responseBody = task.Result;
            Content = JsonConvert.DeserializeObject(responseBody);
        }
        
        public JsonReader(string path)
        {
            using var streamReader = new StreamReader(path);
            Content = JsonConvert.DeserializeObject(streamReader.ReadToEnd());
        }
    }
}
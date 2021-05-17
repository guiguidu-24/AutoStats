using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace AutoStats
{
    //TODO The doc of the strava class
    //TODO Make the Strava class non-static
    public static class Strava
    {
        private static HttpClient httpClient = new HttpClient();
        private static string _accesToken;
        private static int _accesTokenExpiration;
        
        //Todo Make a synchronous ActivitiesAfterAsync
        public static async Task<dynamic> ActivitiesAfterAsync(int dateTime)
        {
            if (UnixTimestamp.ToEpochTime(DateTime.Now) >= _accesTokenExpiration)
                await NewAccesTokenAsync();
            
            var uri = new Uri($"https://www.strava.com/api/v3/athlete/activities?after={dateTime}");
            httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", _accesToken);
            var httpResponse = await httpClient.GetAsync(uri);
            httpResponse.EnsureSuccessStatusCode();

            return new JsonReader(httpResponse).Content;
        }
        
        public static dynamic ActivitiesAfter(int dateTime)
        {
            if (UnixTimestamp.ToEpochTime(DateTime.Now) >= _accesTokenExpiration)
            {
                var newAccTkTask = NewAccesTokenAsync();
                newAccTkTask.Wait();
            }
            
            var uri = new Uri($"https://www.strava.com/api/v3/athlete/activities?after={dateTime}");
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accesToken);
            
            var httpResponseTask = httpClient.GetAsync(uri);
            httpResponseTask.Wait();
            var httpResponse = httpResponseTask.Result;
            httpResponse.EnsureSuccessStatusCode();

            return new JsonReader(httpResponse).Content;
        }

        
        private static async Task NewAccesTokenAsync()
        {
            var jsonSettings = new JsonReader(
                    @"C:\Users\guill\Programmation\dotNET_doc\projets\Console\AutoStats\AutoStats\appsettings.json").Content;

            var parameters = new Dictionary<string, string>()
            {
                {"client_id", (string) jsonSettings["Config"]["StravaApi"]["ClientId"]},
                {"client_secret", (string) jsonSettings["Config"]["StravaApi"]["ClientSecret"]},
                {"refresh_token", (string) jsonSettings["Config"]["StravaApi"]["RefreshToken"]},
                {"grant_type", "refresh_token"}
            };
            
            var uri = new Uri($"https://www.strava.com/api/v3/oauth/token");
            httpClient.DefaultRequestHeaders.Clear();
            
            var httpResponse = await httpClient.PostAsync(uri, new FormUrlEncodedContent(parameters));
            httpResponse.EnsureSuccessStatusCode();
            
            var jsonResponse = new JsonReader(httpResponse).Content;

            _accesToken = (string)jsonResponse["access_token"];
            _accesTokenExpiration = (int)jsonResponse["expires_at"];
        }
    }
}
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace AutoStats
{
    public static class Strava
    {
        private static HttpClient httpClient = new HttpClient();
        private static string _accesToken;
        private static int _accesTokenExpiration;

        public static async Task<dynamic> ActivitiesAfter(int dateTime)
        {
            if (UnixTimestamp.ToEpochTime(DateTime.Now) >= _accesTokenExpiration)
                await NewAccesToken();
            
            var uri = new Uri($"https://www.strava.com/api/v3/athlete/activities?after={dateTime}");
            httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", _accesToken);
            var httpResponse = await httpClient.GetAsync(uri);
            httpResponse.EnsureSuccessStatusCode();

            return new JsonReader(httpResponse).Content;
        }
        
        private static async Task NewAccesToken()
        {
            dynamic jsonSettings =
                new JsonReader(
                        @"C:\Users\guill\Programmation\dotNET_doc\projets\Console\AutoStats\AutoStats\appsettings.json")
                    .Content;

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
            
            dynamic jsonResponse = new JsonReader(httpResponse).Content;

            _accesToken = (string)jsonResponse["access_token"];
            _accesTokenExpiration = (int)jsonResponse["expires_at"];
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace AutoStats
{
    /// <summary>
    /// A class handling the strava api to do specific requests
    /// </summary>
    public class Strava
    {
        private HttpClient httpClient = new HttpClient();
        private readonly Dictionary<string, string> refreshData = new();
        
        private string accesToken;
        private int accesTokenExpiration;
        
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="client_id">The strava client id</param>
        /// <param name="client_secret">The strava client secret</param>
        /// <param name="refresh_token"> The strava refresh token (all reading)</param>
        public Strava(string client_id, string client_secret, string refresh_token)
        {
            refreshData.Add(nameof(client_id), client_id);
            refreshData.Add(nameof(client_secret), client_secret);
            refreshData.Add(nameof(refresh_token), refresh_token);
            
            NewAccesTokenAsync().Wait();
        }
        
        /// <summary>
        /// Get the activities from the strava api asynchronously
        /// </summary>
        /// <param name="dateTime">The date to see after</param>
        /// <returns>The task relative to the method</returns>
        public async Task<dynamic> ActivitiesAfterAsync(int dateTime)
        {
            if (UnixTimestamp.ToEpochTime(DateTime.Now) >= accesTokenExpiration)
                await NewAccesTokenAsync();
            
            var uri = new Uri($"https://www.strava.com/api/v3/athlete/activities?after={dateTime}&per_page=200");
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accesToken);
            
            var httpResponse = await httpClient.GetAsync(uri);
            httpResponse.EnsureSuccessStatusCode();

            return new JsonReader(httpResponse).Content;
        }
        
        /// <summary>
        /// Get the activities from the strava api
        /// </summary>
        /// <param name="dateTime">The date to see after</param>
        public dynamic ActivitiesAfter(int dateTime)
        {
            var activitiesTask = ActivitiesAfterAsync(dateTime);
            activitiesTask.Wait();

            return activitiesTask.Result;
        }

        /// <summary>
        /// Refresh the access token
        /// </summary>
        private async Task NewAccesTokenAsync()
        {
            var parameters = new Dictionary<string, string>()
            {
                {"grant_type", "refresh_token"}
            }.Union(refreshData).ToDictionary(pair => pair.Key, pair => pair.Value);;

            var uri = new Uri($"https://www.strava.com/api/v3/oauth/token");
            httpClient.DefaultRequestHeaders.Clear();
            
            var httpResponse = await httpClient.PostAsync(uri, new FormUrlEncodedContent(parameters));
            httpResponse.EnsureSuccessStatusCode();
            
            var jsonResponse = new JsonReader(httpResponse).Content;

            accesToken = (string)jsonResponse["access_token"];
            accesTokenExpiration = (int)jsonResponse["expires_at"];
        }
    }
}
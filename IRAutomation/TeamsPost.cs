using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace IRAutomation
{
    class TeamsPost
    {
        public string payload;

        public TeamsPost(string s) {

            payload = s;
            PostMessage();

        }


        
        public async void PostMessage()
        {
            string payloadJson = payload;
            var content = new StringContent(payloadJson);
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            var client = new HttpClient();
            Uri uri = new Uri("https://yourteams.webhook.here");
            await client.PostAsync(uri, content);
        }

    }

    
}

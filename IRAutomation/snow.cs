using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Configuration;
using System.IO;
using System.Net;
using System.Text;


namespace IRAutomation
{
    class snow
    {
        public static string CreateIncidentServiceNow(string _username, string _password,string _correlation_id, string _short_description, string _description, string _affected_user,string category)
        {
            try
            {

                string username = _username;
                string password = _password;

               
                string url = "your servicenow connection string goes here";

                var auth = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(username + ":" + password));

                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                request.Headers.Add("Authorization", auth);
                request.Method = "Post";

                using (var streamWriter = new StreamWriter(request.GetRequestStream()))
                {
                    string json = JsonConvert.SerializeObject(new
                    {
                        correlation_display = "Proofpoint API",
                        correlation_id = _correlation_id,
                        short_description = _short_description,
                        assignment_group = "Information Security",
                        impact = "3",
                        severity = "3",
                        u_detection_system = "Proofpoint",
                        description = _description,
                        affected_user = _affected_user
                    });

                    streamWriter.Write(json);
                }

                using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
                {
                    var res = new StreamReader(response.GetResponseStream()).ReadToEnd();

                    JObject joResponse = JObject.Parse(res.ToString());
                    JObject ojObject = (JObject)joResponse["result"];
                    string incNumber = ((JValue)ojObject.SelectToken("number")).Value.ToString();

                    return incNumber;
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }
}

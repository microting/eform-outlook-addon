using eFormShared;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace OutlookExchangeOnlineAPI
{
    public class OutlookExchangeOnlineAPIClient
    {
        string authority = "https://login.microsoftonline.com/REPLACE_ME_GUID/oauth2/token";
        string application_id = "";
        // Used to set the base API endpoint, e.g. "https://outlook.office.com/api/beta"
        public string ApiEndpoint { get; set; }
        public string AccessToken { get; set; }
        string serviceLocation;
        public Log log;


        public OutlookExchangeOnlineAPIClient(string serviceLocation)
        {
            // Set default endpoint
            log.LogStandard("Not Specified", "serviceLocation is set to " + serviceLocation);
            this.serviceLocation = serviceLocation;
            ApiEndpoint = "https://outlook.office.com/api/v2.0";
            AccessToken = GetAppToken(@"cert\cert.pfx", "123qweASDZXC");//the pfx file is encrypted with this password
        }

        private string GetAppToken(string certFile, string certPass)
        {
            string directory_id = File.ReadAllText(GetServiceLocation() + @"cert\directory_id.txt").Trim();
            application_id = File.ReadAllText(GetServiceLocation() + @"cert\application_id.txt").Trim();
            X509Certificate2 cert = new X509Certificate2(certFile, certPass, X509KeyStorageFlags.MachineKeySet);
            AuthenticationContext authContext = new AuthenticationContext(authority.Replace("REPLACE_ME_GUID", directory_id));
            ClientAssertionCertificate assertion = new ClientAssertionCertificate(application_id, cert);
            AuthenticationResult authResult = authContext.AcquireTokenAsync("https://outlook.office.com", assertion).Result;
            return authResult.AccessToken;
        }

        public HttpResponseMessage MakeApiCall(string method, string apiUrl, string userEmail, string payload, Dictionary<string, string> preferHeaders)
        {
            if (string.IsNullOrEmpty(AccessToken))
            {
                throw new ArgumentNullException("AccessToken", "You must supply an access token before making API calls.");
            }

            using (var httpClient = new HttpClient())
            {
                var request = new HttpRequestMessage(new HttpMethod(method), ApiEndpoint + apiUrl);

                // Headers
                // Add the access token in the Authorization header
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", AccessToken);
                // Add a user agent (best practice)
                request.Headers.UserAgent.Add(new System.Net.Http.Headers.ProductInfoHeaderValue("outlook-fetch", "1.0"));
                // Indicate that we want JSON response
                request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                // Set a unique ID on each request (best practice)
                request.Headers.Add("client-request-id", Guid.NewGuid().ToString());
                // Request that the unique ID also be included in the response (to make it easier to correlate request/response)
                request.Headers.Add("return-client-request-id", "true");
                // Set this header to optimize routing of request to appropriate server
                request.Headers.Add("X-AnchorMailbox", userEmail);

                if (preferHeaders != null)
                {
                    foreach (KeyValuePair<string, string> header in preferHeaders)
                    {
                        if (string.IsNullOrEmpty(header.Value))
                        {
                            // Some prefer headers only have a name, no value
                            request.Headers.Add("Prefer", header.Key);
                        }
                        else
                        {
                            request.Headers.Add("Prefer", string.Format("{0}=\"{1}\"", header.Key, header.Value));
                        }
                    }
                }

                // POST and PATCH should have a body
                if ((method.ToUpper() == "POST" || method.ToUpper() == "PATCH") &&
                    !string.IsNullOrEmpty(payload))
                {
                    request.Content = new StringContent(payload);
                    request.Content.Headers.ContentType.MediaType = "application/json";
                }

                var apiResult = httpClient.SendAsync(request).Result;
                return apiResult;
            }
        }

        public CalendarList GetCalendarList(string userEmail, string calendarName)
        {
            string requestUrl = String.Format("/users/{0}/calendars", userEmail);
            HttpResponseMessage result = MakeApiCall("GET", requestUrl, userEmail, null, null);
            string response = result.Content.ReadAsStringAsync().Result;
            return JsonConvert.DeserializeObject<CalendarList>(response);
        }

        public EventList GetCalendarItems(string userEmail, string calendarID, DateTime startDate, DateTime enddate)
        {
            EventList result = new EventList();
            result.value = new List<Event>();
            bool alldone = false;
            int skip = 0;
            string requestUrl;
            while (!alldone)
            {
                requestUrl = String.Format("/users/{0}/calendars/{1}/calendarview?startDateTime={2}&endDateTime={3}&$skip={4}", userEmail, calendarID, startDate, enddate, skip);
                HttpResponseMessage httpresult = MakeApiCall("GET", requestUrl, userEmail, null, null);
                string response = httpresult.Content.ReadAsStringAsync().Result;
                EventList curpage = JsonConvert.DeserializeObject<EventList>(response);
                if (curpage.value.Count >= 10) alldone = false;
                else alldone = true;
                foreach (Event curevent in curpage.value)
                {
                    result.value.Add(curevent);
                }
                skip += curpage.value.Count;
            }
            return result;
        }

        public Event UpdateEvent(string userEmail, string EventID, string update)
        {
            string requestUrl = String.Format("/users/{0}/events/{1}", userEmail, EventID);
            HttpResponseMessage result = MakeApiCall("PATCH", requestUrl, userEmail, update, null);
            string response = result.Content.ReadAsStringAsync().Result;
            Event UpdatedEvent = JsonConvert.DeserializeObject<Event>(response);
            return UpdatedEvent;

        }
        public Event CreateEvent(string userEmail, string CalendarID, string eventObject)
        {
            string requestUrl = String.Format("/users/{0}/calendars/{1}/events", userEmail, CalendarID);
            HttpResponseMessage httpresult = MakeApiCall("POST", requestUrl, userEmail, eventObject, null);
            string response = httpresult.Content.ReadAsStringAsync().Result;
            Event newEvent = JsonConvert.DeserializeObject<Event>(response);
            return newEvent;

        }
        public void DeleteEvent(string userEmail, string EventID)
        {
            string requestUrl = String.Format("/users/{0}/events/{1}", userEmail, EventID);
            HttpResponseMessage result = MakeApiCall("DELETE", requestUrl, userEmail, null, null);
        }

        private string GetServiceLocation()
        {
            if (serviceLocation != "")
                return serviceLocation;

            serviceLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            serviceLocation = Path.GetDirectoryName(serviceLocation) + "\\";
            //LogEvent("serviceLocation:'" + serviceLocation + "'");

            return serviceLocation;
        }
    }
}

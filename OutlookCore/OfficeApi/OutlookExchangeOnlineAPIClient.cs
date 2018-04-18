using eFormShared;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
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
        Tools t = new Tools();
        string certPath = @"cert\cert.pfx";
        string certPass = "123qweASDZXC";


        public OutlookExchangeOnlineAPIClient(string serviceLocation, Log logger)
        {
            // Set default endpoint
            log = logger;
            this.serviceLocation = serviceLocation;
            ApiEndpoint = "https://outlook.office.com/api/v2.0";
            AccessToken = GetAppToken(GetServiceLocation() + certPath, certPass); //the pfx file is encrypted with this password
        }

        public string GetAppToken(string certFile, string certPass)
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

                var apiResult = ExecuteQueryWithIncrementalRetry(request, 3, 30);
                //var apiResult = httpClient.SendAsync(request).Result;
                return apiResult;
            }
        }

        public HttpResponseMessage ExecuteQueryWithIncrementalRetry(HttpRequestMessage request, int retryCount, int delay)
        {
            int retryAttempts = 0;
            int backoffInteval = delay;
            if (retryCount <= 0)
                throw new ArgumentException("Provide a retryCount greater than zero.");
            if (delay <= 0)
                throw new ArgumentException("Provide a delay count greater than zero.");
            HttpResponseMessage result = null;

            while (retryAttempts < retryCount)
            {
                using (var httpClient = new HttpClient())
                {
                    try
                    {

                        log.LogEverything(t.GetMethodName("OutlookExchangeOnlineAPIClient"), "ExecuteQueryWithIncrementalRetry trying to call :" + request.RequestUri);
                        result = httpClient.SendAsync(request).Result;
                        if (!result.StatusCode.Equals(HttpStatusCode.OK))
                        {
                            if (!result.StatusCode.Equals(HttpStatusCode.Created))
                            {

                                if (result.StatusCode.Equals(HttpStatusCode.NoContent))
                                {
                                    log.LogEverything(t.GetMethodName("OutlookExchangeOnlineAPIClient"), "ExecuteQueryWithIncrementalRetry got result NoContent and result.Content is :" + result.Content);
                                    return result;
                                }
                                else
                                {
                                    if (result.StatusCode.Equals(HttpStatusCode.Unauthorized))
                                    {
                                        AccessToken = GetAppToken(GetServiceLocation() + certPath, certPass); //the pfx file is encrypted with this password
                                        log.LogEverything(t.GetMethodName("OutlookExchangeOnlineAPIClient"), "ApiClient.ExecuteQueryWithIncrementalRetry called and status code is Unauthorized so resetting retryAttempts ");
                                        retryAttempts = 0; 
                                    } else
                                    {
                                        if (result.StatusCode.Equals(HttpStatusCode.NotFound))
                                        {
                                            return result;
                                        }
                                        log.LogEverything(t.GetMethodName("OutlookExchangeOnlineAPIClient"), "ApiClient.ExecuteQueryWithIncrementalRetry called and status code is not OK or Created and backoffInteval is now " + backoffInteval.ToString() + " and retryAttempts is " + retryAttempts.ToString());
                                        log.LogEverything(t.GetMethodName("OutlookExchangeOnlineAPIClient"), "ApiClient.ExecuteQueryWithIncrementalRetry called and status code is : " + result.StatusCode.ToString());
                                        System.Threading.Thread.Sleep(backoffInteval * 1000);
                                        retryAttempts++;
                                        backoffInteval = backoffInteval * 2;
                                    }                                                                        
                                }

                            }
                            else
                            {
                                return result;
                            }
                        }
                        else
                        {
                            return result;
                        }
                    }
                    catch (Exception ex)
                    {
                        log.LogEverything(t.GetMethodName("OutlookExchangeOnlineAPIClient"), "ApiClient.ExecuteQueryWithIncrementalRetry throwed an Exception and backoffInteval is now " + backoffInteval.ToString() + " and retryAttempts is " + retryAttempts.ToString());
                        log.LogEverything(t.GetMethodName("OutlookExchangeOnlineAPIClient"), "ApiClient.ExecuteQueryWithIncrementalRetry the exeption is : " + ex.Message);
                        System.Threading.Thread.Sleep(backoffInteval * 1000);
                        retryAttempts++;
                        backoffInteval = backoffInteval * 2;
                    }

                }
            }
            return result;
        }

        public CalendarList GetCalendarList(string userEmail, string calendarName)
        {
            log.LogEverything(t.GetMethodName("OutlookExchangeOnlineAPIClient"), "ApiClient.GetCalendarList called");
            if (string.IsNullOrEmpty(userEmail))
                throw new ArgumentNullException("userEmail cannot be null or empty");
            if (string.IsNullOrEmpty(calendarName))
                throw new ArgumentNullException("calendarName cannot be null or empty");

            string requestUrl = String.Format("/users/{0}/calendars", userEmail);
            HttpResponseMessage result = MakeApiCall("GET", requestUrl, userEmail, null, null);
            string response = result.Content.ReadAsStringAsync().Result;

            log.LogEverything(t.GetMethodName("OutlookExchangeOnlineAPIClient"), "ApiClient.GetCalendarList response is : " + response);
            if (response.Contains("odata"))
            {
                return JsonConvert.DeserializeObject<CalendarList>(response);
            }
            else
            {
                return null;
            }

        }

        public EventList GetCalendarItems(string userEmail, string calendarID, DateTime startDate, DateTime enddate)
        {
            log.LogEverything(t.GetMethodName("OutlookExchangeOnlineAPIClient"), "ApiClient.GetCalendarItems called");

            if (string.IsNullOrEmpty(userEmail))
                throw new ArgumentNullException("userEmail cannot be null or empty");
            if (string.IsNullOrEmpty(calendarID))
                throw new ArgumentNullException("calendarID cannot be null or empty");
            EventList result = new EventList();
            result.value = new List<Event>();
            bool alldone = false;
            int skip = 0;
            string requestUrl;
            while (!alldone)
            {
                requestUrl = String.Format("/users/{0}/calendars/{1}/calendarview?startDateTime={2}&endDateTime={3}&$skip={4}", userEmail, calendarID, startDate.ToString("s"), enddate.ToString("s"), skip);
                // Formating startDate and enddate according to https://docs.microsoft.com/en-us/dotnet/standard/base-types/standard-date-and-time-format-strings#Sortable
                HttpResponseMessage httpresult = MakeApiCall("GET", requestUrl, userEmail, null, null);
                string response = httpresult.Content.ReadAsStringAsync().Result;
                log.LogEverything(t.GetMethodName("OutlookExchangeOnlineAPIClient"), "ApiClient.GetCalendarItems response is " + response);
                if (response.Contains("odata"))
                {
                    EventList curpage = JsonConvert.DeserializeObject<EventList>(response);
                    if (curpage.value.Count >= 10) alldone = false;
                    else alldone = true;
                    foreach (Event curevent in curpage.value)
                    {
                        result.value.Add(curevent);
                    }
                    skip += curpage.value.Count;
                }
                else
                {
                    alldone = true;
                }

            }
            return result;
        }

        public Event GetEvent(string globalId, string userEmail)
        {
            if (string.IsNullOrEmpty(globalId))
                throw new ArgumentNullException("globalId cannot be null or empty");
            if (string.IsNullOrEmpty(userEmail))
                throw new ArgumentNullException("userEmail cannot be null or empty");
            log.LogEverything(t.GetMethodName("OutlookExchangeOnlineAPIClient"), "ApiClient.GetEvent globalId is " + globalId);
            log.LogEverything(t.GetMethodName("OutlookExchangeOnlineAPIClient"), "ApiClient.GetEvent userEmail is " + userEmail);
            string requestUrl = String.Format("/users/{0}/events/{1}", userEmail, globalId);
            HttpResponseMessage result = MakeApiCall("GET", requestUrl, userEmail, null, null);
            string response = result.Content.ReadAsStringAsync().Result;
            log.LogEverything(t.GetMethodName("OutlookExchangeOnlineAPIClient"), "ApiClient.GetEvent response is " + response);
            if (response.Contains("odata"))
            {
                return JsonConvert.DeserializeObject<Event>(response);
            }
            else
            {
                if (response.Contains("error"))
                {
                    throw new Exception("Item not found");
                }
                return null;
            }
        }

        public Event UpdateEvent(string userEmail, string EventID, string update)
        {
            if (string.IsNullOrEmpty(userEmail))
                throw new ArgumentNullException("userEmail cannot be null or empty");
            if (string.IsNullOrEmpty(EventID))
                throw new ArgumentNullException("EventID cannot be null or empty");
            if (string.IsNullOrEmpty(update))
                throw new ArgumentNullException("update cannot be null or empty");
            string requestUrl = String.Format("/users/{0}/events/{1}", userEmail, EventID);
            HttpResponseMessage result = MakeApiCall("PATCH", requestUrl, userEmail, update, null);
            string response = result.Content.ReadAsStringAsync().Result;
            if (response.Contains("odata"))
            {
                return JsonConvert.DeserializeObject<Event>(response);
            }
            else
            {
                return null;
            }


        }
        public Event CreateEvent(string userEmail, string CalendarID, string eventObject)
        {
            if (string.IsNullOrEmpty(userEmail))
                throw new ArgumentNullException("userEmail cannot be null or empty");
            if (string.IsNullOrEmpty(CalendarID))
                throw new ArgumentNullException("CalendarID cannot be null or empty");
            if (string.IsNullOrEmpty(eventObject))
                throw new ArgumentNullException("eventObject cannot be null or empty");
            string requestUrl = String.Format("/users/{0}/calendars/{1}/events", userEmail, CalendarID);
            HttpResponseMessage httpresult = MakeApiCall("POST", requestUrl, userEmail, eventObject, null);
            string response = httpresult.Content.ReadAsStringAsync().Result;
            if (response.Contains("odata"))
            {
                return JsonConvert.DeserializeObject<Event>(response);
            }
            else
            {
                return null;
            }

        }
        public void DeleteEvent(string userEmail, string EventID)
        {
            if (string.IsNullOrEmpty(userEmail))
                throw new ArgumentNullException("userEmail cannot be null or empty");
            if (string.IsNullOrEmpty(EventID))
                throw new ArgumentNullException("EventID cannot be null or empty");
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

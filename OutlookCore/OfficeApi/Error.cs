using Newtonsoft.Json;

namespace OutlookExchangeOnlineAPI
{
    public class Error
    {
        [JsonProperty(PropertyName = "code")]
        public string Code { get; set; }
        [JsonProperty(PropertyName = "message")]
        public string Message { get; set; }
    }

    public class ErrorResponse
    {
        [JsonProperty(PropertyName = "error")]
        public Error Error { get; set; }
    }
}

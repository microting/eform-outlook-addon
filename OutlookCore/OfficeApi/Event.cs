using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookExchangeOnlineAPI
{
    public class ResponseStatus
    {
        public string Response { get; set; }
        public DateTime Time { get; set; }
    }

    public class Body
    {
        public string ContentType { get; set; }
        public string Content { get; set; }
    }

    public class Start
    {
        public DateTime DateTime { get; set; }
        public string TimeZone { get; set; }
    }

    public class End
    {
        public DateTime DateTime { get; set; }
        public string TimeZone { get; set; }
    }

    public class Address
    {
        public string Type { get; set; }
        public string Street { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string CountryOrRegion { get; set; }
        public string PostalCode { get; set; }
    }

    public class Coordinates
    {
        public double? Latitude { get; set; }
        public double? Longitude { get; set; }
    }

    public class Location
    {
        public string DisplayName { get; set; }
        public Address Address { get; set; }
        public Coordinates Coordinates { get; set; }
    }

    public class Organizer
    {
        public EmailAddress EmailAddress { get; set; }
    }

    public class Event
    {
        public string Id { get; set; }
        public DateTime CreatedDateTime { get; set; }
        public DateTime LastModifiedDateTime { get; set; }
        public string ChangeKey { get; set; }
        public List<object> Categories { get; set; }
        public string OriginalStartTimeZone { get; set; }
        public string OriginalEndTimeZone { get; set; }
        public string iCalUId { get; set; }
        public int ReminderMinutesBeforeStart { get; set; }
        public bool IsReminderOn { get; set; }
        public bool HasAttachments { get; set; }
        public string Subject { get; set; }
        public string BodyPreview { get; set; }
        public string Importance { get; set; }
        public string Sensitivity { get; set; }
        public bool IsAllDay { get; set; }
        public bool IsCancelled { get; set; }
        public bool IsOrganizer { get; set; }
        public bool ResponseRequested { get; set; }
        public string SeriesMasterId { get; set; }
        public string ShowAs { get; set; }
        public string Type { get; set; }
        public string WebLink { get; set; }
        public object OnlineMeetingUrl { get; set; }
        public ResponseStatus ResponseStatus { get; set; }
        public Body Body { get; set; }
        public Start Start { get; set; }
        public End End { get; set; }
        public Location Location { get; set; }
        public object Recurrence { get; set; }
        public List<object> Attendees { get; set; }
        public Organizer Organizer { get; set; }
    }

    public class EventList
    {
        public List<Event> value { get; set; }
    }
}

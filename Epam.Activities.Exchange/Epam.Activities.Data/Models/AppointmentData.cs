// License placeholder

using System;
using Microsoft.Exchange.WebServices.Data;

namespace Epam.Activities.Exchange.Data.Models
{
    /// <summary>
    /// Model with necessary data for particular appointment.
    /// </summary>
    public class AppointmentData
    {
        /// <summary>
        /// Gets or sets appointment id.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets appointment subject.
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// Gets or sets appointment location.
        /// </summary>
        public string Location { get; set; }

        /// <summary>
        /// Gets or sets appointment body.
        /// </summary>
        public string Body { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether body is html or not.
        /// </summary>
        public bool IsBodyHtml { get; set; }

        /// <summary>
        /// Gets or sets start time.
        /// </summary>
        public DateTime StartTime { get; set; }
        
        /// <summary>
        /// Gets or sets end time.
        /// </summary>
        public DateTime EndTime { get; set; }

        /// <summary>
        /// Gets or sets type of Appointment.
        /// Detailed description <see ref="https://docs.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.appointmenttype?view=exchange-ews-api"/>
        /// </summary>
        public AppointmentType Type { get; set; } 

        /// <summary>
        /// Gets or sets required attendees collection.
        /// </summary>
        public Attendee[] RequiredAttendees { get; set; }

        /// <summary>
        /// Gets or sets optional attendees collection.
        /// </summary>
        public Attendee[] OptionalAttendees { get; set; }

        /// <summary>
        /// Gets or sets recurrence.
        /// </summary>
        public Recurrence Recurrence { get; set; }
    }
}

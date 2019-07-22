// License placeholder

using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Epam.Activities.Exchange.Services;
using Microsoft.Exchange.WebServices.Data;

namespace Epam.Activities.Exchange.Appointments
{
    /// <summary>
    /// Forward appointment to attendees again.
    /// </summary>
    [Description("Forward appointment to attendees again")]
    public class ForwardToRemind : ExchangeActivityBase
    {
        /// <summary>
        /// Gets or sets appointment Id
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> AppointmentId { get; set; }

        /// <summary>
        /// Gets or sets text for reminder.
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> ReminderText { get; set; }

        /// <inheritdoc />
        protected override void Execute(CodeActivityContext context)
        {
            // The activity will forward the AppointmentId referecned appointment only to not-responding attendees with ReminderText as a forward
            var service = ExchangeHelper.GetService(context.GetValue(OrganizerPassword), context.GetValue(ExchangeUrl), context.GetValue(OrganizerEmail));

            var appointment = AppointmentHelper.GetAppointmentById(service, context.GetValue(AppointmentId));

            if (appointment == null)
            {
                return;
            }

            var toRemind = new List<EmailAddress>();

            toRemind.AddRange(
                appointment
                    .RequiredAttendees
                    .Where(x => x.ResponseType == MeetingResponseType.Unknown || x.ResponseType == MeetingResponseType.NoResponseReceived));
            
            toRemind.AddRange(
                appointment
                    .OptionalAttendees
                    .Where(x => x.ResponseType == MeetingResponseType.Unknown || x.ResponseType == MeetingResponseType.NoResponseReceived));

            // Remind if any
            if (toRemind.Count > 0)
            {
                appointment.Forward(context.GetValue(ReminderText), toRemind);
            }
        }
    }
}

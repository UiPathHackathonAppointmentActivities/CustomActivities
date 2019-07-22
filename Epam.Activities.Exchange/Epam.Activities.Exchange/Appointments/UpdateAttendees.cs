// License placeholder

using System.Activities;
using System.ComponentModel;
using Epam.Activities.Exchange.Services;
using Microsoft.Exchange.WebServices.Data;

namespace Epam.Activities.Exchange.Appointments
{
    /// <summary>
    /// Updates list of required and optional attendees.
    /// </summary>
    [Description("Updates list of required and optional attendees")]
    public class UpdateAttendees : ExchangeActivityBase
    {
        /// <summary>
        /// Gets or sets appointment id.
        /// </summary>
        [RequiredArgument]
        [Category("Input")]
        public InArgument<string> AppointmentId { get; set; }

        /// <summary>
        /// Gets or sets collection of required attendees.
        /// </summary>
        [RequiredArgument]
        [Category("Input")]
        public InArgument<Attendee[]> RequiredAttendees { get; set; }

        /// <summary>
        /// Gets or sets list of optional attendees.
        /// </summary>
        [RequiredArgument]
        [Category("Input")]
        public InArgument<Attendee[]> OptionalAttendees { get; set; }

        /// <inheritdoc />
        protected override void Execute(CodeActivityContext context)
        {
            var service = ExchangeHelper.GetService(
                context.GetValue(OrganizerPassword),
                context.GetValue(ExchangeUrl),
                context.GetValue(OrganizerEmail));

            // GetMaster Item
            // Bind to all RequiredAttendees.
            var meeting = Appointment.Bind(service, new ItemId(context.GetValue(AppointmentId)), AppointmentHelper.GetAttendeesPropertySet());

            AppointmentHelper.UpdateAttendees(meeting, context.GetValue(RequiredAttendees), context.GetValue(OptionalAttendees));

            // Save and Send Updates
            meeting.Update(ConflictResolutionMode.AlwaysOverwrite, SendInvitationsOrCancellationsMode.SendOnlyToChanged);
        }
    }
}

// License placeholder

using System.Activities;
using System.ComponentModel;
using Epam.Activities.Exchange.Services;
using Microsoft.Exchange.WebServices.Data;

namespace Epam.Activities.Exchange.Appointments
{
    /// <inheritdoc />
    /// <summary>
    /// Resolve names and returns collection of <see cref="Attendee"/>.
    /// </summary>
    [Description("Validates names ")]
    public class ResolveNames : ExchangeActivityBase
    {
        /// <summary>
        /// Gets or sets collection of attendee names.
        /// </summary>
        [RequiredArgument]
        [Category("Input")]
        [Description("RequiredAttendees display name")]
        public InArgument<string[]> AttendeeNames { get; set; }

        /// <summary>
        /// Gets or sets indicator if we need to continue in case of errors or not.
        /// </summary>
        [RequiredArgument]
        [Category("Input")]
        [Description("Indicates if we need to continue in case of errors or not")]
        public InArgument<bool> ContinueOnError { get; set; }

        /// <summary>
        /// Gets or sets resolving results.
        /// </summary>
        [RequiredArgument]
        [Category("Output")]
        [Description("RequiredAttendees mail addresses")]
        public OutArgument<Attendee[]> Attendees { get; set; }

        /// <inheritdoc />
        protected override void Execute(CodeActivityContext context)
        {
            var service = ExchangeHelper.GetService(context.GetValue(OrganizerPassword), context.GetValue(ExchangeUrl), context.GetValue(OrganizerEmail));
            var continueOnError = context.GetValue(ContinueOnError);

            var attendees = AppointmentHelper.ResolveAttendeeNames(service, context.GetValue(AttendeeNames), continueOnError);

            context.SetValue(Attendees, attendees.ToArray());
        }
    }
}

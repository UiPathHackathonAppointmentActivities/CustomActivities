// License placeholder

using System.Activities;
using System.ComponentModel;
using System.Data;
using Epam.Activities.Exchange.Appointments;
using Epam.Activities.Exchange.Data.Models;
using Epam.Activities.Exchange.Services;
using Microsoft.Exchange.WebServices.Data;

namespace Epam.Activities.Exchange.BulkOperations
{
    /// <summary>
    /// Synchronize sequence of events with startTime, endTime and location with Exchange Appointment.
    /// </summary>
    [Description("Synchronize sequence of event with startTime, endTime and location with Exchange Appointment")]
    public class SynchronizeAppointment : ExchangeActivityBase
    {
        /// <summary>
        /// Gets or sets subject of appointment.
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> Subject { get; set; }

        /// <summary>
        /// Gets or sets body of appointment.
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> Body { get; set; }

        /// <summary>
        /// Gets or sets indicator that shows is body html.
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        public InArgument<bool> IsBodyHtml { get; set; }

        /// <summary>
        /// Gets or sets list of substrings that appointment body should contain. Used to verify that HTML body is not changed.
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        [Description("List of substrings that appointment body should contain. Used to verify that HTML body is not changed.")]
        public InArgument<string[]> BodyCheckSubstrings { get; set; }

        /// <summary>
        /// Gets or sets collection of required attendees.
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        public InArgument<Attendee[]> RequiredAtendees { get; set; }

        /// <summary>
        /// Gets or sets collection of optional attendees.
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        public InArgument<Attendee[]> OptionalAtendees { get; set; }

        /// <summary>
        /// Gets or sets scheduling data.
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        [Description("Required columns: 'Start', 'End' (of DateTime), 'Location' (of string) ")]
        public InArgument<DataTable> Data { get; set; }

        /// <summary>
        /// Gets or sets resulting patterns.
        /// </summary>
        [Category("Output")]
        public OutArgument<WeeklyPattern[]> MasterPatterns { get; set; }

        /// <summary>
        /// Gets or sets Ids of created appointments.
        /// </summary>
        [Category("Output")]
        public OutArgument<string[]> MasterIds { get; set; }

        /// <inheritdoc />
        protected override void Execute(CodeActivityContext context)
        {
            var synchronizer = new AppointmentsSynchronizer(context.GetValue(OrganizerEmail), context.GetValue(OrganizerPassword), context.GetValue(ExchangeUrl))
            {
                Subject = context.GetValue(Subject),
                Body = context.GetValue(Body),
                BodyIsHtml = context.GetValue(IsBodyHtml),
                BodyCheckSubstrings = context.GetValue(BodyCheckSubstrings),
                RequiredAttendees = context.GetValue(RequiredAtendees),
                OptionalAttendees = context.GetValue(OptionalAtendees),
                Data = context.GetValue(Data)
            };

            synchronizer.Sync();

            context.SetValue(MasterPatterns, synchronizer.WeeklyPatterns);
            context.SetValue(MasterIds, synchronizer.MasterIds);
        }
    }
}

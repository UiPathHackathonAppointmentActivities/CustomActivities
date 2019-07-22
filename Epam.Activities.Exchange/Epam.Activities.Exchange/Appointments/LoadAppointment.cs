// License placeholder

using System;
using System.Activities;
using System.ComponentModel;
using Epam.Activities.Exchange.Services;
using Microsoft.Exchange.WebServices.Data;

namespace Epam.Activities.Exchange.Appointments
{
    /// <summary>
    /// Loads appointment by Id or specific data
    /// </summary>
    [Description("Loads appointment by Id or specific data")]
    public class LoadAppointment : ExchangeActivityBase
    {
        /// <summary>
        /// Gets or sets search appointment by id.
        /// </summary>
        [Description("Search appointment by id.")]
        [RequiredArgument]
        [Category("Input")]
        [OverloadGroup("FindById")]
        public InArgument<string> AppointmentId { get; set; }

        /// <summary>
        /// Gets or sets search appointment by subject and date.
        /// </summary>
        [RequiredArgument]
        [Description("Search appointment by subject and date")]
        [Category("Input")]
        [OverloadGroup("FindByInfo")]
        public InArgument<string> Subject { get; set; }

        /// <summary>
        /// Gets or sets search appointment by subject and date.
        /// </summary>
        [RequiredArgument]
        [Category("Input")]
        [Description("Search appointment by subject and date")]
        [OverloadGroup("FindByInfo")]
        public InArgument<string> AppointmentDate { get; set; }

        /// <summary>
        /// Gets or sets found appointment if any.
        /// </summary>
        [Category("Output")]
        [Description("Found appointment if any")]
        [RequiredArgument]
        public OutArgument<Appointment> FoundAppointment { get; set; }

        /// <inheritdoc />
        protected override void Execute(CodeActivityContext context)
        {
            var service = ExchangeHelper.GetService(context.GetValue(OrganizerPassword), context.GetValue(ExchangeUrl), context.GetValue(OrganizerEmail));

            Appointment meeting;

            var id = context.GetValue(AppointmentId);

            if (string.IsNullOrEmpty(id))
            {
                var start = DateTime.Parse(context.GetValue(AppointmentDate));
                meeting = AppointmentHelper.GetAppointmentBySubject(service, context.GetValue(Subject), start);
            }
            else
            {
                meeting = AppointmentHelper.GetAppointmentById(service, id);
            }

            context.SetValue(FoundAppointment, meeting);
        }
    }
}

// License placeholder

using System;
using System.Activities;
using System.ComponentModel;
using Epam.Activities.Exchange.Services;
using Microsoft.Exchange.WebServices.Data;

namespace Epam.Activities.Exchange.Appointments
{
    /// <inheritdoc />
    /// <summary>
    /// Deletes appointment by ID.
    /// </summary>
    [Description("Deletes appointment by Id")]
    public class DeleteAppointment : ExchangeActivityBase
    {
        /// <summary>
        /// Gets or sets id of appointment to delete.
        /// </summary>
        [RequiredArgument]
        [Category("Input")]
        [Description("Id of appointment to delete")]
        public InArgument<string> AppointmentId { get; set; }

        /// <inheritdoc />
        protected override void Execute(CodeActivityContext context)
        {
            var service = ExchangeHelper.GetService(
                context.GetValue(OrganizerPassword),
                context.GetValue(ExchangeUrl),
                context.GetValue(OrganizerEmail));

            var id = context.GetValue(AppointmentId);

            var meeting = Appointment.Bind(service, new ItemId(id), new PropertySet(AppointmentSchema.Recurrence));

            meeting.Delete(DeleteMode.MoveToDeletedItems, SendCancellationsMode.SendToAllAndSaveCopy);
        }
    }
}

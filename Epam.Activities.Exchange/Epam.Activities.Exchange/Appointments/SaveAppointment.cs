// License placeholder

using System;
using System.Activities;
using System.ComponentModel;
using Epam.Activities.Exchange.Data.Models;
using Epam.Activities.Exchange.Services;
using Microsoft.Exchange.WebServices.Data;

namespace Epam.Activities.Exchange.Appointments
{
    /// <summary>
    /// Creates or updates appointment.
    /// </summary>
    [Description("Creates or updates appointment")]
    public class SaveAppointment : ExchangeActivityBase
    {
        /// <summary>
        /// Gets or sets local instance of appointment to synchronize it with service. 
        /// </summary>
        [Category("Appointment")]
        [OverloadGroup("Appointment")]
        [RequiredArgument]
        public InArgument<AppointmentData> Appointment { get; set; }

        /// <summary>
        /// Gets or sets appointment subject
        /// </summary>
        [Category("AppointmentDetails")]
        [OverloadGroup("AppointmentDetails")]
        [RequiredArgument]
        public InArgument<string> Subject { get; set; }

        /// <summary>
        /// Gets or sets location.
        /// </summary>
        [Category("AppointmentDetails")]
        [OverloadGroup("AppointmentDetails")]
        public InArgument<string> Location { get; set; }

        /// <summary>
        /// Gets or sets body.
        /// </summary>
        [Category("AppointmentDetails")]
        [OverloadGroup("AppointmentDetails")]
        public InArgument<string> Body { get; set; }

        /// <summary>
        /// Gets or sets indicator if body is html.
        /// </summary>
        [Category("AppointmentDetails")]
        [OverloadGroup("AppointmentDetails")]
        public InArgument<bool> IsBodyHtml { get; set; }

        /// <summary>
        /// Gets or sets appointment start time.
        /// </summary>
        [Category("AppointmentDetails")]
        [OverloadGroup("AppointmentDetails")]
        [RequiredArgument]
        public InArgument<DateTime> StartTime { get; set; }
        
        /// <summary>
        /// Gets or sets appointment end time.
        /// </summary>
        [Category("AppointmentDetails")]
        [OverloadGroup("AppointmentDetails")]
        [RequiredArgument]
        public InArgument<DateTime> EndTime { get; set; }

        /// <summary>
        /// Gets or sets appointment type.
        /// </summary>
        [Category("AppointmentDetails")]
        [OverloadGroup("AppointmentDetails")]
        [RequiredArgument]
        public InArgument<AppointmentType> AppointmentType { get; set; }

        /// <summary>
        /// Gets or sets required attendees.
        /// </summary>
        [Category("AppointmentDetails")]
        [OverloadGroup("AppointmentDetails")]
        [RequiredArgument]
        public InArgument<Attendee[]> RequiredAttendees { get; set; }

        /// <summary>
        /// Gets or sets optional attendees.
        /// </summary>
        [Category("AppointmentDetails")]
        [OverloadGroup("AppointmentDetails")]
        public InArgument<Attendee[]> OptionalAttendees { get; set; }

        /// <summary>
        /// Gets or sets recurrence.
        /// </summary>
        [Category("AppointmentDetails")]
        [OverloadGroup("AppointmentDetails")]
        public InArgument<Recurrence> Recurrence { get; set; }
        
        /// <summary>
        /// Gets or sets appointment id.
        /// </summary>
        public OutArgument<string> AppointmentId { get; set; }

        /// <inheritdoc />
        protected override void Execute(CodeActivityContext context)
        {
            var appointmentData = context.GetValue(Appointment) ?? GetAppointmentFromParameters(context);

            var service = ExchangeHelper.GetService(context.GetValue(OrganizerPassword), context.GetValue(ExchangeUrl), context.GetValue(OrganizerEmail));

            var recurrMeeting = new Appointment(service)
            {
                Subject = appointmentData.Subject,
                Body = new MessageBody(appointmentData.IsBodyHtml ? BodyType.HTML : BodyType.Text, appointmentData.Body),
                Start = appointmentData.StartTime,
                End = appointmentData.EndTime,
                Location = appointmentData.Subject,
                Recurrence = appointmentData.Recurrence
            };

            var requiredAttendees = context.GetValue(RequiredAttendees);
            var optionalAttendees = context.GetValue(OptionalAttendees);
            AppointmentHelper.UpdateAttendees(recurrMeeting, requiredAttendees, optionalAttendees);

            // This method results in in a CreateItem call to EWS.
            recurrMeeting.Save(SendInvitationsMode.SendToAllAndSaveCopy);

            context.SetValue(AppointmentId, recurrMeeting.Id);
        }

        /// <summary>
        /// Initialize <see cref="AppointmentData"/> instance from input parameters.
        /// </summary>
        /// <param name="context">Activity context.</param>
        /// <returns>Instance of <see cref="AppointmentData"/></returns>
        private AppointmentData GetAppointmentFromParameters(ActivityContext context)
        {
            return new AppointmentData
            {
                IsBodyHtml = context.GetValue(IsBodyHtml),
                Body = context.GetValue(Body),
                StartTime = context.GetValue(StartTime),
                EndTime = context.GetValue(EndTime),
                Subject = context.GetValue(Subject),
                RequiredAttendees = context.GetValue(RequiredAttendees),
                OptionalAttendees = context.GetValue(OptionalAttendees),
                Recurrence = context.GetValue(Recurrence)
            };
        }
    }
}

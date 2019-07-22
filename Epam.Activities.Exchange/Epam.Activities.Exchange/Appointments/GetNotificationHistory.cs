// License placeholder

using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Epam.Activities.Exchange.Data.Models;
using Epam.Activities.Exchange.Services;
using Microsoft.Exchange.WebServices.Data;

namespace Epam.Activities.Exchange.Appointments
{
    /// <summary>
    /// By Given subject and start date returns all notifications sent, namely - subject, date and time sent, semicolon separated list to.
    /// </summary>
    [Description("By Given subject and start date returns all notifications sent, namely - subject, date and time sent, semicolon separated list to")]
    public class GetNotificationHistory : ExchangeActivityBase
    {
        /// <summary>
        /// Gets or sets subject of appointment to search.
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> AppointmentSubject { get; set; }

        /// <summary>
        /// Gets or sets start date for search.
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        public InArgument<DateTime> AppointmentDateFrom { get; set; }

        /// <summary>
        /// Gets or sets max amount of sent notifications to load.
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        public InArgument<int> Amount { get; set; }

        /// <summary>
        /// Gets or sets notifications history.
        /// </summary>
        [Category("Output")]
        public OutArgument<NotifyHistory[]> SentNotificationsHistory { get; set; }

        /// <inheritdoc />
        protected override void Execute(CodeActivityContext context)
        {
            var service = ExchangeHelper.GetService(context.GetValue(OrganizerPassword), context.GetValue(ExchangeUrl), context.GetValue(OrganizerEmail));
            
            var start = context.GetValue(AppointmentDateFrom);
            var appointmentSubject = context.GetValue(AppointmentSubject);

            var amount = context.GetValue(Amount);

            var itemView = new ItemView(amount)
            {
                PropertySet = new PropertySet(
                    ItemSchema.Subject,
                    AppointmentSchema.Start,
                    ItemSchema.DisplayTo,
                    AppointmentSchema.AppointmentType,
                    ItemSchema.DateTimeSent)
            };

            // Find appointments by subject.
            var substrFilter = new SearchFilter.ContainsSubstring(ItemSchema.Subject, appointmentSubject);
            var startFilter = new SearchFilter.IsGreaterThanOrEqualTo(AppointmentSchema.Start, start);

            var filterList = new List<SearchFilter>
            {
                substrFilter,
                startFilter
            };

            var calendarFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, filterList);

            var results = service.FindItems(WellKnownFolderName.SentItems, calendarFilter, itemView);

            var history = results.Select(item => new NotifyHistory
            {
                Subject = item.Subject,
                SentOn = item.DateTimeSent,
                SentTo = item.DisplayTo
            }).ToArray();

            context.SetValue(SentNotificationsHistory, history);
        }
    }
}

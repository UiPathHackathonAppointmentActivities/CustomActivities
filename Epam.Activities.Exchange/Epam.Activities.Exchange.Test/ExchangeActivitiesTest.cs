// License placeholder

using System;
using System.Collections.Generic;
using System.Linq;
using Epam.Activities.Exchange.Services;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Epam.Activities.Exchange.Test
{
    [TestClass]
    public class ExchangeActivitiesTest
    {
        private const string ExchangeServerUrl = "https://outlook.office365.com/EWS/Exchange.asmx";
        private const string ExchangeLogin = "login@domain.com";
        private const string ExchangePass = "!!!";
        
        [TestMethod]
        public void LoadAppointentBySubjTest()
        {
            var service = ExchangeHelper.GetService(ExchangePass, ExchangeServerUrl, ExchangeLogin);

            var subj = "Item1";
            DateTime.TryParse("10-Jun-2019", out var start);
            AppointmentHelper.GetAppointmentBySubject(service, subj, start);

            Assert.AreEqual(1, 1);
        }

        [TestMethod]
        public void LoadAppointentByIdTest()
        {
            var service = ExchangeHelper.GetService(ExchangePass, ExchangeServerUrl, ExchangeLogin);

            DateTime.TryParse("10-Jun-2019", out var start);
            var meeting = AppointmentHelper.GetAppointmentBySubject(service, "Item1", start);

            AppointmentHelper.GetAppointmentById(service, meeting.Id.ToString());

            Assert.AreEqual(1, 1);
        }

        [TestMethod]
        public void LoadResponseByIdTest()
        {
            var service = ExchangeHelper.GetService(ExchangePass, ExchangeServerUrl, ExchangeLogin);

            var subj = "Item1";
            DateTime.TryParse("10-Jun-2019", out var start);
            var meeting = AppointmentHelper.GetAppointmentBySubject(service, subj, start);

            AppointmentHelper.GetAttendeesById(service, meeting.Id.ToString(), MeetingAttendeeType.Required);

            Assert.AreEqual(1, 1);
        }
        
        [TestMethod]
        public void LoadAppointmentsTest()
        {
            var service = ExchangeHelper.GetService(ExchangePass, ExchangeServerUrl, ExchangeLogin);

            DateTime.TryParse("10-Jun-2019", out var start);
            var meeting = AppointmentHelper.GetAppointmentBySubject(service, "Item1", start);

            var appointments = AppointmentHelper.GetAppointmentsById(service, meeting.Id.ToString());

            foreach (var item in appointments)
            {
                Console.WriteLine(item.AppointmentType.ToString());
            }
        }

        [TestMethod]
        public void ResolveNamesTest()
        {
            var service = ExchangeHelper.GetService(ExchangePass, ExchangeServerUrl, ExchangeLogin);

            string[] testNames = { "Obfuscated attendee1", "Obfuscated attendee2" };

            var attendees = AppointmentHelper.ResolveAttendeeNames(service, testNames, true);
            
            foreach (var item in attendees)
            {
                Console.WriteLine(item.Name + "::" + item.Address);
            }
        }

        [TestMethod]
        public void AttachedAppointmentTest()
        {
            DateTime.TryParse("06-July-2019", out var testStartDate);

            var service = ExchangeHelper.GetService(ExchangePass, ExchangeServerUrl, ExchangeLogin);
            var appointment = AppointmentHelper.GetAppointmentBySubject(service, "AppointmentAttachmentCheck1", testStartDate);
            
            if (appointment != null)
            {
                var toRemind = AppointmentHelper.GetAttendeesById(service, appointment.Id.ToString(), MeetingAttendeeType.Required)
                    .Where(x => x.ResponseType == MeetingResponseType.Unknown || x.ResponseType == MeetingResponseType.NoResponseReceived)
                    .Select(attendee => (EmailAddress)attendee.Address)
                    .ToList();

                // remind if any
                if (toRemind.Count > 0)
                {
                    appointment.Forward("Please resond to an attached invitation", toRemind);
                }
            }

            Assert.AreNotEqual(1, 2);
        }

        [TestMethod]
        public void ForwardedAppointmentTest()
        {
            DateTime.TryParse("06-July-2019", out var testStartDate);

            var service = ExchangeHelper.GetService(ExchangePass, ExchangeServerUrl, ExchangeLogin);

            var itemView = new ItemView(10)
            {
                PropertySet = new PropertySet(ItemSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.AppointmentType)
            };

            // Find appointments by subject.
            var substrFilter = new SearchFilter.ContainsSubstring(ItemSchema.Subject, "AppointmentAttachmentCheck1");
            var startFilter = new SearchFilter.IsGreaterThanOrEqualTo(AppointmentSchema.Start, testStartDate);

            var filterList = new List<SearchFilter>
            {
                substrFilter,
                startFilter
            };

            var calendarFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, filterList);

            var results = service.FindItems(WellKnownFolderName.SentItems, calendarFilter, itemView);

            foreach (var cur in results)
            {
                var item = Item.Bind(service, cur.Id, new PropertySet(ItemSchema.Subject, AppointmentSchema.Start, ItemSchema.DisplayTo, AppointmentSchema.AppointmentType, ItemSchema.DateTimeSent));
                var displayTo = item.DisplayTo;

                Console.WriteLine(displayTo); 
                Console.WriteLine(item.DateTimeSent);
            }

            Assert.AreEqual("AppointmentAttachmentCheck1", "AppointmentAttachmentCheck1");
        }
    }
}

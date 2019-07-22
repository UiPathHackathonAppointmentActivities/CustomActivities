// License placeholder

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using Microsoft.Exchange.WebServices.Data;

namespace Epam.Activities.Exchange.Services
{
    /// <summary>
    /// Static helper for managing appointments.
    /// </summary>
    public static class AppointmentHelper
    {
        /// <summary>
        /// Receive collection of attendees of specified type.
        /// </summary>
        /// <param name="service">Exchange service instance.</param>
        /// <param name="id">Id of appointment.</param>
        /// <param name="attendeeType">Type of attendees to search.</param>
        /// <returns>Collection of attendees of specified type.</returns>
        public static IEnumerable<Attendee> GetAttendeesById(ExchangeService service, string id, MeetingAttendeeType attendeeType)
        {
            // Bind to all RequiredAttendees
            var meeting = Appointment.Bind(service, new ItemId(id), GetAttendeesPropertySet());

            // Return according to requested type
            switch (attendeeType)
            {
                case MeetingAttendeeType.Required:
                {
                    return meeting.RequiredAttendees;
                }

                case MeetingAttendeeType.Optional:
                {
                    return meeting.OptionalAttendees;
                }

                case MeetingAttendeeType.Resource:
                {
                    return meeting.Resources;
                }

                case MeetingAttendeeType.Organizer:
                {
                    return new[] { new Attendee(meeting.Organizer) };
                }

                default:
                {
                    throw new ArgumentOutOfRangeException(nameof(attendeeType), attendeeType, null);
                }
            }
        }

        /// <summary>
        /// Updates attendees collections in appointment
        /// </summary>
        /// <param name="appointment">Appointment to update.</param>
        /// <param name="requiredAttendees">Required attendees new collection.</param>
        /// <param name="optionalAttendees">Optional attendees new collection.</param>
        public static void UpdateAttendees(Appointment appointment, Attendee[] requiredAttendees, Attendee[] optionalAttendees)
        {
            // Update required attendees.
            appointment.RequiredAttendees.Clear();
            if (requiredAttendees != null)
            {
                foreach (var requiredAttendee in requiredAttendees)
                {
                    appointment.RequiredAttendees.Add(requiredAttendee);
                }
            }

            // Update optional attendees.
            appointment.OptionalAttendees.Clear();

            if (optionalAttendees == null)
            {
                return;
            }

            foreach (var optionalAttendee in optionalAttendees)
            {
                appointment.OptionalAttendees.Add(optionalAttendee);
            }
        }

        /// <summary>
        /// Property set of attendees.
        /// </summary>
        /// <returns>Instance of <see cref="PropertySet"/> for attendees</returns>
        public static PropertySet GetAttendeesPropertySet()
        {
            return new PropertySet(
                AppointmentSchema.AppointmentType,
                AppointmentSchema.Recurrence,
                AppointmentSchema.RequiredAttendees,
                AppointmentSchema.OptionalAttendees,
                AppointmentSchema.Organizer,
                AppointmentSchema.Resources);
        }

        /// <summary>
        /// Property set of all fields.
        /// </summary>
        /// <returns>Instance of <see cref="PropertySet"/> for appointment</returns>
        public static PropertySet GetAppointmentPropetySet()
        {
            return new PropertySet(
                ItemSchema.Subject,
                AppointmentSchema.AppointmentType,
                AppointmentSchema.Recurrence,
                AppointmentSchema.IsRecurring,
                AppointmentSchema.FirstOccurrence,
                AppointmentSchema.LastOccurrence,
                AppointmentSchema.Organizer,
                AppointmentSchema.End,
                AppointmentSchema.Start,
                AppointmentSchema.RequiredAttendees,
                AppointmentSchema.ModifiedOccurrences,
                AppointmentSchema.DeletedOccurrences,
                ItemSchema.Body,
                AppointmentSchema.Location);
        }

        /// <summary>
        /// Property set of appointment info.
        /// </summary>
        /// <returns>Instance of <see cref="PropertySet"/> for appointment info.</returns>
        public static PropertySet GetFindAppointmentPropertySet()
        {
            return new PropertySet(
                ItemSchema.Subject,
                AppointmentSchema.AppointmentType,
                AppointmentSchema.IsRecurring,
                AppointmentSchema.Organizer,
                AppointmentSchema.End,
                AppointmentSchema.Start,
                AppointmentSchema.Location);
        }

        /// <summary>
        /// Resolves attendees names. Same logic as "Check names" in outlook.
        /// </summary>
        /// <param name="service">Exchange service.</param>
        /// <param name="names">Collection of names to resolve.</param>
        /// <param name="continueOnError">Indicates if we need proceed processing on error.</param>
        /// <returns>Collection of resolved names.</returns>
        public static List<Attendee> ResolveAttendeeNames(ExchangeService service, string[] names, bool continueOnError = false)
        {
            var attendees = new List<Attendee>();
            
            foreach (var name in names)
            {
                var currentResolved = service.ResolveName(name);

                if (currentResolved.Count == 0 && !continueOnError)
                {
                    throw new InvalidDataException($"Can't resolve given name '{name}'.");
                }

                if (currentResolved.Count > 1)
                {
                    if (!currentResolved.First().Mailbox.Name.ToLower().Equals(name.ToLower()))
                    {
                        throw new InvalidDataException($"Can't resolve given name '{name}' due to ambiguity.");
                    }
                }

                attendees.Add(new Attendee(currentResolved.First().Mailbox));
            }

            return attendees;
        }

        /// <summary>
        /// Secures string.
        /// </summary>
        /// <param name="plainText">String for securing.</param>
        /// <returns>Secured string.</returns>
        public static SecureString GetSecureString(string plainText)
        {
            var secret = new SecureString();

            foreach (var ch in plainText)
            {
                secret.AppendChar(ch);
            }

            return secret;
        }
        
        /// <summary>
        /// Gets appointment by id.
        /// </summary>
        /// <param name="service">Exchange service.</param>
        /// <param name="itemId">Appointment id.</param>
        /// <returns>Instance of <see cref="Appointment"/></returns>
        public static Appointment GetAppointmentById(ExchangeService service, string itemId)
        {
            var meetingProperties = GetAppointmentPropetySet();

            return Appointment.Bind(service, new ItemId(itemId), meetingProperties);
        }
        
        /// <summary>
        /// Gets first found appointment with specific subject starting from specific date.
        /// </summary>
        /// <param name="service">Exchange service.</param>
        /// <param name="subject">Subject of appointment to search.</param>
        /// <param name="start">Start date.</param>
        /// <returns>Found appointment if any.</returns>
        public static Appointment GetAppointmentBySubject(ExchangeService service, string subject, DateTime start)
        {
            Appointment meeting = null;

            var itemView = new ItemView(3)
            {
                PropertySet = new PropertySet(ItemSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.AppointmentType)
            };

            // Find appointments by subject.
            var substrFilter = new SearchFilter.ContainsSubstring(ItemSchema.Subject, subject);
            var startFilter = new SearchFilter.IsGreaterThanOrEqualTo(AppointmentSchema.Start, start);
            
            var filterList = new List<SearchFilter>
            {
                substrFilter,
                startFilter
            };

            var calendarFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, filterList);

            var results = service.FindItems(WellKnownFolderName.Calendar, calendarFilter, itemView);

            if (results.Items.Count == 1)
            {
                meeting = results.Items[0] as Appointment;
            }

            return meeting;
        }

        /// <summary>
        /// Gets all found appointment with specific subject starting from specific date.
        /// </summary>
        /// <param name="service">Exchange service.</param>
        /// <param name="subject">Subject of appointment to search.</param>
        /// <returns>Found appointment if any.</returns>
        public static Appointment[] GetAllAppointmentsBySubject(ExchangeService service, string subject)
        {
            var itemView = new ItemView(3)
            {
                PropertySet = GetFindAppointmentPropertySet()
            };

            var subjectFilter = new SearchFilter.IsEqualTo(ItemSchema.Subject, subject);

            var results = service.FindItems(WellKnownFolderName.Calendar, subjectFilter, itemView);

            var appts = results
                .Items
                .Where(item => item is Appointment)
                .Select(appt => Appointment.Bind(service, appt.Id, GetAppointmentPropetySet()))
                .ToArray();
            
            return appts.Where(x => x.AppointmentType == AppointmentType.RecurringMaster).SelectMany(x => GetAppointmentsById(service, x.Id.UniqueId)).ToArray();
        }

        /// <summary>
        /// Get appointment by id.
        /// </summary>
        /// <param name="service">Exchange service.</param>
        /// <param name="id">Appointment id.</param>
        /// <returns>Appointment if any.</returns>
        public static Appointment[] GetAppointmentsById(ExchangeService service, string id)
        {
            Appointment[] items;

            if (string.IsNullOrEmpty(id))
            {
                return null;
            }

            var masterId = new ItemId(id);

            Console.WriteLine("MasterId::" + masterId);

            var meeting = Appointment.Bind(service, masterId);

            switch (meeting.AppointmentType)
            {
                case AppointmentType.Occurrence:
                case AppointmentType.Exception:
                {
                    meeting = Appointment.BindToRecurringMaster(service, masterId);
                    break;
                }

                case AppointmentType.Single:
                {
                    Console.WriteLine("Single Appointment");

                    items = new[] { meeting };

                    return items;
                }
            }

            Console.WriteLine("Bound to Master");
            
            // iterate through all ocurencies
            Console.WriteLine("Enumeratig Occurrencies");

            var currentIndex = 1;
            var appointments = new List<Appointment>();
            Appointment current = null;

            do
            {
                try
                {
                    current = Appointment.BindToOccurrence(service, masterId, currentIndex++, GetAppointmentPropetySet());

                    appointments.Add(current);
                }
                catch (ServiceResponseException e)
                {
                    if (e.Message == "Occurrence index is out of recurrence range.")
                    {
                        break;
                    }

                    if (e.Message == "Occurrence with this index was previously deleted from the recurrence.")
                    {
                    }
                    else
                    {
                        throw;
                    }
                }
            }
            while (current != null && current.Start <= meeting.LastOccurrence.Start);

            items = appointments.ToArray();

            return items;
        }

        /// <summary>
        /// Check if any updates with attendees needed to prevent excess calls to server.
        /// </summary>
        /// <param name="appointment">Appointment to check.</param>
        /// <param name="requiredAttendees">Collection of required attendees.</param>
        /// <param name="optionalAttendees">Collection of optional attendees.</param>
        /// <returns>True if update needed. False, otherwise.</returns>
        public static bool IsSyncAttendeesNeede(Appointment appointment, Attendee[] requiredAttendees, Attendee[] optionalAttendees)
        {
            var requiredHash = appointment.RequiredAttendees.ToDictionary(x => x.Address, x => x);
            var optionalHash = appointment.OptionalAttendees.ToDictionary(x => x.Address, x => x);

            return 
                requiredAttendees.All(requiredAttendee => requiredHash.ContainsKey(requiredAttendee.Address)) 
                && optionalAttendees.All(optionalAttendee => optionalHash.ContainsKey(optionalAttendee.Address));
        }

        /// <summary>
        /// Compares attendees by their address.
        /// </summary>
        internal class AttendeessComparer : EqualityComparer<Attendee>
        {
            /// <inheritdoc />
            public override bool Equals(Attendee x, Attendee y)
            {
                return y != null && x != null && x.Address.Equals(y.Address, StringComparison.InvariantCultureIgnoreCase);
            }

            /// <inheritdoc />
            public override int GetHashCode(Attendee x)
            {
                return x == null ? 0 : x.GetHashCode();
            }
        }
    }
}

// License placeholder

using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using Epam.Activities.Exchange.Data.Models;
using Microsoft.Exchange.WebServices.Data;

namespace Epam.Activities.Exchange.Services
{
    /// <summary>
    /// Synchronize appointment in Exchange service with provided information.
    /// </summary>
    public class AppointmentsSynchronizer
    {
        /// <summary>
        /// Exchange service.
        /// </summary>
        private readonly ExchangeService _service;

        /// <summary>
        /// Culture info.
        /// </summary>
        private readonly CultureInfo _culture;

        /// <summary>
        /// Calendar instance.
        /// </summary>
        private readonly Calendar _calendar;

        /// <summary>
        /// Collection of <see cref="ScheduleRow"/> describes all occurrences.
        /// </summary>
        private List<ScheduleRow> _schedule = new List<ScheduleRow>();

        /// <summary>
        /// Raw data.
        /// </summary>
        private DataTable _data;
        
        /// <summary>
        /// Initializes a new instance of the <see cref="AppointmentsSynchronizer"/> class
        /// </summary>
        /// <param name="user">Account mail.</param>
        /// <param name="password">Account password.</param>
        /// <param name="exchangeUrl">Exchange service url.</param>
        public AppointmentsSynchronizer(string user, string password, string exchangeUrl)
        {
            _culture = CultureInfo.CurrentCulture;
            _calendar = _culture.Calendar;
            _service = ExchangeHelper.GetService(password, exchangeUrl, user);
        }

        /// <summary>
        /// Gets or sets table with events.
        /// </summary>
        public DataTable Data
        {
            get => _data;
            set
            {
                _data = value;
                _schedule = PrepareSchedule(_data);
            }
        }

        /// <summary>
        /// Gets or sets subject of master appointment.
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// Gets or sets body of master appointment.
        /// </summary>
        public string Body { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether if body is html
        /// </summary>
        public bool BodyIsHtml { get; set; }
        
        /// <summary>
        /// Gets or sets collection of substring to check.
        /// </summary>
        public string[] BodyCheckSubstrings { get; set; }

        /// <summary>
        /// Gets or sets collection of required attendees.
        /// </summary>
        public Attendee[] RequiredAttendees { get; set; }

        /// <summary>
        /// Gets or sets collection of optional attendees.
        /// </summary>
        public Attendee[] OptionalAttendees { get; set; }

        /// <summary>
        /// Gets patterns for scheduling.
        /// </summary>
        public WeeklyPattern[] WeeklyPatterns { get; private set; }

        /// <summary>
        /// Gets occurrences ids.
        /// </summary>
        public string[] MasterIds { get; private set; }

        /// <summary>
        /// Converts table with data to collection of <see cref="ScheduleRow"/>.
        /// </summary>
        /// <param name="data">Schedule data.</param>
        /// <returns>Collection of <see cref="ScheduleRow"/></returns>
        public static List<ScheduleRow> PrepareSchedule(DataTable data)
        {
            if (data == null)
            {
                return new List<ScheduleRow>();
            }

            return data
                .Rows
                .Cast<DataRow>()
                .Select(x => new ScheduleRow
                {
                    Start = x.Field<DateTime>("Start"),
                    End = x.Field<DateTime>("End"),
                    Location = x.Field<string>("Location"),
                    Matched = false
                })
                .OrderBy(x => x.Start)
                .ToList();
        }

        /// <summary>
        /// Determines occurrence pattern for collection of occurrences.
        /// </summary>
        /// <param name="schedule">Collection of occurrences.</param>
        /// <returns>Collection of patterns.</returns>
        public static WeeklyPattern[] DeterminePattern(List<ScheduleRow> schedule)
        {
            var weeksCount = (schedule.Last().Start - schedule.First().Start).Days / 7 + 1;
            var pattern = new Dictionary<TimeSpan, WeeklyPattern>();

            var patternCount = new Dictionary<string, DayPattern>();
            foreach (var row in schedule)
            {
                if (!patternCount.TryGetValue(DayPattern.GetKey(row.Start), out var dayPattern))
                {
                    dayPattern = DayPattern.NewPattern(row.Start, row.End, row.Location);
                    patternCount.Add(DayPattern.GetKey(row.Start), dayPattern);
                }

                dayPattern.Count++;
            }

            Console.WriteLine("Determine pattern metrics: weeksCount={0}, weeskCount mod 2 = {1}", weeksCount, weeksCount / 2);
            var dayPatterns = patternCount.Values
                .Where(dayPatt =>
                {
                    Console.WriteLine("Determine pattern metrics: dayCount - {0} = {1}", dayPatt, dayPatt.Count);
                    return dayPatt.Count >= weeksCount / 2;
                })
                .OrderBy(day => day.DayOfTheWeek);

            foreach (var day in dayPatterns)
            {
                if (!pattern.TryGetValue(day.StartTime, out var master))
                {
                    master = new WeeklyPattern
                    {
                        StartTime = day.StartTime,
                        EndTime = day.EndTime,
                        DaysOfTheWeek = new[] { day.DayOfTheWeek },
                        Location = day.Location
                    };

                    pattern.Add(day.StartTime, master);
                }
                else
                {
                    var list = master.DaysOfTheWeek.ToList();
                    list.Add(day.DayOfTheWeek);
                    master.DaysOfTheWeek = list.ToArray();
                }
            }

            return pattern.Values.ToArray();
        }

        /// <summary>
        /// Synchronize appointment with service.
        /// </summary>
        public void Sync()
        {
            ValidateParameters();

            var appointments = LoadAppointments();

            // If not Schedule then delete all and exit
            if (_schedule.Count == 0)
            {
                if (appointments == null)
                {
                    return;
                }

                Console.WriteLine("Schedule.Count = 0 => Deleting {0} appointments", appointments.Length);

                var appointmentsToDelete = appointments
                    .Where(x => x.AppointmentType == AppointmentType.RecurringMaster || x.AppointmentType == AppointmentType.Single);

                foreach (var item in appointmentsToDelete)
                {
                    item.Delete(DeleteMode.MoveToDeletedItems);
                }

                return;
            }

            // Sync master items
            WeeklyPatterns = DeterminePattern(_schedule);
            if (SyncMaster(WeeklyPatterns, appointments))
            {
                appointments = LoadAppointments();
            }

            MasterIds = appointments.Where(ap => ap.AppointmentType == AppointmentType.RecurringMaster).Select(ap => ap.Id.ToString()).ToArray();
            
            var matchedAppointments = new List<Appointment>();

            bool IsReadyToMatch(ScheduleRow row, Appointment appt)
            {
                return !row.Matched && matchedAppointments.IndexOf(appt) < 0 
                    && appt.AppointmentType != AppointmentType.RecurringMaster;
            }

            bool FromSameDay(ScheduleRow row, Appointment appt)
            {
                return IsReadyToMatch(row, appt) && appt.Start.Date == row.Start.Date;
            }

            bool FromSameWeek(DateTime start, DateTime end)
            {
                return start.Year == end.Year 
                    && _calendar.GetWeekOfYear(
                        start,
                        _culture.DateTimeFormat.CalendarWeekRule,
                        _culture.DateTimeFormat.FirstDayOfWeek) == _calendar.GetWeekOfYear(end, _culture.DateTimeFormat.CalendarWeekRule, _culture.DateTimeFormat.FirstDayOfWeek);
            }

            bool SameWeek(ScheduleRow row, Appointment appt) => IsReadyToMatch(row, appt) && FromSameWeek(appt.Start, row.Start);

            SendInvitationsOrCancellationsMode sendMode;

            // Match by same day
            foreach (var row in _schedule)
            {
                var appt = appointments.FirstOrDefault(item => FromSameDay(row, item));
                if (appt == null)
                {
                    continue;
                }

                matchedAppointments.Add(appt);
                row.Matched = true;

                if (!SyncAppointmentItem(row, appt, out sendMode))
                {
                    continue;
                }

                Console.WriteLine("Updating item in same day {0}, Subject: {1}", row.Start, Subject);
                appt.Update(ConflictResolutionMode.AlwaysOverwrite, sendMode);
            }

            // Match other by week and move appointment in week range
            foreach (var row in _schedule)
            {
                var appt = appointments.FirstOrDefault(item => SameWeek(row, item));
                if (appt != null)
                {
                    matchedAppointments.Add(appt);
                    row.Matched = true;

                    if (SyncAppointmentItem(row, appt, out sendMode))
                    {
                        Console.WriteLine("Updating item in same week {0}, Subject: {1}", row.Start, Subject);
                        appt.Update(ConflictResolutionMode.AlwaysOverwrite, sendMode);
                    }
                }
            }

            // Create missing appointments for rest rows
            foreach (var row in _schedule)
            {
                if (row.Matched)
                {
                    continue;
                }

                Console.WriteLine("Creating new item {0}, Subject: {1}", row.Start, Subject);
                row.Matched = true;
                var newAppt = new Appointment(_service);
                InitNewAppointment(row, newAppt);
                newAppt.Save(SendInvitationsMode.SendToAllAndSaveCopy);
            }

            // Delete not matched appointments
            foreach (var appt in appointments)
            {
                if (matchedAppointments.IndexOf(appt) < 0 && appt.AppointmentType != AppointmentType.RecurringMaster)
                {
                    Console.WriteLine("Deleting item {0}, Subject: {1}", appt.Start, Subject);
                    matchedAppointments.Add(appt);
                    appt.Delete(DeleteMode.MoveToDeletedItems);
                }
            }
        }

        /// <summary>
        /// Logs appointments collection.
        /// </summary>
        /// <param name="appointments">Appointments to log.</param>
        private static void LogAppointments(IEnumerable<Appointment> appointments)
        {
            foreach (var appointment in appointments)
            {
                Console.WriteLine("{0} | {1} | {2} | {3}", appointment.Subject, appointment.AppointmentType, appointment.Start, appointment.End);
                Console.WriteLine("\tLocation = {0}", appointment.Location);
                Console.WriteLine("\tBody Length = {0}", appointment.Body.Text.Length);
            }
        }

        /// <summary>
        /// Synchronize master appointments with server.
        /// </summary>
        /// <param name="weeklyPatterns">Patterns information.</param>
        /// <param name="appointments">Appointments to synchronize.</param>
        /// <returns>True if any changes was performed, False otherwise.</returns>
        private bool SyncMaster(WeeklyPattern[] weeklyPatterns, IEnumerable<Appointment> appointments)
        {
            var processedItems = new List<Appointment>();
            var processedPatterns = new List<WeeklyPattern>();
            
            bool MasterMatchPattern(WeeklyPattern pattern, Appointment master)
            {
                if (processedItems.IndexOf(master) >= 0)
                {
                    return false;
                }

                if (processedPatterns.IndexOf(pattern) >= 0)
                {
                    return false;
                }

                if (master.Start.TimeOfDay != pattern.StartTime)
                {
                    return false;
                }

                if (((Recurrence.WeeklyPattern)master.Recurrence).DaysOfTheWeek.Any(day => !pattern.DaysOfTheWeek.Contains(day)))
                {
                    return false;
                }

                foreach (var day in pattern.DaysOfTheWeek)
                {
                    if (!((Recurrence.WeeklyPattern)master.Recurrence).DaysOfTheWeek.Contains(day))
                    {
                        return false;
                    }
                }

                return true;
            }

            var result = false;
            
            var masterAppointments = appointments.Where(item => item.AppointmentType == AppointmentType.RecurringMaster).ToList();

            // Update matched
            foreach (var masterAppointment in masterAppointments)
            {
                var pattern = weeklyPatterns.FirstOrDefault(patt => MasterMatchPattern(patt, masterAppointment));
                if (pattern == null)
                {
                    continue;
                }

                processedItems.Add(masterAppointment);
                processedPatterns.Add(pattern);

                if (!SyncMasterItem(pattern, masterAppointment, out var sendMode))
                {
                    continue;
                }

                Console.WriteLine("Updating master, Subject {0}", Subject);
                masterAppointment.Update(ConflictResolutionMode.AlwaysOverwrite, sendMode);
                result = true;
            }

            // Delete not matched masters
            foreach (var masterIt in masterAppointments)
            {
                if (processedItems.IndexOf(masterIt) >= 0)
                {
                    continue;
                }

                Console.WriteLine("Deleting master appointment, Subject: {0}", Subject);
                masterIt.Delete(DeleteMode.MoveToDeletedItems);
                result = true;
            }

            // Create new masters for not matched patterns
            foreach (var pattern in weeklyPatterns)
            {
                if (processedPatterns.IndexOf(pattern) >= 0)
                {
                    continue;
                }

                Console.WriteLine("Creating master appointment, Subject: {0}", Subject);
                var master = new Appointment(_service);
                InitNewAppointment(_schedule.First(), master, pattern, _schedule.Last().End);
                master.Save(SendInvitationsMode.SendToAllAndSaveCopy);
                result = true;
            }

            return result;
        }

        /// <summary>
        /// Loads all appointments by subject.
        /// </summary>
        /// <returns>Appointments collection.</returns>
        private Appointment[] LoadAppointments()
        {
            var appointments = AppointmentHelper.GetAllAppointmentsBySubject(_service, Subject);
            LogAppointments(appointments);
            return appointments;
        }

        /// <summary>
        /// Validates parameters.
        /// </summary>
        private void ValidateParameters()
        {
            if (string.IsNullOrEmpty(Subject))
            {
                throw new ApplicationException("EventExchangeSync: Cannot sync event, Subject is Empty");
            }

            if (string.IsNullOrEmpty(Body))
            {
                Console.WriteLine("EventExchangeSync: Warning: Body is empty for Subject: {0}", Subject);
            }

            if (Data == null)
            {
                throw new ApplicationException($"EventExchangeSync: Data for sync is null. Subject: {Subject}");
            }

            if (Data.Rows.Count == 0)
            {
                Console.WriteLine("EventExchangeSync: Warning: Data.Rows count == 0. Subject: {0}", Subject);
            }

            if (RequiredAttendees == null || RequiredAttendees.Length == 0)
            {
                Console.WriteLine("EventExchangeSync: Warning: List of RequiredAttendees is empty. Subject: {0}", Subject);
            }
        }

        /// <summary>
        /// Initializes new appointment and send invitations.
        /// </summary>
        /// <param name="row">New appointment data.</param>
        /// <param name="appointment">New appointment instance.</param>
        /// <param name="weeklyPattern">Weekly pattern.</param>
        /// <param name="masterEndDate">Recurrence end date.</param>
        private void InitNewAppointment(ScheduleRow row, Appointment appointment, WeeklyPattern weeklyPattern = null, DateTime masterEndDate = default(DateTime))
        {
            appointment.Subject = Subject;
            appointment.Body = new MessageBody(BodyIsHtml ? BodyType.HTML : BodyType.Text, Body);
            
            appointment.Location = weeklyPattern != null ? weeklyPattern.Location : row.Location;
            appointment.Start = row.Start;
            appointment.End = row.End;

            if (weeklyPattern != null)
            {
                appointment.Start = row.Start.Date + weeklyPattern.StartTime;
                appointment.End = row.End.Date + weeklyPattern.EndTime;

                var pattern = new Recurrence.WeeklyPattern();
                pattern.DaysOfTheWeek.AddRange(weeklyPattern.DaysOfTheWeek);
                pattern.StartDate = row.Start.Date + weeklyPattern.StartTime;
                pattern.EndDate = masterEndDate.Date + weeklyPattern.EndTime;
                appointment.Recurrence = pattern;
            }

            IsAtendeesUpdated(appointment);
        }

        /// <summary>
        /// Synchronize local appointment with server and send updates.
        /// </summary>
        /// <param name="pattern">Data of occurrence.</param>
        /// <param name="master">Current appointment state.</param>
        /// <param name="sendMode">Send updates mode.</param>
        /// <returns>True if any changes, False otherwise.</returns>
        private bool SyncMasterItem(WeeklyPattern pattern, Appointment master, out SendInvitationsOrCancellationsMode sendMode)
        {
            var isPropertiesChanged = false;
            var isAttendeesChanged = false;

            if (pattern.Location != master.Location)
            {
                master.Location = pattern.Location;
                isPropertiesChanged = true;
            }

            if (IsBodyUpdated(master, BodyCheckSubstrings))
            {
                isPropertiesChanged = true;
            }

            if (IsAtendeesUpdated(master))
            {
                isAttendeesChanged = true;
            }

            if (isAttendeesChanged && !isPropertiesChanged)
            {
                sendMode = SendInvitationsOrCancellationsMode.SendToChangedAndSaveCopy;
            }
            else
            {
                sendMode = SendInvitationsOrCancellationsMode.SendToAllAndSaveCopy;
            }

            return isPropertiesChanged || isAttendeesChanged;
        }

        /// <summary>
        /// Synchronize local appointment with server and send updates.
        /// </summary>
        /// <param name="row">Data of occurrence.</param>
        /// <param name="appointment">Current appointment state.</param>
        /// <param name="sendMode">Send updates mode.</param>
        /// <returns>True if any changes, False otherwise.</returns>
        private bool SyncAppointmentItem(ScheduleRow row, Appointment appointment, out SendInvitationsOrCancellationsMode sendMode)
        {
            var propertiesChanged = false;
            var atendeeChanged = false;

            if (row.Location != appointment.Location)
            {
                appointment.Location = row.Location;
                propertiesChanged = true;
            }

            if (IsBodyUpdated(appointment, BodyCheckSubstrings))
            {
                propertiesChanged = true;
            }

            if (appointment.AppointmentType == AppointmentType.Single && IsAtendeesUpdated(appointment))
            {
                atendeeChanged = true;
            }

            if (row.Start != appointment.Start || row.End != appointment.End)
            {
                appointment.Start = row.Start;
                appointment.End = row.End;
                propertiesChanged = true;
            }

            if (atendeeChanged && !propertiesChanged)
            {
                sendMode = SendInvitationsOrCancellationsMode.SendToChangedAndSaveCopy;
            }
            else
            {
                sendMode = SendInvitationsOrCancellationsMode.SendToAllAndSaveCopy;
            }

            return propertiesChanged || atendeeChanged;
        }

        /// <summary>
        /// Indicates whether collections of attendees changed or not.
        /// </summary>
        /// <param name="appointment">Appointment to check.</param>
        /// <returns>True if any updates found, False otherwise.</returns>
        private bool IsAtendeesUpdated(Appointment appointment)
        {
            return AppointmentHelper.IsSyncAttendeesNeede(appointment, RequiredAttendees, OptionalAttendees);
        }

        /// <summary>
        /// Indicates whether body passed body checks or not.
        /// </summary>
        /// <param name="appointment">Appointment to check.</param>
        /// <param name="bodyCheckSubstrings">Substrings to check.</param>
        /// <returns>True if body passed body checks, False otherwise.</returns>
        private bool IsBodyUpdated(Item appointment, string[] bodyCheckSubstrings)
        {
            if (BodyIsHtml && bodyCheckSubstrings.Length > 0)
            {
                return bodyCheckSubstrings.All(substr => appointment.Body.Text.Contains(substr));
            }

            if (Body == appointment.Body.Text && !(BodyIsHtml ^ (appointment.Body.BodyType == BodyType.HTML)))
            {
                return false;
            }

            appointment.Body.Text = Body;
            appointment.Body.BodyType = BodyIsHtml ? BodyType.HTML : BodyType.Text;
            return true;
        }
    }
}

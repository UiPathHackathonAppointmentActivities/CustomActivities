// License placeholder

using System;
using Microsoft.Exchange.WebServices.Data;

namespace Epam.Activities.Exchange.Data.Models
{
    /// <summary>
    /// Model for daily recurrence pattern.
    /// </summary>
    public class DayPattern
    {
        /// <summary>
        /// Gets or sets day of the week.
        /// </summary>
        public DayOfTheWeek DayOfTheWeek { get; set; }

        /// <summary>
        /// Gets or sets start time of appointment.
        /// </summary>
        public TimeSpan StartTime { get; set; }

        /// <summary>
        /// Gets or sets end time of appointment.
        /// </summary>
        public TimeSpan EndTime { get; set; }

        /// <summary>
        /// Gets or sets amount of occurrences.
        /// </summary>
        public int Count { get; set; }

        /// <summary>
        /// Gets or sets location if any.
        /// </summary>
        public string Location { get; set; }

        /// <summary>
        /// Builds key in format 'DOW - HH:MM'. 
        /// </summary>
        /// <param name="date">Date time to build from.</param>
        /// <returns>E.g. "Monday - 11:00"</returns>
        public static string GetKey(DateTime date)
        {
            return $"{date.DayOfWeek} - {date.TimeOfDay:hh\\:mm}";
        }

        /// <summary>
        /// Initialize instance of <see cref="DayPattern"/>
        /// </summary>
        /// <param name="start">Start time of appointment.</param>
        /// <param name="end">End time of appointment.</param>
        /// <param name="location">Location of appointment.</param>
        /// <returns>Instance of <see cref="DayPattern"/></returns>
        public static DayPattern NewPattern(DateTime start, DateTime end, string location)
        {
            return new DayPattern
            {
                DayOfTheWeek = (DayOfTheWeek)start.DayOfWeek,
                StartTime = start.TimeOfDay,
                EndTime = end.TimeOfDay,
                Location = location,
                Count = 0
            };
        }

        /// <inheritdoc />
        public override string ToString()
        {
            return $"{DayOfTheWeek} - {StartTime:hh\\:mm}";
        }
    }
}
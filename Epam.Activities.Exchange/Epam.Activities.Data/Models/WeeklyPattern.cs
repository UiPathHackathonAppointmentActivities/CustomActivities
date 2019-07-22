// License placeholder

using System;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;

namespace Epam.Activities.Exchange.Data.Models
{
    /// <summary>
    /// Model for Weekly pattern.
    /// </summary>
    public class WeeklyPattern
    {
        /// <summary>
        /// Gets or sets days of week to determine recurrence.
        /// </summary>
        public DayOfTheWeek[] DaysOfTheWeek { get; set; }

        /// <summary>
        /// Gets or sets start time of each appointment.
        /// </summary>
        public TimeSpan StartTime { get; set; }

        /// <summary>
        /// Gets or sets end time of each appointment.
        /// </summary>
        public TimeSpan EndTime { get; set; }

        /// <summary>
        /// Gets or sets location of appointment.
        /// </summary>
        public string Location { get; set; }

        /// <inheritdoc />
        public override string ToString()
        {
            return $"{string.Join("/", DaysOfTheWeek.Select(day => day.ToString()))}, {StartTime:hh\\:mm}-{EndTime:hh\\:mm}";
        }
    }
}

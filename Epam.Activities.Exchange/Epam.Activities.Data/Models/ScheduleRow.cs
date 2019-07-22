// License placeholder

using System;

namespace Epam.Activities.Exchange.Data.Models
{
    /// <summary>
    /// Model for data to recognize recurrence.
    /// </summary>
    public class ScheduleRow
    {
        /// <summary>
        /// Gets or sets start time.
        /// </summary>
        public DateTime Start { get; set; }

        /// <summary>
        /// Gets or sets end time.
        /// </summary>
        public DateTime End { get; set; }

        /// <summary>
        /// Gets or sets location.
        /// </summary>
        public string Location { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether matched a pattern or not. 
        /// </summary>
        public bool Matched { get; set; }
    }
}
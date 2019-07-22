// License placeholder

using System;

namespace Epam.Activities.Exchange.Data.Models
{
    /// <summary>
    /// Model for notifications history.
    /// </summary>
    public class NotifyHistory
    {
        /// <summary>
        /// Gets or sets subject of notification.
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// Gets or sets date when notification was sent.
        /// </summary>
        public DateTime SentOn { get; set; }

        /// <summary>
        /// Gets or sets person who notification was sent to.
        /// </summary>
        public string SentTo { get; set; }
    }
}
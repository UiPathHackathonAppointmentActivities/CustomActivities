// License placeholder

using System.Activities;
using System.ComponentModel;

namespace Epam.Activities.Exchange.Appointments
{
    /// <inheritdoc />
    /// <summary>
    /// Abstract class for exchange activities.
    /// </summary>
    public abstract class ExchangeActivityBase : CodeActivity
    {
        /// <summary>
        /// Gets or sets exchange Service Url.
        /// </summary>
        [Category("Exchange Service")]
        [Description("Exchange service url")]
        public InArgument<string> ExchangeUrl { get; set; }

        /// <summary>
        /// Gets or sets service account email.
        /// </summary>
        [Category("Exchange Service")]
        [Description("Login of account that will be used for working with exchange")]
        [RequiredArgument]
        public InArgument<string> OrganizerEmail { get; set; }

        /// <summary>
        /// Gets or sets service account password.
        /// </summary>
        [Category("Exchange Service")]
        [Description("Password of account that will be used for working with exchange")]
        [RequiredArgument]
        public InArgument<string> OrganizerPassword { get; set; }
    }
}

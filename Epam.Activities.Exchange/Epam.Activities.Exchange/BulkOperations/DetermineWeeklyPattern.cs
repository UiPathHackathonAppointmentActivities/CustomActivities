// License placeholder

using System.Activities;
using System.ComponentModel;
using System.Data;
using Epam.Activities.Exchange.Data.Models;
using Epam.Activities.Exchange.Services;

namespace Epam.Activities.Exchange.BulkOperations
{
    /// <summary>
    /// Determines weekly pattern based on information about all occurrences.
    /// </summary>
    public class DetermineWeeklyPattern : CodeActivity
    {
        /// <summary>
        /// Gets or sets data with occurrences information.
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        [Description("Required columns: Start (of DateTime), End (of DateTime), Location (of string)")]
        public InArgument<DataTable> Data { get; set; }

        /// <summary>
        /// Gets or sets resolved patterns.
        /// </summary>
        [Category("Output")]
        public OutArgument<WeeklyPattern[]> WeeklyPatterns { get; set; }
        
        /// <inheritdoc />
        protected override void Execute(CodeActivityContext context)
        {
            var data = context.GetValue(Data);
            var schedule = AppointmentsSynchronizer.PrepareSchedule(data);
            var pattern = AppointmentsSynchronizer.DeterminePattern(schedule);

            context.SetValue(WeeklyPatterns, pattern);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using MeetingCostAlert.Properties;
using Microsoft.Office.Interop.Outlook;
using System.Globalization;
using System.Windows.Forms;

namespace MeetingCostAlert
{
    public partial class ThisAddIn
    {
        private int hourlyRate;
        private int trivialCostThreshold;
        private string locale;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Settings.Default.Upgrade();

            this.hourlyRate = Settings.Default.HourlyRate;
            this.trivialCostThreshold = Settings.Default.TrivialCostThreshold;
            this.locale = Settings.Default.Locale;

            Application.ItemSend += Application_ItemSend;
        }

        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            if (Item is MeetingItem meeting)
            {
                AppointmentItem appointment = meeting.GetAssociatedAppointment(false);
                RecurrencePattern recurrancePattern = appointment.GetRecurrencePattern();

                if (meeting.Class != OlObjectClass.olMeetingRequest)
                {
                    return;
                }

                int attendees = appointment.Recipients.Count;
                TimeSpan duration = TimeSpan.FromMinutes(appointment.Duration);

                double cost = duration.TotalHours * attendees * hourlyRate;
                double timeCost = duration.TotalHours * attendees;

                if (cost <= this.trivialCostThreshold)
                {
                    return;
                }

                string message = appointment.IsRecurring
                    ? string.Format(
                        CultureInfo.CreateSpecificCulture(this.locale),
                        "The cost of this recurring meeting is an estimated {1:C0} per occurrence and {3:C0} in total.\n" +
                        "Do you wish to continue?",
                        timeCost,
                        cost,
                        timeCost * recurrancePattern.Occurrences,
                        cost * recurrancePattern.Occurrences)
                    : string.Format(
                        CultureInfo.CreateSpecificCulture(this.locale),
                        "The cost of this meeting is an estimated {1:C0}\n" +
                        "Do you wish to continue?",
                        timeCost,
                        cost);

                DialogResult result = MessageBox.Show(
                    message,
                    "Meeting Cost",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                Cancel = result != DialogResult.Yes;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}

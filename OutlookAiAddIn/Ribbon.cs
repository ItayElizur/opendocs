using System.Runtime.InteropServices;
using OfficeAi.Shared;

namespace OutlookAiAddIn
{
    // The OpenDocs button goes on the main window's Mail tab (idMso "TabMail", not Word/Excel/PowerPoint's "TabHome"),
    // on the tab of a mail compose window, and on the appointment and meeting editor's tab. Windows that only read an
    // item (a received email, a meeting request) get no button. See ThisAddIn.cs.md.
    [ComVisible(true)]
    public class Ribbon : RibbonBase
    {
        protected override string TabIdMsoFor(string ribbonID)
        {
            switch (ribbonID)
            {
                case "Microsoft.Outlook.Explorer": return "TabMail";
                case "Microsoft.Outlook.Mail.Compose": return "TabNewMailMessage";
                case "Microsoft.Outlook.Appointment": return "TabAppointment";
                default: return null;
            }
        }

        protected override void TogglePane(object context)
        {
            Globals.ThisAddIn.TogglePane(context);
        }

        protected override void TogglePane()
        {
            Globals.ThisAddIn.TogglePane(null);
        }

        protected override int GetOfficeUiLanguageId()
        {
            return Globals.ThisAddIn.GetOfficeUiLanguageId();
        }
    }
}

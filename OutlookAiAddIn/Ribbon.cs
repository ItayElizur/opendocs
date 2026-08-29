using System.Runtime.InteropServices;
using OfficeAi.Shared;

namespace OutlookAiAddIn
{
    // Explorer-only: the Airchat button goes on the Explorer's Mail tab
    // (idMso "TabMail", not Word/Excel/PowerPoint's "TabHome") and is
    // suppressed on every Inspector (pop-out read/compose) ribbon surface.
    [ComVisible(true)]
    public class Ribbon : RibbonBase
    {
        protected override string HomeTabIdMso
        {
            get { return "TabMail"; }
        }

        protected override bool ProvidesRibbonFor(string ribbonID)
        {
            return ribbonID == "Microsoft.Outlook.Explorer";
        }

        protected override void TogglePane()
        {
            Globals.ThisAddIn.TogglePane();
        }
    }
}

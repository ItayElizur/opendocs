using System.Runtime.InteropServices;
using OfficeAi.Shared;

namespace WordAiAddIn
{
    [ComVisible(true)]
    public class Ribbon : RibbonBase
    {
        protected override void TogglePane()
        {
            Globals.ThisAddIn.TogglePane();
        }
    }
}

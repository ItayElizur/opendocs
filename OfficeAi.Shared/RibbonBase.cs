using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace OfficeAi.Shared
{
    // Shared across Word/Excel/PowerPoint's Ribbon.cs - the three copies
    // differed only by namespace. TogglePane() is the one app-specific hook
    // (Globals.ThisAddIn is VSTO-generated per project, so it can't be
    // referenced from here); each app's concrete Ribbon overrides it.
    public abstract class RibbonBase : IRibbonExtensibility
    {
        private IRibbonUI _ribbon;

        // Word/Excel/PowerPoint all inject into TabHome and answer every
        // ribbonID. Outlook's Explorer home tab is TabMail, and Outlook calls
        // GetCustomUI once per ribbon surface (Explorer, each Inspector class) -
        // OutlookAiAddIn overrides both hooks so the button appears only on the
        // Explorer's Mail tab and nowhere in the pop-out windows.
        protected virtual string HomeTabIdMso { get { return "TabHome"; } }

        protected virtual bool ProvidesRibbonFor(string ribbonID) { return true; }

        public string GetCustomUI(string ribbonID)
        {
            if (!ProvidesRibbonFor(ribbonID)) return null;
            return
@"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab idMso=""" + HomeTabIdMso + @""">
        <group id=""OpenDocsGroup"" getLabel=""GetGroupLabel"">
          <button id=""OpenDocsToggleButton"" getLabel=""GetButtonLabel"" size=""large""
                  getImage=""GetLogoImage"" onAction=""OnToggleTaskPane"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        // Office's UI display language decides the transliterated Hebrew vs
        // Latin brand name; getLabel callbacks are the standard Ribbon XML
        // mechanism for this. GetOfficeUiLanguageId() delegates to
        // Globals.ThisAddIn's own guarded copy rather than re-issuing the COM
        // call here. See RibbonBase.cs.md.
        public string GetGroupLabel(IRibbonControl control)
        {
            return OfficeLanguage.ResolveBrandName(GetOfficeUiLanguageId());
        }

        public string GetButtonLabel(IRibbonControl control)
        {
            return OfficeLanguage.ResolveBrandName(GetOfficeUiLanguageId());
        }

        protected abstract int GetOfficeUiLanguageId();

        // Reads the same web/logo.png the WebView2 header uses (copied there
        // at build time) - one image file drives both surfaces.
        // AppDomain.CurrentDomain.BaseDirectory resolves to the host
        // process's own directory even though this code lives in the shared
        // assembly - it's a property of the AppDomain, not of the reading
        // assembly. See RibbonBase.cs.md.
        public stdole.IPictureDisp GetLogoImage(IRibbonControl control)
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "web", "logo.png");
            using (Image image = Image.FromFile(path))
            {
                return PictureConverter.ImageToPictureDisp(image);
            }
        }

        public void OnToggleTaskPane(IRibbonControl control)
        {
            TogglePane();
        }

        protected abstract void TogglePane();
    }

    internal static class PictureConverter
    {
        private class AxHostConverter : AxHost
        {
            private AxHostConverter() : base(string.Empty) { }

            public static stdole.IPictureDisp Convert(Image image)
            {
                return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }
        }

        public static stdole.IPictureDisp ImageToPictureDisp(Image image)
        {
            return AxHostConverter.Convert(image);
        }
    }
}

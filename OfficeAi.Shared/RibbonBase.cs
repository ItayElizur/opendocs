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

        // The built-in tab the button goes on for one ribbon surface, or null for no button there. Word/Excel/PowerPoint
        // use TabHome everywhere. Outlook calls GetCustomUI once per surface (the main window, each kind of item window)
        // and overrides this to pick its tabs.
        protected virtual string TabIdMsoFor(string ribbonID) { return "TabHome"; }

        public string GetCustomUI(string ribbonID)
        {
            string tab = TabIdMsoFor(ribbonID);
            if (tab == null) return null;
            return
@"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab idMso=""" + tab + @""">
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
            TogglePane((object)control.Context);
        }

        // context is the window whose button was pressed (Outlook passes its Explorer or Inspector).
        protected virtual void TogglePane(object context) { TogglePane(); }

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

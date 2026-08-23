using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace ExcelAiAddIn
{
    [ComVisible(true)]
    public class Ribbon : IRibbonExtensibility
    {
        private IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            return
@"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab idMso=""TabHome"">
        <group id=""AirchatGroup"" label=""Airchat Office"">
          <button id=""AirchatToggleButton"" label=""Airchat Office"" size=""large""
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

        // Reads the same web/logo.png the WebView2-hosted header uses (copied
        // there at build time from shared/chat-ui/logo.png) - one physical
        // image file drives both surfaces, so editing it updates both.
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
            Globals.ThisAddIn.TogglePane();
        }
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

using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddIn1
{
    public class CustomRibbon : IRibbonExtensibility
    {
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PowerPointAddIn1.CustomRibbon.xml");
        }

        public void OnMyButtonClick(IRibbonControl control)
        {
            System.Windows.Forms.MessageBox.Show("Button clicked!");
        }

        // Helper method to load the XML from the embedded resource
        private static string GetResourceText(string resourceName)
        {
            var asm = typeof(CustomRibbon).Assembly;
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                using (var reader = new System.IO.StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }
    }
}

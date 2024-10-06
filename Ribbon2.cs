using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public partial class Ribbon2
    {
        private void Ribbon2_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnSplit_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application application = Globals.ThisAddIn.Application;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Drawing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane taskPaneValue;
        private Ribbon2 ribbon2;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.PresentationNewSlide +=
            new PowerPoint.EApplication_PresentationNewSlideEventHandler(PresentationNewSlide);
            ribbon2 = new Ribbon2();
        }

        void PresentationNewSlide(PowerPoint.Slide Sld)
        {
            Shape textBox = Sld.Shapes.AddTextbox(
            Office.MsoTextOrientation.msoTextOrientationHorizontal, 10, 20, 500,50);
            textBox.TextFrame.TextRange.InsertAfter("This text is written by Thien");
            Sld.FollowMasterBackground = Office.MsoTriState.msoFalse;
            //Sld.Background.Fill.TwoColorGradient(Office.MsoGradientStyle.msoGradientFromCenter, 2);
            //Sld.Background.Fill.GradientAngle = 90;
            //Sld.Background.Fill.GradientStops.Insert(ColorTranslator.ToOle(Color.LightBlue), 0, 0);
            //Sld.Background.Fill.GradientStops.Insert(ColorTranslator.ToOle(Color.AliceBlue), 0.5f, 0);
            //Sld.Background.Fill.BackColor.RGB = ColorTranslator.ToOle(Color.LightBlue);
            //Sld.Background.Fill.Solid();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

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

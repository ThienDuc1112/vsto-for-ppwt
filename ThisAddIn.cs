using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Drawing;
using Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Tools;

namespace PowerPointAddIn1
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane taskPaneValue;
        private Ribbon2 ribbon2;
        private Office.CommandBarButton customButton;
        private Office.CommandBar textCommandBar;
        private Office.CommandBarControl customMenu;
        private CustomTaskPane myTaskPane;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //this.Application.PresentationNewSlide +=
            //new PowerPoint.EApplication_PresentationNewSlideEventHandler(PresentationNewSlide);
            //Application.SlideSelectionChanged += Application_SlideSelectionChanged;
            // this.Application.WindowBeforeRightClick += new PowerPoint.EApplication_WindowBeforeRightClickEventHandler(AddCustomContextMenuItem);
            UserControl1 myPaneControl = new UserControl1();

            // Add a custom task pane to the PowerPoint application
            myTaskPane = CustomTaskPanes.Add(myPaneControl, "My Custom Task Pane");

            // Set the dock position to the right side
            myTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;

            // Set the width of the task pane (e.g., 300 pixels)
            myTaskPane.Width = 500;

            // Make the task pane visible
            myTaskPane.Visible = true;
        }

        private void Application_SlideSelectionChanged(PowerPoint.SlideRange slideRange)
        {
            // Clear existing context menu if it exists
            if (customMenu != null)
            {
                customMenu.Delete();
                customMenu = null;
            }
            foreach (PowerPoint.Slide selectedSlide in slideRange)
            {
                // Check if there is selected text
                if (selectedSlide.Shapes[1].HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    var textFrame = selectedSlide.Shapes[1].TextFrame;
                    var textRange = textFrame.TextRange;

                    // Create the custom context menu
                    textCommandBar = Application.CommandBars["Text"];
                    customMenu = textCommandBar.Controls.Add(Office.MsoControlType.msoControlPopup, Type.Missing, Type.Missing, 1, true);
                    customMenu.Caption = "Demo"; // Set the caption for the submenu

                    // Add "Clean Up" button
                    Office.CommandBarControl cleanUpButton = customMenu.Control.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, 1, true);
                    cleanUpButton.Caption = "Clean Up";
                    cleanUpButton.OnAction = "CustomButton_Click2";


                }
            }
        }

            private void AddCustomContextMenuItem(PowerPoint.Selection sel, ref bool cance)
        {
                // Get the context menu (CommandBar) for text selection
                Office.CommandBar contextMenu = Application.CommandBars["Text"];

            // Add a new button to the context menu
            customButton = (Office.CommandBarButton)contextMenu.Controls.Add(
                    Office.MsoControlType.msoControlButton,
                    missing, missing, missing, true);

                // Set properties for the button
                customButton.Caption = "Custom Action";
                customButton.Tag = "CustomButton";
                customButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(CustomButton_Click2);
            
        }

        // Handler for the custom button click event
        private void CustomButton_Click2(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            MessageBox.Show("Custom button clicked!");
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new CustomRibbon();
        }



        private void ShowCustomPopup(string selectedText)
        {
            // Create a new form (or context menu) with a custom button
            Form popupForm = new Form();
            popupForm.StartPosition = FormStartPosition.Manual;
            popupForm.Location = Cursor.Position;
            popupForm.Width = 200;
            popupForm.Height = 100;

            Label label = new Label();
            label.Text = $"Selected Text: {selectedText}";
            label.Dock = DockStyle.Top;
            popupForm.Controls.Add(label);

            Button customButton = new Button();
            customButton.Text = "My Custom Button";
            customButton.Dock = DockStyle.Bottom;
            customButton.Click += (s, e) => { MessageBox.Show("Button Clicked!"); };
            popupForm.Controls.Add(customButton);

            // Show the custom popup
            popupForm.ShowDialog();
        }



        void PresentationNewSlide(PowerPoint.Slide Sld)
        {


            //CreateNewChartInExcel();
            //Shape textBox = Sld.Shapes.AddTextbox(
            //Office.MsoTextOrientation.msoTextOrientationHorizontal, 10, 20, 500,50);
            //textBox.TextFrame.TextRange.InsertAfter("This text is written by Thien");
            //Sld.FollowMasterBackground = Office.MsoTriState.msoFalse;
            //Sld.Background.Fill.TwoColorGradient(Office.MsoGradientStyle.msoGradientFromCenter, 2);
            //Sld.Background.Fill.GradientAngle = 90;
            //Sld.Background.Fill.GradientStops.Insert(ColorTranslator.ToOle(Color.LightBlue), 0, 0);
            //Sld.Background.Fill.GradientStops.Insert(ColorTranslator.ToOle(Color.AliceBlue), 0.5f, 0);
            //Sld.Background.Fill.BackColor.RGB = ColorTranslator.ToOle(Color.LightBlue);
            //Sld.Background.Fill.Solid();
        }
        private void AddCustomButtonToContextMenu(PowerPoint.Selection sel, ref bool cancel)
        {
           
            try
            {
                Microsoft.Office.Core.CommandBar contextMenu = null;

                // Check what kind of item was right-clicked (text, shape, etc.)
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    // Right-clicked on text
                    contextMenu = this.Application.CommandBars["Text"];
                }
                else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    // Right-clicked on a shape
                    contextMenu = this.Application.CommandBars["Shape"];
                }

                if (contextMenu != null)
                {
                    // Check if the custom button already exists to prevent duplicates
                    Microsoft.Office.Core.CommandBarButton customButton = (Microsoft.Office.Core.CommandBarButton)
                        contextMenu.FindControl(Microsoft.Office.Core.MsoControlType.msoControlButton, 0, "MYRIGHTCLICKMENU", Missing.Value, Missing.Value);

                    if (customButton == null)
                    {
                        // Add the custom button to the context menu
                        customButton = (Microsoft.Office.Core.CommandBarButton)contextMenu.Controls.Add(
                            Microsoft.Office.Core.MsoControlType.msoControlButton, Missing.Value, Missing.Value, contextMenu.Controls.Count + 1, true);

                        // Set the button's properties
                        customButton.Caption = "My Custom Button";
                        customButton.Tag = "MYRIGHTCLICKMENU";
                        customButton.BeginGroup = true; // Adds a separator line above the button

                        // Handle the button click event
                        customButton.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(CustomButton_Click);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message);
            }
        }


        private void CustomButton_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            // Define what happens when the custom button is clicked
            System.Windows.Forms.MessageBox.Show("Custom button clicked!");
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

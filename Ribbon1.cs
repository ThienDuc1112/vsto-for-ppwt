using Microsoft.Office.Tools.Ribbon;
using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.IO;
using System.Diagnostics;
using System.Net.Http;
using System.Drawing;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.PowerPoint;
using XlChartType = Microsoft.Office.Core.XlChartType;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.Remoting;

namespace PowerPointAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnInsertNewSlide_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Presentation presentation = pptApp.ActivePresentation;
            PowerPoint.Slide newSlide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);

            PowerPoint.TextRange title = newSlide.Shapes[1].TextFrame.TextRange;
            title.Text = "Welcome to my presentation";
            //set bold
            title.Font.Bold = MsoTriState.msoTrue;
            title.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);
            title.Font.Size = 32;

            PowerPoint.TextRange bodyTextRange = newSlide.Shapes[2].TextFrame.TextRange;
            bodyTextRange.Text = "This slide is created by a dev";

            // Set font color, size, italic for body text
            bodyTextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
            bodyTextRange.Font.Size = 20;
            bodyTextRange.Font.Italic = MsoTriState.msoTrue;
        }

        private void btnCopySlide_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Presentation presentation = pptApp.ActivePresentation;
            PowerPoint.SlideRange selectedSlides = pptApp.ActiveWindow.Selection.SlideRange;

            if (selectedSlides != null & selectedSlides.Count > 0)
            {
                PowerPoint.SlideRange duplicatedSlide = selectedSlides.Duplicate();
                duplicatedSlide.MoveTo(selectedSlides.SlideIndex + selectedSlides.Count);
            }
        }

        private void btnAutomate_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Presentation presentation = pptApp.ActivePresentation;

            if (presentation.SlideShowSettings.AdvanceMode == PowerPoint.PpSlideShowAdvanceMode.ppSlideShowUseSlideTimings)
            {
                MessageBox.Show("This show is already set to advance automatically");
            }
            else
            {
                presentation.SlideShowSettings.AdvanceMode = PowerPoint.PpSlideShowAdvanceMode.ppSlideShowUseSlideTimings;
            }

            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                if (slide.SlideShowTransition.AdvanceOnTime != Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    slide.SlideShowTransition.AdvanceOnTime = Microsoft.Office.Core.MsoTriState.msoTrue;
                    slide.SlideShowTransition.AdvanceTime = 3;
                }
            }

            presentation.SlideShowSettings.Run();
        }

        private void btnAddNote_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.SlideRange currentSlide = pptApp.ActiveWindow.Selection.SlideRange;

            if (currentSlide != null)
            {
                using (noteForm noteForm = new noteForm())
                {
                    if (noteForm.ShowDialog() == DialogResult.OK)
                    {
                        string speakerNote = noteForm.Note;

                        foreach (PowerPoint.Slide slide in currentSlide)
                        {
                            slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = speakerNote;
                        }
                        MessageBox.Show("Speaker notes added to the selected slide(s)!", "Success");
                    }
                }
            }
        }

        private void btnExport_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Presentation presentation = app.ActivePresentation;

            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "File PDF (*.pdf)|.*pdf";
                saveFileDialog.Title = "Export to PDF";
                saveFileDialog.FileName = presentation.Name.Replace(".pptx", ".pdf");

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string pdfPath = saveFileDialog.FileName;

                    presentation.ExportAsFixedFormat(pdfPath,
                        PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                        PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen,
                        Office.MsoTriState.msoFalse,
                        PowerPoint.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                        PowerPoint.PpPrintOutputType.ppPrintOutputSlides,
                        MsoTriState.msoFalse, null,
                        PowerPoint.PpPrintRangeType.ppPrintAll,
                        string.Empty, true, true, true, true, false, Type.Missing
                    );

                    if (File.Exists(pdfPath))
                    {
                        Process.Start(pdfPath);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnImportImages_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Presentation presentation = app.ActivePresentation;

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Image Files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp",
                Title = "Select multiple images",
                Multiselect = true,
            };

            try
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    foreach (string imagePath in openFileDialog.FileNames)
                    {
                        PowerPoint.Slide slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);

                        PowerPoint.Shape insertedImage = slide.Shapes.AddPicture(
                            FileName: imagePath,
                            LinkToFile: MsoTriState.msoFalse,
                            SaveWithDocument: MsoTriState.msoTrue,
                            Left: 50, Top: 50,
                            Width: slide.CustomLayout.Width - 100,
                            Height: slide.CustomLayout.Height - 100);

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnApplyEffect_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Presentation presentation = app.ActivePresentation;

            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                slide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectWipeRight;
                slide.SlideShowTransition.Duration = 2.0f;

                slide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoTrue;
                slide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoFalse;
            }
        }

        private void btnSaveSlide_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation presentation = app.ActivePresentation;
                PowerPoint.SlideRange selectedSlides = app.ActiveWindow.Selection.SlideRange;

                if (selectedSlides != null && selectedSlides.Count > 0)
                {
                    PowerPoint.Presentation newPresentation = app.Presentations.Add(Office.MsoTriState.msoTrue);
                    foreach (PowerPoint.Slide slide in selectedSlides)
                    {
                        slide.Copy();
                        newPresentation.Slides.Paste();
                    }

                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "New Slide (*.pptx)|*.pptx";
                    saveFileDialog.Title = "Save as a new slide";
                    saveFileDialog.FileName = "Exported" + selectedSlides.Name + ".pptx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string newPath = saveFileDialog.FileName;
                        newPresentation.SaveAs(newPath, PowerPoint.PpSaveAsFileType.ppSaveAsPresentation, MsoTriState.msoTrue);
                    }
                    else
                    {
                        newPresentation.Close();
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnTableTitle_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation presentation = app.ActivePresentation;

                if (presentation.Slides.Count == 0)
                {
                    MessageBox.Show("The presentation has no slides.", "No Slides", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                PowerPoint.Slide tocSlide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutText);
                tocSlide.Shapes[1].TextFrame.TextRange.Text = "Table of Contents";

                int row = presentation.Slides.Count - 1;
                int columns = 2;
                PowerPoint.Shape tableShape = tocSlide.Shapes.AddTable(row, columns, 30, 30, 400, 500);
                PowerPoint.Table table = tableShape.Table;

                for (int i = 2; i <= presentation.Slides.Count; i++)
                {
                    PowerPoint.Slide slide = presentation.Slides[i];
                    string slideTitle = slide.Shapes[1].TextFrame.TextRange.Text;

                    if (!string.IsNullOrEmpty(slideTitle))
                    {
                        table.Cell(i, 2).Shape.TextFrame.TextRange.Text = $"{i - 1}. {slideTitle}";
                    }
                    else
                    {
                        table.Cell(i, 2).Shape.TextFrame.TextRange.Text = $"{i - 1}. [Untitled Slide]";
                    }

                    for (int column = 1; column <= columns; column++)
                    {
                        PowerPoint.Cell cell = table.Cell(i, column);
                        PowerPoint.TextRange textRange = cell.Shape.TextFrame.TextRange;
                        textRange.Font.Size = 10;
                        textRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                        textRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkOrange);
                        textRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;

                        cell.Shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);
                        cell.Borders[PowerPoint.PpBorderType.ppBorderTop].ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkOrange);
                        cell.Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkOrange);
                        cell.Borders[PowerPoint.PpBorderType.ppBorderLeft].ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkOrange);
                        cell.Borders[PowerPoint.PpBorderType.ppBorderRight].ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkOrange);
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnSearch_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation presentation = app.ActivePresentation;
                PowerPoint.SlideRange selectedSlides = app.ActiveWindow.Selection.SlideRange;
                string text = String.Empty;
                int foundCount = 0;

                using (noteForm noteForm = new noteForm())
                {
                    if (noteForm.ShowDialog() == DialogResult.OK)
                    {
                        text = noteForm.Note.Trim().Replace(" ", "");
                    }
                }

                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            PowerPoint.TextRange textRange = shape.TextFrame.TextRange;
                            int startIndex = 1; // PowerPoint TextRange is 1-based
                            while (startIndex <= textRange.Length)
                            {
                                int foundIndex = textRange.Text.IndexOf(text, startIndex - 1, StringComparison.OrdinalIgnoreCase);
                                if (foundIndex == -1)
                                {
                                    startIndex = textRange.Length + 1;
                                }
                                else
                                {
                                    PowerPoint.TextRange foundText = textRange.Characters(foundIndex + 1, text.Length);
                                    foundText.Font.Color.RGB = ColorTranslator.ToOle(Color.Yellow);
                                    startIndex = foundIndex + text.Length + 1;
                                    foundCount++;
                                }
                            }
                        }
                    }
                }

                if (foundCount > 0)
                {
                    MessageBox.Show($"There is found {foundCount} occurences of {text} ");
                }
                else
                {
                    MessageBox.Show($"No occurrences of '{text}' found.", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private async void btnCallApi_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Presentation presentation = app.ActivePresentation;
            PowerPoint.Slide tocSlide = null;

            if (presentation.Slides[1].Shapes[1].HasTextFrame == MsoTriState.msoTrue && presentation.Slides[1].Shapes[1].TextFrame.HasText == MsoTriState.msoTrue
                && presentation.Slides[1].Shapes[1].TextFrame.TextRange.Text.Contains("Table of Weather"))
            {
                tocSlide = presentation.Slides[1];
            }
            else
            {
                tocSlide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutText);
                tocSlide.Shapes[1].TextFrame.TextRange.Text = "Table of Weather";
            }
            int rows = 5;
            int columns = 2;
            float slideWidth = tocSlide.Master.Width;
            float slideHeight = tocSlide.Master.Height;
            PowerPoint.Shape tableShape = tocSlide.Shapes.AddTable(rows, columns, 30, 30, 400, 250);
            PowerPoint.Table table = tableShape.Table;
            table.Parent.Left = (slideWidth - 400) / 2;
            table.Parent.Top = (slideHeight - 250) / 2;

            string cityName = ddCity.SelectedItem.Label;

            dynamic res = await getWeather(cityName);
            try
            {
                table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "City";
                table.Cell(1, 2).Shape.TextFrame.TextRange.Text = res.location.name;
                table.Cell(2, 1).Shape.TextFrame.TextRange.Text = "Country";
                table.Cell(2, 2).Shape.TextFrame.TextRange.Text = res.location.country;
                table.Cell(3, 1).Shape.TextFrame.TextRange.Text = "Temperature(C)";
                table.Cell(3, 2).Shape.TextFrame.TextRange.Text = res.current.temp_c;
                table.Cell(4, 1).Shape.TextFrame.TextRange.Text = "Temperature(F)";
                table.Cell(4, 2).Shape.TextFrame.TextRange.Text = res.current.temp_f;
                table.Cell(5, 1).Shape.TextFrame.TextRange.Text = "Humidity";
                table.Cell(5, 2).Shape.TextFrame.TextRange.Text = res.current.humidity;

                for (int i = 1; i <= rows; i++)
                {
                    for (int j = 1; j <= columns; j++)
                    {
                        PowerPoint.Cell cell = table.Cell(i, j);
                        PowerPoint.TextRange textRange = cell.Shape.TextFrame.TextRange;
                        textRange.Font.Size = 16;
                        textRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                        textRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkOrange);
                        textRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                        textRange.Font.Name = "Times New Roman";
                        cell.Shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                    }
                }

                for (int j = 1; j <= columns; j++)
                {
                    table.Columns[j].Width = 200;
                }

                for (int i = 1; i <= rows; i++)
                {
                    table.Rows[i].Height = 50;
                }

                tocSlide.Select();
            }
            catch (RuntimeBinderException ex)
            {
                MessageBox.Show($"Error accessing property: {ex.Message}");
            }

        }

        private async Task<dynamic> getWeather(string cityName)
        {
            using (var client = new HttpClient())
            {
                var request = new HttpRequestMessage
                {
                    Method = HttpMethod.Get,
                    RequestUri = new Uri($"https://weatherapi-com.p.rapidapi.com/current.json?q={Uri.EscapeDataString(cityName)}"),
                    Headers =
                {
                    { "x-rapidapi-key", "4c25a72f26mshe5d9266689291b0p148052jsn22091d6ae3c2" },
                    { "x-rapidapi-host", "weatherapi-com.p.rapidapi.com" },
                },
                };

                using (var response = await client.SendAsync(request))
                {
                    response.EnsureSuccessStatusCode();
                    dynamic body = JsonConvert.DeserializeObject(await response.Content.ReadAsStringAsync());
                    return body;
                }
            }
        }

        private void btnFormatImage_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Presentation presentation = app.ActivePresentation;
            string imagePath = "C:\\Users\\Admin\\Pictures\\Screenshots\\test2.png";

            PowerPoint.Slide slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);

            //PowerPoint.Shape insertedImage = slide.Shapes.AddPicture(
            //    FileName: imagePath,
            //    LinkToFile: MsoTriState.msoFalse,
            //    SaveWithDocument: MsoTriState.msoTrue,
            //    Left: 50, Top: 50,
            //    Width: slide.CustomLayout.Width - 200,
            //    Height: slide.CustomLayout.Height - 200);
            PowerPoint.Shape insertedImage = slide.Shapes.AddShape(
                Office.MsoAutoShapeType.msoShapeRectangle,
                100, 100, 300, 200);

            insertedImage.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Aqua);
            insertedImage.Line.Weight = 3;
            insertedImage.Fill.TwoColorGradient(Office.MsoGradientStyle.msoGradientHorizontal, 2);
            insertedImage.Fill.GradientAngle = 90;
            insertedImage.Fill.GradientStops.Insert(
                ColorTranslator.ToOle(Color.Green), 0, 0.5f
                );
            insertedImage.Fill.GradientStops.Insert(
                ColorTranslator.ToOle(Color.Red), 0.5f, 0.5f
                );
            insertedImage.Fill.GradientStops.Insert(
            ColorTranslator.ToOle(Color.DarkBlue),
            0.8f,
            0);
        }

        private void btnTriggleVisibility_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.SlideRange slide = app.ActiveWindow.Selection.SlideRange;

            if (slide.Count > 0)
            {
                foreach (PowerPoint.Slide item in slide)
                {
                    foreach (PowerPoint.Shape shape in item.Shapes)
                    {
                        if (shape.Visible == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoFalse)
                        {
                            //shape.Delete();
                            shape.Visible = MsoTriState.msoFalse;
                        }
                        else
                        {
                            shape.Visible = MsoTriState.msoTrue;
                        }
                    }
                }
            }
        }

        private void btnAddText_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Presentation presentation = app.ActivePresentation;
            PowerPoint.Slide slide = presentation.Slides[1];
            string imagePath = "D:\\studying\\self-learning\\vsto\\assets\\icons8-checkbox-48.png";

            PowerPoint.Shape shape = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                140, 100, 500, 300);

            string[] listItems = new string[] {
                    "Command-line Unix",
                    "Vim",
                    "HTML",
                    "CSS",
                    "Python",
                    "JavaScript",
                    "SQL"
                    };

            PowerPoint.TextRange textRange = shape.TextFrame.TextRange;
            foreach (string item in listItems)
            {

                PowerPoint.TextRange paragraph = textRange.Paragraphs(textRange.Paragraphs().Count + 1);
                paragraph.Text = item + "\r";
                paragraph.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                paragraph.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);
                //paragraph.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNumbered;
                //paragraph.ParagraphFormat.Bullet.Style = PowerPoint.PpNumberedBulletStyle.ppBulletArabicPeriod;
                paragraph.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletPicture;
                if (System.IO.File.Exists(imagePath))
                {
                    paragraph.ParagraphFormat.Bullet.Picture(imagePath);
                }
                else
                {
                    MessageBox.Show("Not found");
                }
            }
        }

        private void btnChangeFS_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            Slide slide = app.ActiveWindow.Selection.SlideRange[1];

            PowerPoint.Shape shape = slide.Shapes[1];
            string originalText = shape.TextFrame.TextRange.Text;

            string[] wordsToStyle = new string[] { "bold", "red", "underlined" };
            //string[] splitArray = originalText.Split(new string[] { "bold", "red", "underlined" }, StringSplitOptions.None);

            PowerPoint.TextRange textRange = shape.TextFrame.TextRange;

            foreach (string word in wordsToStyle)
            {
                ChangeStyle(textRange, originalText, word);
            }
        }

        private void ChangeStyle(PowerPoint.TextRange textRange, string originalText, string targetWord)
        {
            int startIndex = originalText.IndexOf(targetWord);
            if (startIndex >= 0)
            {
                int wordLength = targetWord.Length;

                TextRange specificText = textRange.Characters(startIndex + 1, wordLength);
                switch (targetWord)
                {
                    case "bold":
                        specificText.Font.Bold = MsoTriState.msoTrue;
                        break;
                    case "red":
                        specificText.Font.Color.RGB = ColorTranslator.ToOle(Color.Red);
                        break;
                    case "underlined":
                        specificText.Font.Underline = MsoTriState.msoTrue;
                        break;
                }

                specificText.Font.Size = 20;
            }
        }

        private void btnSplit_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application application = Globals.ThisAddIn.Application;
            Slide slide = application.ActivePresentation.Slides[1];
            PowerPoint.Shape shape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 150, 100, 400, 200);
            shape.TextFrame.TextRange.Text = "This is a test paragraph. " +
            "We are splitting this text into two columns. " +
            "Each column will have equal text.";

            shape.TextFrame2.Column.Number = 2;
            shape.TextFrame2.Column.Spacing = 25f;
            shape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
            shape.Line.Weight = 5;

        }

        private void btnApplyBg_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application application = Globals.ThisAddIn.Application;
            Slide slide = application.ActiveWindow.Selection.SlideRange[1];

            string selectedBg = ddBackgroundSelecting.SelectedItem.Label.ToString();

            switch (selectedBg)
            {
                case "Horizontal Gradient":
                    slide.Background.Fill.TwoColorGradient(MsoGradientStyle.msoGradientHorizontal, 1);
                    slide.Background.Fill.GradientStops[1].Color.RGB = ColorTranslator.ToOle(Color.Blue);
                    slide.Background.Fill.GradientStops[2].Color.RGB = ColorTranslator.ToOle(Color.LightBlue);
                    break;

                case "Vertical Gradient":
                    slide.Background.Fill.TwoColorGradient(MsoGradientStyle.msoGradientVertical, 2);
                    slide.Background.Fill.GradientStops[1].Color.RGB = ColorTranslator.ToOle(Color.Red);
                    slide.Background.Fill.GradientStops[2].Color.RGB = ColorTranslator.ToOle(Color.Yellow);
                    break;
                case "Diagonal Gradient":
                    slide.Background.Fill.TwoColorGradient(MsoGradientStyle.msoGradientDiagonalUp, 2);
                    slide.Background.Fill.GradientStops[1].Color.RGB = ColorTranslator.ToOle(Color.Green);
                    slide.Background.Fill.GradientStops[2].Color.RGB = ColorTranslator.ToOle(Color.LightGreen);
                    break;
                case "Rectangular Gradient":
                    slide.Background.Fill.TwoColorGradient(MsoGradientStyle.msoGradientFromCorner, 2);
                    slide.Background.Fill.GradientStops[1].Color.RGB = ColorTranslator.ToOle(Color.Purple);
                    slide.Background.Fill.GradientStops[2].Color.RGB = ColorTranslator.ToOle(Color.MediumPurple);
                    break;
                case "Path Gradient":
                    slide.Background.Fill.TwoColorGradient(MsoGradientStyle.msoGradientFromTitle, 2);
                    slide.Background.Fill.GradientStops[1].Color.RGB = ColorTranslator.ToOle(Color.Orange);
                    slide.Background.Fill.GradientStops[2].Color.RGB = ColorTranslator.ToOle(Color.DarkOrange);
                    break;
                case "Center Gradient":
                    slide.Background.Fill.TwoColorGradient(MsoGradientStyle.msoGradientFromCenter, 2);
                    slide.Background.Fill.GradientStops.Insert(ColorTranslator.ToOle(Color.LightBlue), 0, 0);
                    slide.Background.Fill.GradientStops.Insert(ColorTranslator.ToOle(Color.AliceBlue), 0.5f, 0.9f);
                    break;
                default:
                    break;
            }
        }

        private void btnRemoveSlide_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            Slide slide = app.ActiveWindow.Selection.SlideRange[1];
            Slide previousSlide = null;

            if (slide.SlideIndex > 1)
            {
                previousSlide = app.ActivePresentation.Slides[slide.SlideIndex - 1];
            }
            if (previousSlide != null)
            {
                slide.Delete();
                app.ActiveWindow.View.GotoSlide(previousSlide.SlideIndex);
            }
        }

        //private void btnReverse_Click(object sender, RibbonControlEventArgs e)
        //{
        //    PowerPoint.Application app = Globals.ThisAddIn.Application;
        //    PowerPoint.Presentation presentation = app.ActivePresentation;
        //    PowerPoint.Presentation newPresentation = app.Presentations.Add();
        //    string currentPath = presentation.Path;
        //    int lastIndex = presentation.Slides.Count;
        //    string clipboardError;

        //    for (int i = lastIndex; i >= 1; i--)
        //    {
        //        PowerPoint.Slide sourceSlide = presentation.Slides[i];

        //        // Add a new blank slide in the new presentation
        //        PowerPoint.CustomLayout layout = newPresentation.SlideMaster.CustomLayouts[1];
        //        PowerPoint.Slide newSlide = newPresentation.Slides.AddSlide(newPresentation.Slides.Count + 1, layout);

        //        foreach (PowerPoint.Shape shape in sourceSlide.Shapes)
        //        {
        //            try
        //            {
        //                shape.Copy();
        //                if (IsClipboardDataValidForPowerPoint(out clipboardError))
        //                {
        //                    newSlide.Shapes.Paste(); 
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                MessageBox.Show("Error during copy/paste: " + ex.Message);
        //            }
        //        }
        //    }

        //    string newPath = System.IO.Path.Combine(currentPath, "newPresentation.pptx");
        //    try
        //    {
        //        newPresentation.SaveAs(newPath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Office.MsoTriState.msoTrue);
        //        MessageBox.Show("Presentation saved as " + newPath);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error saving presentation: " + ex.Message);
        //    }
        //}

        private void btnReverse_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Presentation presentation = app.ActivePresentation;
            PowerPoint.Presentation newPresentation = app.Presentations.Add();
            string currentPath = presentation.Path;
            int lastIndex = presentation.Slides.Count;
            string clipboardError;

            // Loop through the slides in reverse order
            for (int i = lastIndex; i >= 1; i--) // Use 1-based index for PowerPoint slides
            {
                presentation.Slides[i].Copy();
                if (IsClipboardDataValidForPowerPoint(out clipboardError))
                {
                    newPresentation.Slides.Paste(newPresentation.Slides.Count + 1);
                }
            }

            string newPath = System.IO.Path.Combine(currentPath, "newPresentation.pptx");

            try
            {
                // Save the new presentation
                newPresentation.SaveAs(newPath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Office.MsoTriState.msoTrue);
                MessageBox.Show("Presentation saved as " + newPath);
            }
            catch (Exception ex)
            {
                // Handle exceptions (e.g., issues with file saving)
                MessageBox.Show("Error: " + ex.Message);
            }
        }



        public bool IsClipboardDataValidForPowerPoint(out string errorMessage)
        {
            errorMessage = null;

            try
            {
                // Check if the clipboard contains shapes or slide data
                if (Clipboard.ContainsData(DataFormats.EnhancedMetafile) || Clipboard.ContainsData(DataFormats.Bitmap) || Clipboard.ContainsData(DataFormats.Text))
                {
                    return true; // Clipboard contains valid data
                }
                else
                {
                    errorMessage = "Clipboard does not contain valid PowerPoint data.";
                    return false;
                }
            }
            catch (Exception ex)
            {
                errorMessage = "Error checking clipboard: " + ex.Message;
                return false;
            }
        }

        private void btnSplitPP_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application application = Globals.ThisAddIn.Application;
            PowerPoint.Presentation presentation = application.ActivePresentation;
            string currentPath = presentation.Path;
            int totalSlide = presentation.Slides.Count;
            string outputFolder = System.IO.Path.Combine(currentPath, "newPresentations");
            string clipboardError;

            for (int i = 1; i <= totalSlide; i++)
            {
                PowerPoint.Presentation newPresentation = application.Presentations.Add();
                presentation.Slides[i].Copy();
                if (IsClipboardDataValidForPowerPoint(out clipboardError))
                {
                    newPresentation.Slides.Paste(newPresentation.Slides.Count + 1);
                }

                string newPath = System.IO.Path.Combine(outputFolder, $"Slide_{i}.pptx");
                try
                {
                    newPresentation.SaveAs(newPath, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                    newPresentation.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error saving presentation: {ex.Message}");
                }
            }
        }

        private void btnCombinePp_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            Presentation mainPresentation = app.ActivePresentation;

            string filePath1 = "";
            string filePath2 = "";

            using (combineFilesForm combineFileForm = new combineFilesForm())
            {
                if (combineFileForm.ShowDialog() == DialogResult.OK)
                {
                    filePath1 = combineFileForm.filePath1.Trim();
                    filePath2 = combineFileForm.filePath2.Trim();
                }
                else
                {
                    return;
                }
            }
            try
            {
                DuplicateSlidesFromPresentation(app, mainPresentation, filePath1);
                DuplicateSlidesFromPresentation(app, mainPresentation, filePath2);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }

        }

        private void DuplicateSlidesFromPresentation(PowerPoint.Application app, PowerPoint.Presentation mainPresentation, string filePath)
        {
            PowerPoint.Presentation sourcePresentation = null;

            try
            {
                sourcePresentation = app.Presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                for (int i = 1; i <= sourcePresentation.Slides.Count; i += 2)
                {
                    // Add a new blank slide to the destination presentation
                    PowerPoint.CustomLayout layout = mainPresentation.SlideMaster.CustomLayouts[1];
                    PowerPoint.Slide newSlide = mainPresentation.Slides.AddSlide(mainPresentation.Slides.Count + 1, layout);

                    // Copy content from the current slide
                    PowerPoint.Slide slide1 = sourcePresentation.Slides[i];
                    CopySlideContent(slide1, newSlide);

                    // Check if there is a next slide to combine (in case of an odd number of slides)
                    if (i + 1 <= sourcePresentation.Slides.Count)
                    {
                        PowerPoint.Slide slide2 = sourcePresentation.Slides[i + 1];
                        CopySlideContent(slide2, newSlide);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while copying slides: {ex.Message}");
            }
            finally
            {
                // Ensure the source presentation is always closed
                if (sourcePresentation != null)
                {
                    sourcePresentation.Close();
                }
            }
        }

        private void CopySlideContent(PowerPoint.Slide sourceSlide, PowerPoint.Slide targetSlide)
        {
            foreach (PowerPoint.Shape shape in sourceSlide.Shapes)
            {
                shape.Copy(); // Copy each shape
                targetSlide.Shapes.Paste(); // Paste it into the target slide
            }
        }

        //private void CopySlidesFromPresentation(PowerPoint.Application app, Presentation mainPresentation, string filePath)
        //{
        //    Presentation sourcePresentation = app.Presentations.Open(filePath);

        //    foreach (Slide slide in sourcePresentation.Slides)
        //    {
        //        slide.Copy();
        //        if (IsClipboardDataValidForPowerPoint(out string clipboardError))
        //        {
        //            mainPresentation.Slides.Paste(mainPresentation.Slides.Count + 1);
        //            Slide newSlide = mainPresentation.Slides[mainPresentation.Slides.Count];
        //            newSlide.CustomLayout = slide.CustomLayout;
        //            newSlide.FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse;
        //            CopyBackground(slide, newSlide);
        //        }
        //        else
        //        {
        //            MessageBox.Show($"Clipboard error: {clipboardError}");
        //        }
        //    }

        //    sourcePresentation.Close();
        //}

        private void CopyBackground(PowerPoint.Slide sourceSlide, PowerPoint.Slide targetSlide)
        {
            if (sourceSlide.Background.Fill.Type == Office.MsoFillType.msoFillPicture)
            {
                string picturePath = "";
                if (sourceSlide.Shapes.Count > 0)
                {
                    PowerPoint.Shape backgroundShape = sourceSlide.Shapes[1];
                    if (backgroundShape.Fill.Type == Office.MsoFillType.msoFillPicture)
                    {
                        backgroundShape.Copy();
                        targetSlide.FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse;
                        targetSlide.Shapes.Paste();
                    }
                }
            }
            else
            {
                MessageBox.Show("The source slide does not have a picture background.");
            }
        }

        private void btnAddShape_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            Slide slide = app.ActivePresentation.Slides.Add(app.ActivePresentation.Slides.Count + 1, PpSlideLayout.ppLayoutBlank);

            slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 50, 100, 200, 100);
            slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeQuadArrowCallout, 50, 250, 200, 100);
        }

        private void btnAddChart_Click(object sender, RibbonControlEventArgs e)
        {

            GEN_VSTO_Chart();
        }
        public static void GEN_VSTO_Chart()

        {
            PowerPoint.Application objPPT = Globals.ThisAddIn.Application;
            Presentation objPres = objPPT.ActivePresentation;

            Slide objSlide = objPPT.ActivePresentation.Slides.Add(objPPT.ActivePresentation.Slides.Count + 1, PpSlideLayout.ppLayoutBlank);

            //Select firs slide and set its layout

            objSlide.Select();

            objSlide.Layout = Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank;

            if (objSlide != null)
            {
                PowerPoint.Chart chart = objSlide.Shapes.AddChart2(-1,
          XlChartType.xl3DColumn, // Use PowerPoint's XlChartType
          20F, 30F, 400F, 300F, true).Chart;

                //Access the added chart

                Microsoft.Office.Interop.PowerPoint.Chart ppChart = objSlide.Shapes[1].Chart;

                //Access the chart data

                Microsoft.Office.Interop.PowerPoint.ChartData chartData = ppChart.ChartData;

                //Create instance to Excel workbook to work with chart data

                Microsoft.Office.Interop.Excel.Workbook dataWorkbook = (Microsoft.Office.Interop.Excel.Workbook)chartData.Workbook;

                //Accessing the data worksheet for chart

                Microsoft.Office.Interop.Excel.Worksheet dataSheet = dataWorkbook.Worksheets[1];

                //Setting the range of chart

                Microsoft.Office.Interop.Excel.Range tRange = dataSheet.Cells.get_Range("A1", "B5");

                //Applying the set range on chart data table

                Microsoft.Office.Interop.Excel.ListObject tbl1 = dataSheet.ListObjects["Table1"];

                tbl1.Resize(tRange);

                //Setting values for categories and respective series data

                ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("A2"))).FormulaR1C1 = "Bikes";

                ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("A3"))).FormulaR1C1 = "Accessories";

                ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("A4"))).FormulaR1C1 = "Repairs";

                ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("A5"))).FormulaR1C1 = "Clothing";

                ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("B2"))).FormulaR1C1 = "1000";

                ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("B3"))).FormulaR1C1 = "2500";

                ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("B4"))).FormulaR1C1 = "4000";

                ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("B5"))).FormulaR1C1 = "3000";

                //Setting chart title

                ppChart.ChartTitle.Font.Italic = true;

                ppChart.ChartTitle.Text = "2007 Sales";

                ppChart.ChartTitle.Font.Size = 18;

                ppChart.ChartTitle.Font.Color = Color.Black.ToArgb();

                ppChart.ChartTitle.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;

                ppChart.ChartTitle.Format.Line.ForeColor.RGB = Color.Black.ToArgb();

                //Accessing Chart value axis

                Microsoft.Office.Interop.PowerPoint.Axis valaxis = ppChart.Axes(Microsoft.Office.Interop.PowerPoint.XlAxisType.xlValue, Microsoft.Office.Interop.PowerPoint.XlAxisGroup.xlPrimary);

                //Setting values axis units

                valaxis.MajorUnit = 2000.0F;

                valaxis.MinorUnit = 1000.0F;

                valaxis.MinimumScale = 0.0F;

                valaxis.MaximumScale = 4000.0F;

                //Accessing Chart Depth axis

                Microsoft.Office.Interop.PowerPoint.Axis Depthaxis = ppChart.Axes(Microsoft.Office.Interop.PowerPoint.XlAxisType.xlSeriesAxis, Microsoft.Office.Interop.PowerPoint.XlAxisGroup.xlPrimary);

                Depthaxis.Delete();

                //Setting chart rotation

                ppChart.Rotation = 20; //Y-Value

                ppChart.Elevation = 15; //X-Value

                ppChart.RightAngleAxes = false;

                // Save the presentation as a PPTX

                objPres.SaveAs("VSTOSampleChart", Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);

                //Close Workbook and presentation

                dataWorkbook.Application.Quit();

                objPres.Application.Quit();
            }     

        }

    }
}

using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;
using System.Drawing.Imaging;
using System.Net;
using Control = System.Windows.Forms.Control;

namespace MultiMediaSolver
{
    public partial class Form1 : Form
    {
        int amt = 15;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            tabControl1.TabPages.Clear();

            for (int i = 1; i <= amt; i++)
            {
                tabControl1.TabPages.Add(i.ToString());

                Globals.slides.Add(i, new());
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Window will be frozen while loading. Expect ETA 1-3 minutes on most computers.");
            foreach (System.Windows.Forms.TextBox tb in Controls.OfType<System.Windows.Forms.TextBox>())
            {
                if (string.IsNullOrEmpty(tb.Text))
                {
                    MessageBox.Show(tb.Name.Replace("tb_", "") + " is empty.");
                    return;
                }
            }


            foreach (int slideNumber in Globals.slides.Keys)
            {

                switch (slideNumber)
                {
                    case 1:
                        Globals.slides[slideNumber] = Globals.getTitlePage(tb_title.Text, tb_name.Text, tb_date.Text);
                        break;
                    default:
                        if (Enumerable.Range(2, 4).Contains(slideNumber))
                        {
                            switch (slideNumber)
                            {
                                case 2:
                                    Globals.slides[slideNumber] = Globals.getInfoPage("Common Tasks for " + tb_name1.Text, API.getRandomTasks(), API.getImageUrl(tb_name1.Text));
                                    break;
                                case 3:
                                    Globals.slides[slideNumber] = Globals.getInfoPage("Common Work Activities for " + tb_name1.Text, API.getRandomActivities(), API.getImageUrl(tb_name1.Text));
                                    break;
                                case 4:
                                    Globals.slides[slideNumber] = Globals.getInfoPage("Common Work Contexts for " + tb_name1.Text, API.getRandomContexts(), API.getImageUrl(tb_name1.Text));
                                    break;
                            }
                        }
                        if (Enumerable.Range(5, 7).Contains(slideNumber))
                        {
                            switch (slideNumber)
                            {
                                case 5:
                                    Globals.slides[slideNumber] = Globals.getInfoPage("Common Tasks for " + tb_name2.Text, API.getRandomTasks(), API.getImageUrl(tb_name2.Text));
                                    break;
                                case 6:
                                    Globals.slides[slideNumber] = Globals.getInfoPage("Common Work Activities for " + tb_name2.Text, API.getRandomActivities(), API.getImageUrl(tb_name2.Text));
                                    break;
                                case 7:
                                    Globals.slides[slideNumber] = Globals.getInfoPage("Common Work Contexts for " + tb_name2.Text, API.getRandomContexts(), API.getImageUrl(tb_name2.Text));
                                    break;
                            }
                        }
                        if (Enumerable.Range(8, 10).Contains(slideNumber))
                        {
                            switch (slideNumber)
                            {
                                case 8:
                                    Globals.slides[slideNumber] = Globals.getInfoPage("Common Tasks for " + tb_name3.Text, API.getRandomTasks(), API.getImageUrl(tb_name3.Text));
                                    break;
                                case 9:
                                    Globals.slides[slideNumber] = Globals.getInfoPage("Common Work Activities for " + tb_name3.Text, API.getRandomActivities(), API.getImageUrl(tb_name3.Text));
                                    break;
                                case 10:
                                    Globals.slides[slideNumber] = Globals.getInfoPage("Common Work Contexts for " + tb_name3.Text, API.getRandomContexts(), API.getImageUrl(tb_name3.Text));
                                    break;
                            }
                        }
                        if (Enumerable.Range(11, 13).Contains(slideNumber))
                        {
                            switch (slideNumber)
                            {
                                case 11:
                                    Globals.slides[slideNumber] = Globals.getInfoPage("Common Tasks for " + tb_name4.Text, API.getRandomTasks(), API.getImageUrl(tb_name4.Text));
                                    break;
                                case 12:
                                    Globals.slides[slideNumber] = Globals.getInfoPage("Common Work Activities for " + tb_name4.Text, API.getRandomActivities(), API.getImageUrl(tb_name4.Text));
                                    break;
                                case 13:
                                    Globals.slides[slideNumber] = Globals.getInfoPage("Common Work Contexts for " + tb_name4.Text, API.getRandomContexts(), API.getImageUrl(tb_name4.Text));
                                    break;
                            }
                        }
                        if (Enumerable.Range(14, 15).Contains(slideNumber))
                        {
                            switch (slideNumber)
                            {
                                case 14:
                                    Globals.slides[slideNumber] = Globals.getComparisonPage(tb_name1.Text, tb_name2.Text, tb_name3.Text, tb_name4.Text,
                                        tb_annual1.Text, tb_annual2.Text, tb_annual3.Text, tb_annual4.Text,
                                        tb_rank1.Text, tb_rank2.Text, tb_rank3.Text, tb_rank4.Text);
                                    break;
                                case 15:
                                    List<string> images = new();
                                    foreach (TabPage t in Globals.slides.Values)
                                    {
                                        if (t.Tag != null)
                                        {
                                            images.Add(t.Tag.ToString());
                                        }
                                    }

                                    Globals.slides[slideNumber] = Globals.getCitedPage(tb_link1.Text, tb_link2.Text, tb_link3.Text, tb_link4.Text, images);
                                    break;
                            }
                        }
                        break;
                }
            }

            tabControl1.TabPages.Clear();

            for (int i = 1; i <= amt; i++)
            {
                TabPage page = Globals.slides[i];
                page.Text = i.ToString();
                tabControl1.TabPages.Add(page);

            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            tb_link1.Text = "http://www.onetonline.org/link/details/21-2011.00#Tasks";
            tb_link2.Text = "http://www.onetonline.org/link/details/39-5092.00#Tasks";
            tb_link3.Text = "http://www.onetonline.org/link/details/21-1013.00#Tasks";
            tb_link4.Text = "http://www.onetonline.org/link/details/21-1093.00#Tasks";

            tb_name.Text = "Jeffery";
            tb_title.Text = "Careers";
            tb_date.Text = "3/24/2022";

            tb_name1.Text = "Clergy";
            tb_name2.Text = "Manicurists and Pedicurists";
            tb_name3.Text = "Marriage and Family Therapists";
            tb_name4.Text = "Social and Human Service Assistants";

            tb_annual1.Text = "$51,940";
            tb_annual2.Text = "$27,870";
            tb_annual3.Text = "$51,340";
            tb_annual4.Text = "$35,960";

            tb_rank1.Text = "best";
            tb_rank2.Text = "worst";
            tb_rank3.Text = "middle";
            tb_rank4.Text = "worst";

        }

        private void tabPage1_Resize(object sender, EventArgs e)
        {
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            Globals.tabSize = tabControl1.Size;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Window will be frozen while loading. Expect ETA 1-3 minutes on most computers.");
            int doneamt = 0;
            using (Presentation pres = new Presentation())
            {
                ISlideCollection slides = pres.Slides;
                int i = 0;
                foreach (int tb in Globals.slides.Keys)
                {
                    try
                    {
                        pres.Slides.AddEmptySlide(pres.LayoutSlides[tb - 1]);
                        ISlide sld = pres.Slides[tb - 1];
                        foreach (System.Windows.Forms.Control c in Globals.slides[tb].Controls)
                        {

                            switch (c.Name)
                            {
                                case "maintitle":
                                    IAutoShape ashp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 30, 30, 663, 100);
                                    ashp.AddTextFrame(c.Text);
                                    ITextFrame txtFrame = ashp.TextFrame;
                                    IPortion port = txtFrame.Paragraphs[0].Portions[0];
                                    port.PortionFormat.FontHeight = 40;

                                    break;
                                case "bydate":
                                    IAutoShape ashp2 = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 30, 375, 663, 100);
                                    ashp2.FillFormat.FillType = FillType.NoFill;
                                    ashp2.AddTextFrame(c.Text);
                                    ITextFrame txtFrame2 = ashp2.TextFrame;
                                    IPortion port2 = txtFrame2.Paragraphs[0].Portions[0];
                                    port2.PortionFormat.FillFormat.FillType = FillType.Solid;
                                    port2.PortionFormat.FillFormat.SolidFillColor.Color = System.Drawing.Color.Black;
                                    port2.PortionFormat.FontHeight = 37;
                                    break;
                                case "infotitle":
                                    IAutoShape ashp3 = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 30, 30, 663, 100);
                                    ashp3.AddTextFrame(c.Text);
                                    ITextFrame txtFrame3 = ashp3.TextFrame;
                                    IPortion port3 = txtFrame3.Paragraphs[0].Portions[0];
                                    port3.PortionFormat.FontHeight = 40;
                                    break;
                                case "infobox":
                                    IAutoShape ashp4 = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 30, 210, 320, 300);
                                    ashp4.FillFormat.FillType = FillType.NoFill;
                                    ashp4.AddTextFrame(c.Text);
                                    ITextFrame txtFrame4 = ashp4.TextFrame;
                                    IPortion port4 = txtFrame4.Paragraphs[0].Portions[0];
                                    port4.PortionFormat.FillFormat.FillType = FillType.Solid;
                                    port4.PortionFormat.FillFormat.SolidFillColor.Color = System.Drawing.Color.Black;
                                    port4.PortionFormat.FontHeight = 16;
                                    break;
                                case "infopic":
                                    byte[] imageData;
                                    using (WebClient webClient = new WebClient())
                                    {
                                        imageData = webClient.DownloadData(new Uri(Globals.slides[tb].Tag.ToString()));
                                    }

                                    IPPImage image = pres.Images.AddImage(imageData);
                                    sld.Shapes.AddPictureFrame(ShapeType.Rectangle, 360, 210, 320, 300, image);

                                    break;
                            }
                        }
                        i++;
                    }
                    catch (Exception eee) { }
                }
                doneamt = i;
                pres.Save(@"P1.pptx", SaveFormat.Pptx);
            }

            for (int j = 0; j <= doneamt; j++)
            {
                Globals.slides.Remove(j);
            }
            int again = doneamt;
            using (Presentation pres = new Presentation())
            {
                ISlideCollection slides = pres.Slides;
                int i = 0;
                int z = 0;
                foreach (int tb in Globals.slides.Keys)
                {
                    try
                    {
                        pres.Slides.AddEmptySlide(pres.LayoutSlides[z]);
                        ISlide sld = pres.Slides[z];
                        foreach (System.Windows.Forms.Control c in Globals.slides[tb].Controls)
                        {

                            switch (c.Name)
                            {
                                case "maintitle":
                                    IAutoShape ashp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 30, 30, 663, 100);
                                    ashp.AddTextFrame(c.Text);
                                    ITextFrame txtFrame = ashp.TextFrame;
                                    IPortion port = txtFrame.Paragraphs[0].Portions[0];
                                    port.PortionFormat.FontHeight = 40;

                                    break;
                                case "bydate":
                                    IAutoShape ashp2 = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 30, 375, 663, 100);
                                    ashp2.FillFormat.FillType = FillType.NoFill;
                                    ashp2.AddTextFrame(c.Text);
                                    ITextFrame txtFrame2 = ashp2.TextFrame;
                                    IPortion port2 = txtFrame2.Paragraphs[0].Portions[0];
                                    port2.PortionFormat.FillFormat.FillType = FillType.Solid;
                                    port2.PortionFormat.FillFormat.SolidFillColor.Color = System.Drawing.Color.Black;
                                    port2.PortionFormat.FontHeight = 37;
                                    break;
                                case "infotitle":
                                    IAutoShape ashp3 = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 30, 30, 663, 100);
                                    ashp3.AddTextFrame(c.Text);
                                    ITextFrame txtFrame3 = ashp3.TextFrame;
                                    IPortion port3 = txtFrame3.Paragraphs[0].Portions[0];
                                    port3.PortionFormat.FontHeight = 40;
                                    break;
                                case "infobox":
                                    IAutoShape ashp4 = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 30, 210, 320, 300);
                                    ashp4.FillFormat.FillType = FillType.NoFill;
                                    ashp4.AddTextFrame(c.Text);
                                    ITextFrame txtFrame4 = ashp4.TextFrame;
                                    IPortion port4 = txtFrame4.Paragraphs[0].Portions[0];
                                    port4.PortionFormat.FillFormat.FillType = FillType.Solid;
                                    port4.PortionFormat.FillFormat.SolidFillColor.Color = System.Drawing.Color.Black;
                                    port4.PortionFormat.FontHeight = 16;
                                    break;
                                case "infopic":
                                    byte[] imageData;
                                    using (WebClient webClient = new WebClient())
                                    {
                                        imageData = webClient.DownloadData(new Uri(Globals.slides[tb].Tag.ToString()));
                                    }

                                    IPPImage image = pres.Images.AddImage(imageData);
                                    sld.Shapes.AddPictureFrame(ShapeType.Rectangle, 360, 210, 320, 300, image);

                                    break;
                                case "cited":
                                    IAutoShape ashp7 = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 30, 30, 663, 100);
                                    ashp7.AddTextFrame("Works Cited");
                                    ITextFrame txtFrame7 = ashp7.TextFrame;
                                    IPortion port7 = txtFrame7.Paragraphs[0].Portions[0];
                                    port7.PortionFormat.FontHeight = 40;

                                    IAutoShape ashp5 = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 30, 210, 640, 300);
                                    ashp5.AddTextFrame(c.Text);
                                    ITextFrame txtFrame5 = ashp5.TextFrame;
                                    IPortion port5 = txtFrame5.Paragraphs[0].Portions[0];
                                    port5.PortionFormat.FillFormat.FillType = FillType.Solid;
                                    port5.PortionFormat.FillFormat.SolidFillColor.Color = System.Drawing.Color.Black;
                                    port5.PortionFormat.FontHeight = 16;

                                    break;
                                case "grid":
                                    IAutoShape ashp6 = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 30, 30, 663, 100);
                                    ashp6.AddTextFrame("Career Comparison");
                                    ITextFrame txtFrame6 = ashp6.TextFrame;
                                    IPortion port6 = txtFrame6.Paragraphs[0].Portions[0];
                                    port6.PortionFormat.FontHeight = 40;

                                    DataGridView grid = (DataGridView)c;
                                    IChart chart = sld.Shapes.AddChart(Aspose.Slides.Charts.ChartType.ClusteredColumn, 30, 210, 640, 300);
                                    chart.ChartTitle.AddTextFrameForOverriding("Career Comparison");
                                    chart.ChartTitle.TextFrameForOverriding.TextFrameFormat.CenterText = NullableBool.True;
                                    chart.ChartTitle.Height = 40;
                                    chart.HasTitle = true;

                                    // Sets the first series to show values
                                    chart.ChartData.Series[0].Labels.DefaultDataLabelFormat.ShowValue = true;

                                    // Sets the index for the chart data sheet
                                    int defaultWorksheetIndex = 0;

                                    // Gets the chart data worksheet
                                    IChartDataWorkbook fact = chart.ChartData.ChartDataWorkbook;

                                    chart.ChartData.Series.Clear();
                                    chart.ChartData.Categories.Clear();
                                    int s = chart.ChartData.Series.Count;
                                    s = chart.ChartData.Categories.Count;


                                    // Adds new series
                                    chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 1, ""), chart.Type);
                                    chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 2, ""), chart.Type);
                                    chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 3, "Income"), chart.Type);

                                    // Adds new categories
                                    chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 1, 0, grid.Rows[0].Cells[0].Value));
                                    chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 2, 0, grid.Rows[1].Cells[0].Value));
                                    chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 3, 0, grid.Rows[2].Cells[0].Value));

                                    IChartSeries series = chart.ChartData.Series[0];



                                    series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 1, 1, int.Parse(grid.Rows[0].Cells[1].Value.ToString().Replace("$", "").Replace(",", ""))));
                                    series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 2, 1, int.Parse(grid.Rows[1].Cells[1].Value.ToString().Replace("$", "").Replace(",", ""))));
                                    series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 3, 1, int.Parse(grid.Rows[2].Cells[1].Value.ToString().Replace("$", "").Replace(",", ""))));

                                    // Sets the fill color for series
                                    series.Format.Fill.FillType = FillType.Solid;
                                    series.Format.Fill.SolidFillColor.Color = System.Drawing.Color.LightBlue;

                                    // Sets the first label to show Category name
                                    IDataLabel lbl = series.DataPoints[0].Label;
                                    lbl.DataLabelFormat.ShowCategoryName = true;

                                    lbl = series.DataPoints[1].Label;
                                    lbl.DataLabelFormat.ShowSeriesName = false;

                                    // Sets the series to show the value for the third label
                                    lbl = series.DataPoints[2].Label;
                                    lbl.DataLabelFormat.ShowValue = false;
                                    lbl.DataLabelFormat.ShowSeriesName = false;
                                    lbl.DataLabelFormat.Separator = "/";

                                    break;
                            }
                        }
                        i++;
                    }
                    catch (Exception asdasd) { }
                    z++;
                }
                again += i;
                pres.Save(@"P2.pptx", SaveFormat.Pptx);
            }
            //CombineMultipleFilesIntoSingleFile(@"P1.pptx", @"P2.pptx", @"Careers PowerPoint.pptx");
            MessageBox.Show("Completed / Still very early build. Exports 2 files, P1.pptx, P2.pptx." + Environment.NewLine + "You can just copy and paste the slides over." + Environment.NewLine + "Also contains spam on most pages which can easily be removed by delete key.");
        }

        // Binary File Copy
        private static void CombineMultipleFilesIntoSingleFile(string inputDirectoryPath, string inputFileNamePattern, string outputFilePath)
        {
            string[] inputFilePaths = Directory.GetFiles(inputDirectoryPath, inputFileNamePattern);
            Console.WriteLine("Number of files: {0}.", inputFilePaths.Length);
            foreach (var inputFilePath in inputFilePaths)
            {
                using (var outputStream = File.AppendText(outputFilePath))
                {
                    // Buffer size can be passed as the second argument.
                    outputStream.WriteLine(File.ReadAllText(inputFilePath));
                    Console.WriteLine("The file {0} has been processed.", inputFilePath);

                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            /*
            if (!System.IO.Directory.Exists(@"images"))
            {
                System.IO.Directory.CreateDirectory(@"images");

            }
            else
            {
                System.IO.Directory.Delete(@"images");
                System.IO.Directory.CreateDirectory(@"images");
            }*/

            int i = 0;
            foreach (TabPage p in tabControl1.TabPages)
            {
                SaveAsBitmap(p, System.IO.Directory.GetCurrentDirectory() + "\\images\\slide" + i.ToString() + ".bmp");
                i++;
            }

        }
        public void SaveAsBitmap(Control control, string fileName)
        {
            //get the instance of the graphics from the control
            Graphics g = control.CreateGraphics();

            //new bitmap object to save the image
            Bitmap bmp = new Bitmap(control.Width, control.Height);

            //Drawing control to the bitmap
            control.DrawToBitmap(bmp, new Rectangle(0, 0, control.Width, control.Height));

            bmp.Save(fileName);
            bmp.Dispose();

        }
    }
}
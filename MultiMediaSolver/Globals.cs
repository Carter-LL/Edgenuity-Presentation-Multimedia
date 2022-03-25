using AutoMapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultiMediaSolver
{
    public static class Globals
    {
        public static Size tabSize = new Size(740, 511);

        public static Dictionary<int, TabPage> slides = new();

        private static Color backColor = Color.DarkCyan;
        private static Color infobackColor = Color.Snow;
        private static Color infoforecolor = Color.DarkBlue;

        private static TabPage tab_TitlePage()
        {
            TabPage page = new();
            page.Size = tabSize;
            page.BackColor = backColor;
            return page;

        }

        private static TabPage tab_InfoPage()
        {
            TabPage page = new();
            page.Size = tabSize;
            page.BackColor = infobackColor;
            return page;
        }

        private static RichTextBox textbox_title()
        {
            RichTextBox tb = new();
            tb.Font = new Font("Arial", 25.0f, FontStyle.Bold);
            tb.ForeColor = Color.Black;
            tb.Location = new Point(30, tabSize.Height / 3);
            tb.BackColor = backColor;
            tb.BorderStyle = BorderStyle.None;
            tb.WordWrap = true;
            return tb;
        }

        private static TextBox textbox_byDate()
        {
            TextBox tb = new();
            tb.Font = new Font("Arial", 19.0f);
            tb.ForeColor = Color.WhiteSmoke;
            tb.Location = new Point(30, (tabSize.Height / 3) + 150);
            tb.BackColor = backColor;
            tb.BorderStyle = BorderStyle.None;
            return tb;
        }

        private static RichTextBox richtextbox_info()
        {
            RichTextBox rtb = new();
            rtb.Font = new Font("Arial", 10.0f);
            rtb.ForeColor = infoforecolor;
            rtb.Location = new Point(30, (tabSize.Height / 3) + 10);
            rtb.BackColor = infobackColor;
            rtb.BorderStyle = BorderStyle.None;
            rtb.WordWrap = true;
            return rtb;
        }

        private static PictureBox pictureBox_category()
        {
            PictureBox pb = new();
            pb.Location = new Point(30, (tabSize.Height / 3) + 50);
            pb.BorderStyle = BorderStyle.None;
            pb.Size = new Size(300, 300);
            pb.BackgroundImageLayout = ImageLayout.Stretch;
            pb.Anchor = (AnchorStyles.Right | AnchorStyles.Bottom);
            return pb;
        }


        private static Size MeasureText(string text, Font font)
        {
            return TextRenderer.MeasureText(text, font);
        }


        public static TabPage getTitlePage(string title, string name, string date)
        {
            RichTextBox tb_title = textbox_title();
            tb_title.Text = title;
            tb_title.Name = "maintitle";

            tb_title.Width = MeasureText(tb_title.Text, tb_title.Font).Width;
            tb_title.Height = MeasureText(tb_title.Text, tb_title.Font).Height;

            TextBox tb_byDate = textbox_byDate();
            tb_byDate.Text = name + " - " + date;
            tb_byDate.Name = "bydate";

            tb_byDate.Width = MeasureText(tb_byDate.Text, tb_byDate.Font).Width;
            tb_byDate.Height = MeasureText(tb_byDate.Text, tb_byDate.Font).Height;

            TabPage tb_page = tab_TitlePage();

            tb_page.Controls.Add(tb_title);
            tb_page.Controls.Add(tb_byDate);

            return tb_page;
        }
        public static TabPage getComparisonPage(string name1, string name2, string name3, string name4, 
            string annual1, string annual2, string annual3, string annual4,
            string rate1, string rate2, string rate3, string rate4)
        {
            RichTextBox tb_title = textbox_title();
            tb_title.Text = "Career Comparison";
            tb_title.Location = new Point(50, 30);
            tb_title.BackColor = infobackColor;

            tb_title.Width = tabSize.Width - 20;
            tb_title.Height = 80;

            DataGridView grid = new DataGridView();
            grid.Name = "grid";
            
            grid.Font = new Font("Arial", 12.0f);
            grid.Location = new Point(60, (tabSize.Height / 3) + 10);
            grid.Width = 330;
            grid.Height = 300;


            grid.ColumnCount = 3;
            grid.Columns[0].Name = "Carrer";
            grid.Columns[1].Name = "Annual Income";
            grid.Columns[2].Name = "Ranking";

            string[] row = new string[] { name1, annual1, rate1 };
            grid.Rows.Add(row);

             row = new string[] { name2, annual2, rate2 };
            grid.Rows.Add(row);

            row = new string[] { name3, annual3, rate3 };
            grid.Rows.Add(row);

            row = new string[] { name4, annual4, rate4 };
            grid.Rows.Add(row);

            TabPage tb_page = tab_InfoPage();
            tb_page.Controls.Add(tb_title);
            tb_page.Controls.Add(grid);

            return tb_page;

        }

        public static TabPage getCitedPage(string link1, string link2, string link3, string link4, List<string> images)
        {
            RichTextBox tb_title = textbox_title();
            tb_title.Text = "Works Cited";
            tb_title.Location = new Point(50, 30);
            tb_title.BackColor = infobackColor;

            tb_title.Width = tabSize.Width - 20;
            tb_title.Height = 150;

            RichTextBox rtb_info = richtextbox_info();
            rtb_info.Font = new Font("Arial", 12.0f);
            rtb_info.Location = new Point(60, (tabSize.Height / 3) + 10);
            rtb_info.WordWrap = false;
            rtb_info.Name = "cited";

            rtb_info.Text = link1 + Environment.NewLine + Environment.NewLine + link2 + Environment.NewLine + Environment.NewLine + link3 + Environment.NewLine + Environment.NewLine + link4 + Environment.NewLine + Environment.NewLine;

            foreach(string s in images)
            {
                //rtb_info.Text += s + Environment.NewLine + Environment.NewLine;
            }

            rtb_info.Width = tabSize.Width - 30;
            rtb_info.Height = MeasureText(rtb_info.Text, rtb_info.Font).Height + 200;

            TabPage tb_page = tab_InfoPage();

            tb_page.Controls.Add(tb_title);
            tb_page.Controls.Add(rtb_info);

            return tb_page;
        }

        public static TabPage getInfoPage(string title, string info, string imageUrl)
        {
            RichTextBox tb_title = textbox_title();
            tb_title.Text = title;
            tb_title.Location = new Point(50, 30);
            tb_title.BackColor = infobackColor;
            tb_title.Name = "infotitle";

            tb_title.Width = tabSize.Width - 20;
            tb_title.Height = 150;

            RichTextBox rtb_info = richtextbox_info();
            rtb_info.Text = info;
            rtb_info.Name = "infobox";

            rtb_info.Width = 370;
            rtb_info.Height = MeasureText(rtb_info.Text, rtb_info.Font).Height + 200;

            PictureBox pb_category = pictureBox_category();
            pb_category.Name = "infopic";

            byte[] image = API.GetImage(imageUrl);
            using (var ms = new MemoryStream(image))
            {
                pb_category.BackgroundImage = Image.FromStream(ms);
            }


            pb_category.Location = new Point(390, (tabSize.Height / 4));


            TabPage tb_page = tab_InfoPage();

            tb_page.Tag = imageUrl;

            tb_page.Controls.Add(tb_title);
            tb_page.Controls.Add(rtb_info);
            tb_page.Controls.Add(pb_category);

            return tb_page;
        }

    }
}

using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Words.NET;

namespace NameSheetGen
{
    public partial class Form1 : Form
    {
        private String[] holder;
        private String sectionNumber;
        private bool pdf = true, docx;
        public Form1()
        {
            InitializeComponent();

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            generate();
        }

        private void generate()
        {

            holder = richTextBox1.Text.Split('\n');
            sectionNumber = textBox1.Text;
            string path = SaveFile();
            if (docx)
            {
                string fileName = @Path.Combine(path, "NameSheet.docx");

                var lineFormat = new Formatting();
                lineFormat.Size = 20D;
                lineFormat.Bold = true;
                lineFormat.FontFamily = new Xceed.Words.NET.Font("Calibri");

                var secFormat = new Formatting();
                secFormat.Size = 16D;
                secFormat.Bold = false;
                secFormat.FontFamily = new Xceed.Words.NET.Font("Calibri");
                var doc = DocX.Create(fileName);

                for (int i = 0; i < holder.Length; i++)
                {
                    doc.InsertParagraph(sectionNumber, false, secFormat).Alignment = Alignment.left;
                    doc.InsertParagraph(holder[i], false, lineFormat).Alignment = Alignment.left;
                    for(int j = 0; j < 48; j++)
                    {
                        doc.InsertParagraph();
                    }                    
                }
                doc.Save();

            }
            if (pdf)
            {
                string fileName = @Path.Combine(path, "NameSheet.pdf");
                PdfDocument doc = new PdfDocument();
                for(int i = 0; i < holder.Length; i++)
                {
                    PdfPage page = doc.AddPage();
                    XGraphics gfx = XGraphics.FromPdfPage(page);
                    XFont lineFont = new XFont("Calibri", 24, XFontStyle.Bold);
                    XFont secFont = new XFont("Calibri", 16, XFontStyle.Regular);
                    gfx.DrawString(sectionNumber, secFont, XBrushes.Black, new XRect(50, 50, page.Width, page.Height), XStringFormats.TopLeft);
                    gfx.DrawString(holder[i], lineFont, XBrushes.Black, new XRect(50, 70, page.Width, page.Height), XStringFormats.TopLeft);

                }
                doc.Save(fileName);
            }
        }

        private string SaveFile()
        {
            // Show the FolderBrowserDialog.
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                string folderName = folderBrowserDialog1.SelectedPath;
                MessageBox.Show("Saved To:\n" + folderName);
                return folderName;
            }
            return " ";
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                docx = true;
            }
            else
            {
                docx = false;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show("This software is developed specifically for generating name sheets in Northwest Missouri State University.\nDeveloper Name: Kancharla, Sai Krishna Teja.\nReport any bugs to baymaxsoftwares@gmail.com");
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/tejavictory/name_sheet_gen");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                pdf = true;
            }
            else
            {
                pdf = false;
            }
        }
    }
}

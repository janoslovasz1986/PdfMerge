using iTextSharp.text;
using iTextSharp.text.pdf;
using Kevsoft.PDFtk;
using PdfMerge.Model;
using System.Collections;
using System.Collections.Generic;
using System.Reflection.PortableExecutable;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static PdfMerge.Form1;
using System.Linq;
using com.itextpdf.text.pdf;
using System.util;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace PdfMerge
{
    public partial class Form1 : Form
    {

        int index = 0;
        int indexOfItemUnderMouseToDrop = -1;

        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            listBox1.AllowDrop = true;
            listBox1.DragDrop += new DragEventHandler(listBox1_DragDrop);
            listBox1.DragEnter += new DragEventHandler(listBox1_DragEnter);
        }



        void listBox1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;

        }


        void listBox1_DragDrop(object sender, DragEventArgs e)
        {
            string item = listBox1.Items[index].ToString();

            listBox1.Items.RemoveAt(index);

            if (indexOfItemUnderMouseToDrop == -1)
            {
                listBox1.Items.Insert(listBox1.Items.Count, item);
            }

            if (indexOfItemUnderMouseToDrop < listBox1.Items.Count && indexOfItemUnderMouseToDrop != -1)
            {
                listBox1.Items.Insert(indexOfItemUnderMouseToDrop, item);
            }
            index = 0;
            indexOfItemUnderMouseToDrop = -1;

            GlobalVariables.listOfPdfFiles = listBox1.Items.Cast<String>().ToList();
        }

        private void listBox1_DragOver(object sender, DragEventArgs e)
        {

            indexOfItemUnderMouseToDrop =
                listBox1.IndexFromPoint(listBox1.PointToClient(new Point(e.X, e.Y)));
            // Determine whether string data exists in the drop data. If not, then
            // the drop effect reflects that the drop cannot occur.
            if (!e.Data.GetDataPresent(typeof(System.String)))
            {
                e.Effect = DragDropEffects.None;
                //DropLocationLabel.Text = "None - no string data.";
                return;
            }

            else if ((e.AllowedEffect & DragDropEffects.Move) == DragDropEffects.Move)
            {
                // By default, the drop action should be move, if allowed.
                e.Effect = DragDropEffects.Move;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }

            // Updates the label text.
            if (indexOfItemUnderMouseToDrop != ListBox.NoMatches)
            {
                //DropLocationLabel.Text = "Drops before item #" + (indexOfItemUnderMouseToDrop + 1);
            }
            else
            {
                //DropLocationLabel.Text = "Drops at the end.";
            }


        }



        void button1_Click(object sender, EventArgs e)
        {

            //int pos = 0;

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Pdf files (*.pdf)|*.pdf";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                if (!(pdfElementListFullPathContainsString(GlobalVariables.listOfPdfElements, dialog.FileName)))
                {
                    PdfElement pdfElement = new PdfElement(dialog.FileName, removeFullPath(dialog.FileName));
                    GlobalVariables.listOfPdfElements.Add(pdfElement);
                }
            }
            listBox1.Items.Clear();
            listBox1.Items.AddRange(getListElementshortPath(GlobalVariables.listOfPdfElements).ToArray());

            GlobalVariables.listOfPdfFiles = listBox1.Items.Cast<String>().ToList();

        }

        void openFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }




        void button2_Click(object sender, EventArgs e)
        {
            //string outputPdfPath = @"C:\\Temp\\pdf\\new.pdf";
            string outputPdfPath = "";
            outputPdfPath = showSaveFileDialog();


            if (GlobalVariables.listOfPdfFiles.Count != 0 && outputPdfPath != "")
            {
                PdfReader reader = null;
                Document sourceDocument = null;
                PdfCopy pdfCopyProvider = null;
                PdfImportedPage importedPage;


                sourceDocument = new Document();
                pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));

                //Open the output file
                sourceDocument.Open();

                try
                {
                    List<String> fullpathOfPdfFiles = replaceShortPathToFullPath(GlobalVariables.listOfPdfFiles);


                    for (int i = 0; i < fullpathOfPdfFiles.Count; i++)
                    {
                        int pages = get_pageCcount(fullpathOfPdfFiles[i]);

                        reader = new PdfReader(fullpathOfPdfFiles[i]);
                        //Add pages of current file
                        for (int j = 1; j <= pages; j++)
                        {
                            importedPage = pdfCopyProvider.GetImportedPage(reader, j);
                            pdfCopyProvider.AddPage(importedPage);
                        }

                        reader.Close();
                    }

                    //At the end save the output file
                    sourceDocument.Close();
                    System.Windows.Forms.MessageBox.Show("A pdf összefüzése sikeres volt!");
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("A pdf összefüzése sikertelen volt!" + ex.ToString());
                }

                clearGlobalData();
                listBox1.Items.Clear();
            }
            else if (GlobalVariables.listOfPdfFiles.Count == 0)
                System.Windows.Forms.MessageBox.Show("Nem adott meg összefûzendõ pdf-eket");

        }

        static int get_pageCcount(string file)
        {
            using (StreamReader sr = new StreamReader(File.OpenRead(file)))
            {
                return new Regex(@"/Type\s*/Page[^s]").Matches(sr.ReadToEnd()).Count;

            }
        }


        void listBox1_MouseDown(object sender, MouseEventArgs e)
        {
            ListBox listbox1 = sender as ListBox;
            index = listbox1.IndexFromPoint(e.X, e.Y);
            if (index >= 0 & e.Button == MouseButtons.Left)
                listbox1.DoDragDrop(listbox1.Items[index].ToString(), DragDropEffects.Move);
        }

        void listBox1_MouseUp(object sender, MouseEventArgs e)
        {
            ListBox listbox1 = sender as ListBox;
            index = listbox1.IndexFromPoint(e.X, e.Y);
            if (index >= 0 & e.Button == MouseButtons.Left)
                listbox1.DoDragDrop(listbox1.Items[index].ToString(), DragDropEffects.Move);
        }


        public static Boolean pdfElementListFullPathContainsString(List<PdfElement> list, string s)
        {
            for (int i = 0; i < list.Count; i++)
            {
                if (list.ElementAt(i).Fullpath == s)
                    return true;
            }

            return false;
        }


        public static String removeFullPath(string input)
        {
            int pos;
            pos = input.LastIndexOf('\\');
            string found = input.Substring(pos + 1);

            return found;
        }


        public static List<String> getListElementshortPath(List<PdfElement> list)
        {
            List<String> output = new List<String>();

            for (int i = 0; i < list.Count; i++)
            {
                var item = list.ElementAt(i);
                var firstItemShortpath = item.Shortpath;

                output.Add(firstItemShortpath);
            }

            return output;

        }



        public static List<String> getListElementFullPath(List<PdfElement> list)
        {
            List<String> output = new List<String>();

            for (int i = 0; i < list.Count; i++)
            {
                var item = list.ElementAt(i);
                var firstItemShortpath = item.Fullpath;

                output.Add(firstItemShortpath);
            }

            return output;

        }

        //public static List<PdfElement> replaceShortPathToFullPath(List<String> list)
        public static List<String> replaceShortPathToFullPath(List<String> list)
        {

            List<String> _toPrintOrder = new List<string>();
            for (int j = 0; j < list.Count; j++)
            {
                for (int i = 0; i < GlobalVariables.listOfPdfElements.Count; i++)
                {
                    if (list[j] == GlobalVariables.listOfPdfElements.ElementAt(i).Shortpath)
                    {

                        String s = GlobalVariables.listOfPdfElements.ElementAt(i).Fullpath;
                        _toPrintOrder.Add(s);
                    }
                }
            }

            return _toPrintOrder;
        }


        public static String showSaveFileDialog()
        {
            String outputPdfPath = "";

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Pdf files (*.pdf)|*.pdf";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //MessageBox.Show("You selected file: " + saveFileDialog.FileName);
                outputPdfPath = saveFileDialog.FileName;
            }



            return outputPdfPath;


        }
    


        public static void clearGlobalData() {

            GlobalVariables.listItemFiles.Clear();
            GlobalVariables.listOfPdfFiles.Clear();
            GlobalVariables.listOfPdfElements.Clear();

        
        }


    }

    public static class GlobalVariables {

        public static List<String> listOfPdfFiles = new List<String>();
        public static List<String> listItemFiles = new List<String>();
        public static List<PdfElement> listOfPdfElements = new List<PdfElement>();

    }
}

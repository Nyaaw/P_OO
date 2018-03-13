using System;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Net;

namespace SmartThesaurus
{
    class Program
    {
        [STAThread]
        public static void Main()
        {
            ///////////////LIRE UN FICHIER TXT////////////////////////

            //try
            //{   // Open the text file using a stream reader.
            //    using (StreamReader sr = new StreamReader(@"..\..\TestsLecture\hello.txt"))
            //    {
            //        // Read the stream to a string, and write the string to the console.
            //        String line = sr.ReadToEnd();
            //        Console.WriteLine(line);
            //    }
            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine("The file could not be read:");
            //    Console.WriteLine(e.Message);
            //}
            //Console.ReadLine();

            ///////////////////////////////////////////////////////////


            ///////////////LIRE UN FICHIER DOCX////////////////////////

            //Word.Application wordApp = new Word.Application();

            //// Input box is used to get the path of the file which has to be 
            ////uploaded into textbox.

            //string filePath = @"C:\Users\dutoitrugu\Desktop\Test.docx";

            //object file = filePath;

            //object nullobj = System.Reflection.Missing.Value;

            //// here on Document.Open there should be 9 arg.

            //Word.Document doc = wordApp.Documents.Open(ref file,
            //ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj,
            //ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj,
            //ref nullobj);

            //// Here the word content is copeied into a string which helps to
            ////store it into textbox.

            //Word.Document doc1 = wordApp.ActiveDocument;

            //string m_Content = doc1.Content.Text;

            //Console.Write(m_Content);
            //Console.ReadLine();

            //// the content need to be stored into the DB.

            //doc.Close(ref nullobj, ref nullobj, ref nullobj);

            //////////////////////////////////////////////////////////

            /////////////////Lire un document Excel///////////////////

            //const string fileName = @"C:\Users\dutoitrugu\Desktop\test.xlsx";

            //// Ouvrir un ducument Excel
            //var application = new Microsoft.Office.Interop.Excel.Application();
            //var workbook = application.Workbooks.Open(fileName);
            //var worksheet = workbook.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;

            //int i = 1;
            //int j = 1;

            //string text = "";

            ////Vas lire tout le texte cellule par cellule jusqu'à ce que la cellule sois null et vas l'ajouter dans une variable séparer d'espace
            //do
            //{
            //    doc = app.Documents.Open(ref path, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            //    string text = doc.Content.Text;
            //    File.WriteAllText(txtPath, text, new UTF8Encoding());
            //    Console.WriteLine("Converted!");

            //    doc.ActiveWindow.Selection.WholeStory();
            //    doc.ActiveWindow.Selection.Copy();

            //    string ClipboardText;

            //    if (Clipboard.ContainsText(TextDataFormat.Text))
            //    {
            //        ClipboardText = Clipboard.GetText(TextDataFormat.Text);
            //    }
            //    else
            //    {
            //        ClipboardText = "Data not found";
            //    }

            //    Console.WriteLine(ClipboardText);

            //    doc.Close(WdSaveOptions.wdDoNotSaveChanges, WdOriginalFormat.wdOriginalDocumentFormat);
            //}
            //catch
            //{
            //    Console.WriteLine("An error occured. Please check the file path to your word document, and whether the word document is valid.");
            //}
            //finally
            //{
            //    object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
            //    app.Quit(ref saveChanges, ref missing, ref missing);
            //}

            //Console.ReadLine();

            //Save.
            //workbook.Save();
            ///////////////////////////////////////////////////////////

            //////////////////LIRE UN FICHIER PW/////////////////////

            //string presentation_text = ReadPPT();

            //Console.WriteLine(presentation_text);

            //Console.ReadLine();

            //////////////////////////////////////////////////////////

            //////////////////LIRE PAGE WEB///////////////////////////


            //Vas enregistrer la page web demandée pour pouvoir la lire avec Word par la suite
            using (WebClient client = new WebClient()) 
            {
                client.DownloadFile("https://www.etml.ch/", @"C:\Users\dutoitrugu\Desktop\localfile.html");              
            }


           
            //Lecture du fichier html par word

            Word.Application wordApp = new Word.Application();

            // Input box is used to get the path of the file which has to be 
            //uploaded into textbox.

            string filePath = @"C:\Users\dutoitrugu\Desktop\localfile.html";

            object file = filePath;

            object nullobj = System.Reflection.Missing.Value;

            // here on Document.Open there should be 9 arg.

            Word.Document doc = wordApp.Documents.Open(ref file,
            ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj,
            ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj,
            ref nullobj);

            // Here the word content is copeied into a string which helps to
            //store it into textbox.

            Word.Document doc1 = wordApp.ActiveDocument;


 
            string m_Content = doc1.Content.Text;

            //Ecris le texte dans un fichier txt

            try
            {

                //Pass the filepath and filename to the StreamWriter Constructor
                StreamWriter sw = new StreamWriter(@"C:\Users\dutoitrugu\Desktop\Test2.txt");

              

                //Write a second line of text
                sw.WriteLine(m_Content);

                //Close the file
                sw.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
            finally
            {
                Console.WriteLine("Executing finally block.");
            }

            Console.ReadLine();

            // the content need to be stored into the DB.

            doc.Close(ref nullobj, ref nullobj, ref nullobj);

            //////////////////////////////////////////////////////////
        }

        //private static string ReadPPT()
        //{
        //    PowerPoint.Application PowerPoint_App = new PowerPoint.Application();
        //    Microsoft.Office.Interop.PowerPoint.Presentations multi_presentations = PowerPoint_App.Presentations;
        //    PowerPoint.Presentation presentation = multi_presentations.Open(@"C:\Users\dutoitrugu\Desktop\Test.pptx");
        //    string presentation_text = "";
        //    for (int i = 0; i < presentation.Slides.Count; i++)
        //    {
        //        foreach (var item in presentation.Slides[i + 1].Shapes)
        //        {
        //            var shape = (PowerPoint.Shape)item;
        //            if (shape.HasTextFrame == MsoTriState.msoTrue)
        //            {
        //                if (shape.TextFrame.HasText == MsoTriState.msoTrue)
        //                {
        //                    var textRange = shape.TextFrame.TextRange;
        //                    var text = textRange.Text;
        //                    presentation_text += text + " ";
        //                }
        //            }
        //        }
        //    }

        //    presentation.Close();
        //    //PowerPoint_App.;
        //    PowerPoint_App.Quit();

        //    Process[] pros = Process.GetProcesses();
        //    for (int i = 0; i < pros.Count(); i++)
        //    {
        //        if (pros[i].ProcessName.ToLower().Contains("powerpnt"))
        //        {
        //            pros[i].Kill();
        //        }
        //    }

        //    //Marshal.ReleaseComObject(presentation);
        //    //Marshal.ReleaseComObject(PowerPoint_App);

        //    return presentation_text;
        //}
    }
}

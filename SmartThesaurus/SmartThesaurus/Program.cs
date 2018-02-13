using System;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Windows;


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

            object path = @"F:\ETML\P_OO\P_OO\SmartThesaurus\SmartThesaurus\TestsLecture\LOL.docx";

            string txtPath = "TxtOfWord.txt";

            Word.Application app = new Word.Application();
            Document doc;
            object missing = Type.Missing;
            object readOnly = true;
            try


            //////////////////////////////////////////////////////////

            /////////////////Lire un document Excel///////////////////

            const string fileName = @"C:\Users\dutoitrugu\Desktop\test.xlsx";
           
            // Ouvrir un ducument Excel
            var application = new Microsoft.Office.Interop.Excel.Application();
            var workbook = application.Workbooks.Open(fileName);
            var worksheet = workbook.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;

            int i = 1;
            int j = 1;

            string text = "";

            //Vas lire tout le texte cellule par cellule jusqu'à ce que la cellule sois null et vas l'ajouter dans une variable séparer d'espace
            do
            {
                doc = app.Documents.Open(ref path, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                string text = doc.Content.Text;
                File.WriteAllText(txtPath, text, new UTF8Encoding());
                Console.WriteLine("Converted!");

                doc.ActiveWindow.Selection.WholeStory();
                doc.ActiveWindow.Selection.Copy();

                string ClipboardText;

                if (Clipboard.ContainsText(TextDataFormat.Text))
                {
                    ClipboardText = Clipboard.GetText(TextDataFormat.Text);
                }
                else
                {
                    ClipboardText = "Data not found";
                }

                Console.WriteLine(ClipboardText);

                doc.Close(WdSaveOptions.wdDoNotSaveChanges, WdOriginalFormat.wdOriginalDocumentFormat);
            }
            catch
            {
                Console.WriteLine("An error occured. Please check the file path to your word document, and whether the word document is valid.");
            }
            finally
            {
                object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                app.Quit(ref saveChanges, ref missing, ref missing);
            }

            Console.ReadLine();

            //Save.
            //workbook.Save();
            ///////////////////////////////////////////////////////////

            //////////////////LIRE UN FICHIER XLX/////////////////////

            //////////////////////////////////////////////////////////
        }


    }
}

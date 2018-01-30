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
            Word.Document doc;
            object missing = Type.Missing;
            object readOnly = true;
            try
            {
               doc = app.Documents.Open(ref path, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                //string text = doc.Content.Text;
                //File.WriteAllText(txtPath, text, new UTF8Encoding());
                //Console.WriteLine("Converted!");

                //doc.ActiveWindow.Selection.WholeStory();
                //doc.ActiveWindow.Selection.Copy();
                //IDataObject idata = Clipboard.GetDataObject();
                //Console.WriteLine(Clipboard.GetText());

                doc.Range().Copy();
                IDataObject n = Clipboard.GetDataObject();
                Console.WriteLine(n.GetData(DataFormats.Text).ToString());

                doc.Close(WdSaveOptions.wdDoNotSaveChanges, WdOriginalFormat.wdOriginalDocumentFormat);
            }
            catch
            {
                Console.WriteLine("An error occured. Please check the file path to your word document, and whether the word document is valid.");
            }
            finally
            {
                object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                app.Quit(ref saveChanges, ref missing, ref missing);
            }

            Console.ReadLine();

            ///////////////////////////////////////////////////////////

            //////////////////LIRE UN FICHIER XLX/////////////////////
            
            //////////////////////////////////////////////////////////
        }


    }
}

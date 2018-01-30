using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO;
using Word = Microsoft.Office.Interop.Word;



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
                ////string text = doc.Content.Text;
                string text = doc.Content.FormattedText.Text;
                File.WriteAllText(txtPath, text, new UTF8Encoding());
                Console.WriteLine("Converted!");
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

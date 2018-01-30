using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Word = Microsoft.Office.Interop.Word;


namespace SmartThesaurus
{
    class Program
    {
        public static void Main()
        {
            try
            {   // Open the text file using a stream reader.
                using (StreamReader sr = new StreamReader(@"..\..\TestsLecture\hello.txt"))
                {
                    // Read the stream to a string, and write the string to the console.
                    String line = sr.ReadToEnd();
                    Console.WriteLine(line);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(e.Message);
            }
            Console.ReadLine();

            ///////////////LIRE UN FICHIER TXT////////////////////////

            //using (StreamReader reader = new StreamReader("../../Test2.docx"))
            //{
            //    string[] test;

            //    string content = reader.ReadToEnd();
            //    reader.Close();

            //    test = content.Split(' ');

            //    foreach(string element in test)
            //    {
            //        Console.WriteLine(element);
            //    }




            //    Console.ReadLine();
            //}

            ///////////////////////////////////////////////////////////


            ///////////////LIRE UN FICHIER DOCX////////////////////////

            object path = @"K:\INF\Eleves\Temp\LOL.docx";

            string txtPath = "TxtOfWord.txt";

            Word.Application app = new Word.Application();
            Word.Document doc;
            object missing = Type.Missing;
            object readOnly = true;
            try
            {
                doc = app.Documents.Open(ref path, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                string text = doc.Content.Text;
                //File.WriteAllText(txtPath, text);
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

            ///////////////////////////////////////////////////////////

            //////////////////LIRE UN FICHIER XLX/////////////////////













            //////////////////////////////////////////////////////
        }


    }
}

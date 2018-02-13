using System;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Windows;
using RasterEdge.Imaging.Basic;
using RasterEdge.XDoc.Converter;
using RasterEdge.XDoc.PDF;
using System.Collections.Generic;
using System.Drawing;
using Microsoft.Office.Interop.Excel;

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

            //object path = @"F:\ETML\P_OO\P_OO\SmartThesaurus\SmartThesaurus\TestsLecture\LOL.docx";

            //string txtPath = "TxtOfWord.txt";

            //Word.Application app = new Word.Application();
            //Word.Document doc;
            //object missing = Type.Missing;
            //object readOnly = true;
            //try
            //{
            //   doc = app.Documents.Open(ref path, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            //    //string text = doc.Content.Text;
            //    //File.WriteAllText(txtPath, text, new UTF8Encoding());
            //    //Console.WriteLine("Converted!");

            //    //doc.ActiveWindow.Selection.WholeStory();
            //    //doc.ActiveWindow.Selection.Copy();
            //    //IDataObject idata = Clipboard.GetDataObject();
            //    //Console.WriteLine(Clipboard.GetText());

            //    doc.Range().Copy();
            //    IDataObject n = Clipboard.GetDataObject();
            //    Console.WriteLine(n.GetData(DataFormats.Text).ToString());

            //    doc.Close(WdSaveOptions.wdDoNotSaveChanges, WdOriginalFormat.wdOriginalDocumentFormat);
            //}
            //catch
            //{
            //    Console.WriteLine("An error occured. Please check the file path to your word document, and whether the word document is valid.");
            //}
            //finally
            //{
            //    object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
            //    app.Quit(ref saveChanges, ref missing, ref missing);
            //}

            //Console.ReadLine();

            ///////////////////////////////////////////////////////////

            //////////////////LIRE UN FICHIER PDF/////////////////////


            //convertPdfToText();



            //////////////////////////////////////////////////////////

            /////////////////Lire un document Excel///////////////////

            const string fileName = @"C:\Users\dutoitrugu\Desktop\test.xlsx";
           
            // Open Excel and get first worksheet.
            var application = new Microsoft.Office.Interop.Excel.Application();
            var workbook = application.Workbooks.Open(fileName);
            var worksheet = workbook.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;

            int i = 1;
            int j = 1;

            string text = "";

            do
            {
                text += worksheet.Cells[i, j].value + " ";

                j++;

                if(j == 20)
                {
                    j = 1;
                    i++;
                }

            } while (worksheet.Cells[i, j].value != null);

            

            Console.WriteLine(text);

            Console.ReadLine();

            // Save.
            //workbook.Save();

            /////////////////////////////////////////////////////////
        }


        //#region pdf to text (file to file)
        //internal static void convertPdfToText()
        //{
        //    String inputFilePath = @"C:\Users\dutoitrugu\Desktop\2C-E-P_Web2-ISI001-CdC.pdf";
        //    String outputFilePath = @"C:\Users\dutoitrugu\Desktop\Test.txt";
        //    StreamWriter writer = new StreamWriter(outputFilePath);
        //    PDFDocument doc = new PDFDocument(inputFilePath);
        //    PDFTextMgr textMgr = PDFTextHandler.ExportPDFTextManager(doc);
        //    int pageCount = doc.GetPageCount();
        //    for (int i = 0; i < pageCount; i++)
        //    {
        //        PDFPage page = (PDFPage)doc.GetPage(i);
        //        List<PDFTextLine> pageTextLines = textMgr.ExtractTextLine(page);


        //        writeTextLines(pageTextLines, writer);
        //    }
        //    writer.Close();
        //}
        //#endregion


        //private static void writeTextLines(List<PDFTextLine> pageTextLines, StreamWriter writer)
        //{
        //    String lineText = "";
        //    float positionY = 0f;
        //    float height = 0f;
        //    float positionX = 0f;


        //    if (pageTextLines != null)
        //    {
        //        for (int i = 0; i < pageTextLines.Count; i++)
        //        {
        //            RectangleF rectangle = pageTextLines[i].GetBoundary();
        //            if (i != 0 && !isEqual(positionY + height, rectangle.Y + rectangle.Height))
        //            {
        //                writer.WriteLine(lineText);
        //                lineText = "";
        //            }
        //            if (positionX > rectangle.X)
        //            {
        //                lineText = getTextLineContent(pageTextLines[i]) + " " + lineText;
        //            }
        //            else
        //            {
        //                lineText += getTextLineContent(pageTextLines[i]);
        //                lineText += "    ";
        //            }
        //            positionY = rectangle.Y;
        //            height = rectangle.Height;
        //            positionX = rectangle.X;
        //            if (i == pageTextLines.Count - 1)
        //            {
        //                writer.WriteLine(lineText);
        //            }
        //        }
        //    }


        //    writer.WriteLine(" ");
        //    writer.WriteLine(" ");
        //    writer.Flush();
        //}

        //private static String getTextLineContent(PDFTextLine pdfTextLine)
        //{
        //    List<PDFTextWord> words = pdfTextLine.GetTextWord();
        //    String wordText = "";
        //    float positionX = 0;
        //    float width = 0;
        //    for (int i = 0; i < words.Count; i++)
        //    {
        //        RectangleF rectange = words[i].GetBoundary();
        //        if (i != 0 && !isEqual(positionX + width, rectange.X))
        //            wordText += " ";
        //        wordText += words[i].GetContent();
        //        positionX = rectange.X;
        //        width = rectange.Width;
        //    }

        //    return wordText;
        //}

        //private static bool isEqual(float first, float second)
        //{
        //    if (first - second < 2F && first - second > -2F)
        //        return true;
        //    return false;
        //}

    }
}

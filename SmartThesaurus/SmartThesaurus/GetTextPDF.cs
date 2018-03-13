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
		
		#region pdf to text (file to file)
        internal static void convertPdfToText()
        {
            String inputFilePath = @"F:\..\..\TestsLecture\hello.pdf";
            String outputFilePath = @"F:\..\..\TestsLecture\Test.txt";
            StreamWriter writer = new StreamWriter(outputFilePath);
            PDFDocument doc = new PDFDocument(inputFilePath);
            PDFTextMgr textMgr = PDFTextHandler.ExportPDFTextManager(doc);
            int pageCount = doc.GetPageCount();
            for (int i = 0; i < pageCount; i++)
            {
                PDFPage page = (PDFPage)doc.GetPage(i);
                List<PDFTextLine> pageTextLines = textMgr.ExtractTextLine(page);


                writeTextLines(pageTextLines, writer);
            }
            writer.Close();
        }
        #endregion


        private static void writeTextLines(List<PDFTextLine> pageTextLines, StreamWriter writer)
        {
            String lineText = "";
            float positionY = 0f;
            float height = 0f;
            float positionX = 0f;


            if (pageTextLines != null)
            {
                for (int i = 0; i < pageTextLines.Count; i++)
                {
                    RectangleF rectangle = pageTextLines[i].GetBoundary();
                    if (i != 0 && !isEqual(positionY + height, rectangle.Y + rectangle.Height))
                    {
                        writer.WriteLine(lineText);
                        lineText = "";
                    }
                    if (positionX > rectangle.X)
                    {
                        lineText = getTextLineContent(pageTextLines[i]) + " " + lineText;
                    }
                    else
                    {
                        lineText += getTextLineContent(pageTextLines[i]);
                        lineText += "    ";
                    }
                    positionY = rectangle.Y;
                    height = rectangle.Height;
                    positionX = rectangle.X;
                    if (i == pageTextLines.Count - 1)
                    {
                        writer.WriteLine(lineText);
                    }
                }
            }


            writer.WriteLine(" ");
            writer.WriteLine(" ");
            writer.Flush();
        }

        private static String getTextLineContent(PDFTextLine pdfTextLine)
        {
            List<PDFTextWord> words = pdfTextLine.GetTextWord();
            String wordText = "";
            float positionX = 0;
            float width = 0;
            for (int i = 0; i < words.Count; i++)
            {
                RectangleF rectange = words[i].GetBoundary();
                if (i != 0 && !isEqual(positionX + width, rectange.X))
                    wordText += " ";
                wordText += words[i].GetContent();
                positionX = rectange.X;
                width = rectange.Width;
            }

            return wordText;
        }

        private static bool isEqual(float first, float second)
        {
            if (first - second < 2F && first - second > -2F)
                return true;
            return false;
        }
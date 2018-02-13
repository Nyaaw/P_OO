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
using System.Threading;

namespace SmartThesaurus
{
    class Program
    {

        [STAThread]
        public static void Main()
        {
            /////////////////Lire un document Excel///////////////////

            const string fileName = @"F:\ETML\P_OO\P_OO\SmartThesaurus\SmartThesaurus\TestsLecture\test.xlsx";

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

                worksheet.Cells.FindNext("");
                //if (worksheet.Cells[i, j].text != null)
                //    text += worksheet.Cells[i, j].text + " ";

                if (i % 100 == 0)
                    Console.WriteLine(i);

                j++;

                if (j == 20)
                {
                    j = 1;
                    i++;
                }

            } while (i != 50);

            Console.WriteLine(text);

            //Save.
            workbook.Save();
            workbook.Close();

            Console.ReadLine();

            /* UTILISER CE CODE ???

            var usedRange = worksheet.UsedRange;
            int startRow = usedRange.Row;
            int endRow = startRow + usedRange.Rows.Count - 1;
            int startColumn = usedRange.Column;
            int endColumn = startColumn + usedRange.Columns.Count - 1;
            for (int row = startRow; row <= endRow; row++)
            {
                Excel.Range lastCell = worksheet.Cells[row, endColumn];
                if (lastCell.Value2 == null)
                    lastCell = lastCell.End[Excel.XlDirection.xlToLeft];
                var lastColumn = lastCell.Column;
                Console.WriteLine($"{row}: {lastColumn}");
            }
            */
        }
    }
}

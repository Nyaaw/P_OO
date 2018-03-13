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


			/////////////////Lire un document Excel///////////////////

            const string fileName = @"F:\..\..\TestsLecture\test.xlsx";
           
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

            //Save.
            //workbook.Save();

            /////////////////////////////////////////////////////////
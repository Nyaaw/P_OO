using System;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Word;
			
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
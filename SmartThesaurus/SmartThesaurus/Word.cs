using System;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Word;
			
object path = @"F:\ETML\P_OO\P_OO\SmartThesaurus\SmartThesaurus\TestsLecture\LOL.docx";

string txtPath = "TxtOfWord.txt";

Word.Application app = new Word.Application();
Document doc;
object missing = Type.Missing;
object readOnly = true;
try { 
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
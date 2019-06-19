using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;
namespace WordAPI
{
	class Program
	{
		static void Main(string[] args)
		{
			WordAPI word = new WordAPI(@"C:\Users\Komil\Desktop\test2.docx");
			//word.GetBoldText();
			word.GetRegularText("*");
			WordAPIFile file = new WordAPIFile(@"C:\Users\Komil\Desktop\sam.docx", word.Subjects);
			file.WriteToTable();
			//file.WriteToEmptyDoc();
			Console.WriteLine("Done");
			Console.ReadKey();
		}
	}
}

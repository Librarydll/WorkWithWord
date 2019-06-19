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
			WordAPI word = new WordAPI(@"C:\Users\Komil\Desktop\tt.docx");
			word.GetRegularText("*");
			foreach (var s in word.Subjects)
			{
				Console.WriteLine(s.Question+"\n"+s.Answer+"\n");
			}
			Console.ReadKey();
		}
	}
}

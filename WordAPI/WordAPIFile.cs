using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;
namespace WordAPI
{
	public class WordAPIFile
	{
		IEnumerable<Subject> Subjects { get; }
		DocX _document;
		public WordAPIFile(string path,IEnumerable<Subject> _sub)
		{
			Subjects = _sub;
			_document = DocX.Create(path);
		}
		~WordAPIFile()
		{
			_document.Dispose();
		}

		public void WriteToTable()
		{
			var t = _document.AddTable(Subjects.Count(), 3);
			int row = 0;
			int count = 1;
			foreach (var item in Subjects)
			{
				t.Rows[row].Cells[0].Paragraphs[0].Append(count.ToString());
				t.Rows[row].Cells[1].Paragraphs[0].Append(item.Question);
				t.Rows[row].Cells[2].Paragraphs[0].Append(item.Answer);
				count++;
				row++;
			}
			_document.InsertTable(t);
			_document.Save();
		}

		public void WriteToEmptyDoc()
		{
			foreach (var item in Subjects)
			{
				_document.InsertParagraph(item.Question).FontSize(18).Bold();
				_document.InsertParagraph(item.Answer).FontSize(14);
			}
			_document.Save();
		}
	}
}

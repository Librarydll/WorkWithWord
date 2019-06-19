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

		public void WriteDoc()
		{
			var t = _document.AddTable(Subjects.Count(), 3);
			int row = 0;
			int cell = 0;
			int count = 1;
			foreach (var item in Subjects)
			{
				t.Rows[row].Cells[cell].Paragraphs[0].Append(count.ToString());
				cell++;
				t.Rows[row].Cells[cell].Paragraphs[0].Append(item.Question);
				cell++;
				t.Rows[row].Cells[cell].Paragraphs[0].Append(item.Answer);
				count++;
				row++;
				cell = 0;
			}
			_document.InsertTable(t);
			_document.Save();
		}

	}
}

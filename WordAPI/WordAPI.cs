using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;
namespace WordAPI
{
	public class WordAPI 
	{
		private readonly string _path = "";
		private DocX _document;
		public IList<Subject> Subjects { get; }
		public WordAPI(string path)
		{
			_path = path;
			if (!File.Exists(_path))
				throw new FileNotFoundException();
			_document = DocX.Load(_path);
			Subjects = new List<Subject>();
		}
		public void GetBoldText()
		{
			string question = "";

			ClearSubject();
			int i = 1;
			foreach (var t in _document.Paragraphs)
			{
				if (t.IsBold(t.Text))
				{
					if (i % 2 == 0)
					{
						Subjects.Add(new Subject() { Question = question, Answer = t.Text });
					}
					else
					{
						question = t.Text;
					}
					i++;
				}
			}
		}
		public void GetRegularText(string reg="*")
		{
			string question = "";
			ClearSubject();
			foreach (var t in _document.Paragraphs)
			{
				if (t.IsBold(t.Text)&&t.Text.IsQuestion())
				{			
					question = t.Text;
				}

				if (t.Text.IsMatch(reg))
				{
					Subjects.Add(new Subject()
					{
						Question = question.RemoveReg(reg),
						Answer = t.Text.RemoveReg(reg)
					});
				}
			}
		}

		public void ClearSubject()
		{
			if(Subjects.Count>0)
			Subjects.Clear();
		}
	}
}

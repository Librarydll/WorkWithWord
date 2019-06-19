using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;
namespace WordAPI
{
	public static class WordLinq
	{
		public static bool IsBold(this Paragraph paragraph,string text)
		{
			 return !string.IsNullOrWhiteSpace(text)&& paragraph.MagicText[0].formatting?.Bold == true;
		}
	}
}

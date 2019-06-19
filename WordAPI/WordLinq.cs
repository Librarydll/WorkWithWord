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

		public static bool IsMatch(this string text,string reg)
		{
			return !string.IsNullOrWhiteSpace(text) && text.StartsWith(reg)==true;
		}

		public static string RemoveReg(this string text,string reg)
		{
			return text.Replace(reg, "").Trim();
		}
	}
}

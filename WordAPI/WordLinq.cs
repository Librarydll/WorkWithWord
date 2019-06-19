using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Xceed.Words.NET;
namespace WordAPI
{
	public static class WordLinq
	{
		static Regex regex ;
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

		public static bool IsQuestion(this string text)
		{	string pattern = @"^\G\d{1,}\)|\.";
			regex = new Regex(pattern);
			var firstStage = regex.IsMatch(text);
			if (!firstStage)
				return false;

			return text.EndsWith("?") || text.EndsWith(":") || text.EndsWith(".") || text.EndsWith("...");
		}

		public static bool IsAnswer(this string text)
		{
			string pattern = @"^[a-zа-я]\.|\)";
			regex = new Regex(pattern, RegexOptions.IgnoreCase);

			return regex.IsMatch(text);


		}
	}
}

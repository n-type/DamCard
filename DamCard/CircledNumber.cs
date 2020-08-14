using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace DamCard
{
    public static class CircledNumber
    {
        public static bool Search(string text)
        {
            return Regex.IsMatch(text, "①|②|③|④|⑤|⑥|⑦|⑧|⑨");
        }

        public static int SearchMaxNumber(string text)
        {
            for(int i = 9; i > 0 ; i--)
            {
                if (text.IndexOf(i.ToCircledNumber()) > 0) return i;
            }
            return 0;
        }

        public static string ExtractTextByCircledNumber(string text,int number)
        {
            if (number > 9) throw new ArgumentOutOfRangeException("number", "Number must be in the range 1-9");
            var target = number.ToCircledNumber();
            var next = (number + 1).ToCircledNumber();

            int targetPosition = text.IndexOf(target);
            // 指定した丸数字が無い場合、空白で返す
            if (targetPosition < 0) return string.Empty;
            var startPosition = targetPosition + 1;

            int endPosition = text.IndexOf(next);
            return endPosition > 0 ? text.Substring(startPosition, endPosition - startPosition).Trim() : text.Substring(startPosition).Trim();
        }

        public static string ToCircledNumber(this int number)
        {
            return number switch
            {
                1 => "①",
                2 => "②",
                3 => "③",
                4 => "④",
                5 => "⑤",
                6 => "⑥",
                7 => "⑦",
                8 => "⑧",
                9 => "⑨",
                _ => "",
            };
        }

        public static int ToNumber(this string circledNumber)
        {
            return circledNumber switch
            {
                "①" => 1,
                "②" => 2,
                "③" => 3,
                "④" => 4,
                "⑤" => 5,
                "⑥" => 6,
                "⑦" => 7,
                "⑧" => 8,
                "⑨" => 9,
                _ => 0,
            };
        }

    }
}

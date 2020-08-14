using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DamCard
{
    /// <summary>
    /// 変換タイプ
    /// </summary>
    public enum ConvertTypes
    {
        Numeric = 1,
        Symbol = 2,
        Alphabet = 3,
        Katakana = 4,
        All = 99
    }

    /// <summary>
    /// 文字変換クラス
    /// </summary>
    public static class StringConverter
    {
        /// <summary>
        /// 半角数字
        /// </summary>
        public static readonly string[] HalfWidthNumeric = new string[]
        {
            "0","1","2","3","4","5","6","7","8","9"
        };
        /// <summary>
        /// 全角数字
        /// </summary>
        public static readonly string[] FullWidthNumeric = new string[]
        {
            "０","１","２","３","４","５","６","７","８","９"
        };
        /// <summary>
        /// 半角記号
        /// (チルダを除く)
        /// </summary>
        public static readonly string[] HalfWidthSymbol = new string[]
        {
            "!","\"","#","$","%","&","'","(",")","*","+",",","-",".","/",":",";","<","=",">","?","@","[","\\","]","^","_","`","{","|","}"
        };
        /// <summary>
        /// 全角記号
        /// (チルダを除く)
        /// </summary>
        public static readonly string[] FullWidthSymbol = new string[]
        {
            "！","”","＃","＄","％","＆","’","（","）","＊","＋","，","－","．","／","：","；","＜","＝","＞","？","＠","［","￥","］","＾","＿","｀","｛","｜","｝"
        };
        /// <summary>
        /// 半角英字
        /// </summary>
        public static readonly string[] HalfWidthAlphabet = new string[]
        {
            "A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z",
            "a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z"
        };
        /// <summary>
        /// 全角英字
        /// </summary>
        public static readonly string[] FullWidthAlphabet = new string[]
        {
            "Ａ","Ｂ","Ｃ","Ｄ","Ｅ","Ｆ","Ｇ","Ｈ","Ｉ","Ｊ","Ｋ","Ｌ","Ｍ","Ｎ","Ｏ","Ｐ","Ｑ","Ｒ","Ｓ","Ｔ","Ｕ","Ｖ","Ｗ","Ｘ","Ｙ","Ｚ",
            "ａ","ｂ","ｃ","ｄ","ｅ","ｆ","ｇ","ｈ","ｉ","ｊ","ｋ","ｌ","ｍ","ｎ","ｏ","ｐ","ｑ","ｒ","ｓ","ｔ","ｕ","ｖ","ｗ","ｘ","ｙ","ｚ"
        };
        /// <summary>
        /// 半角カタカナ
        /// </summary>
        public static readonly string[] HalfWidthKatakana = new string[]
        {
            "ｱ","ｲ","ｳ","ｴ","ｵ",
            "ｶ","ｷ","ｸ","ｹ","ｺ",
            "ｻ","ｼ","ｽ","ｾ","ｿ",
            "ﾀ","ﾁ","ﾂ","ﾃ","ﾄ",
            "ﾅ","ﾆ","ﾇ","ﾈ","ﾉ",
            "ﾊ","ﾋ","ﾌ","ﾍ","ﾎ",
            "ﾏ","ﾐ","ﾑ","ﾒ","ﾓ",
            "ﾔ","ﾕ","ﾖ",
            "ﾗ","ﾘ","ﾙ","ﾚ","ﾛ",
            "ﾜ","ｦ","ﾝ",
            "ｧ","ｨ","ｩ","ｪ","ｫ",
            "ｯ","ｬ","ｭ","ｮ","ｰ",
            "ｶﾞ","ｷﾞ","ｸﾞ","ｹﾞ","ｺﾞ",
            "ｻﾞ","ｼﾞ","ｽﾞ","ｾﾞ","ｿﾞ",
            "ﾀﾞ","ﾁﾞ","ﾂﾞ","ﾃﾞ","ﾄﾞ",
            "ﾊﾞ","ﾋﾞ","ﾌﾞ","ﾍﾞ","ﾎﾞ",
            "ﾊﾟ","ﾋﾟ","ﾌﾟ","ﾍﾟ","ﾎﾟ",
            "ｳﾞ"
        };
        /// <summary>
        /// 全角カタカナ
        /// </summary>
        public static readonly string[] FullWidthKatakana = new string[]
        {
            "ア","イ","ウ","エ","オ",
            "カ","キ","ク","ケ","コ",
            "サ","シ","ス","セ","ソ",
            "タ","チ","ツ","テ","ト",
            "ナ","ニ","ヌ","ネ","ノ",
            "ハ","ヒ","フ","ヘ","ホ",
            "マ","ミ","ム","メ","モ",
            "ヤ","ユ","ヨ",
            "ラ","リ","ル","レ","ロ",
            "ワ","ヲ","ン",
            "ァ","ィ","ゥ","ェ","ォ",
            "ッ","ャ","ュ","ョ","ー",
            "ガ","ギ","グ","ゲ","ゴ",
            "ザ","ジ","ズ","ゼ","ゾ",
            "ダ","ヂ","ヅ","デ","ド",
            "バ","ビ","ブ","ベ","ボ",
            "パ","ピ","プ","ペ","ポ",
            "ヴ"
        };

        /// <summary>
        /// 全角半角変換
        /// </summary>
        /// <param name="convertString">対象文字列</param>
        /// <param name="convertTypes">変換対象</param>
        /// <returns>変換後文字列</returns>
        public static string ToHalf(this string convertString, ConvertTypes convertTypes)
        {
            if (convertTypes == ConvertTypes.All || convertTypes == ConvertTypes.Numeric)
            {
                for (int i = 0; i < FullWidthNumeric.Length; i++)
                {
                    convertString = convertString.Replace(FullWidthNumeric[i], HalfWidthNumeric[i]);
                }
            }

            if (convertTypes == ConvertTypes.All || convertTypes == ConvertTypes.Symbol)
            {
                for (int i = 0; i < FullWidthSymbol.Length; i++)
                {
                    convertString = convertString.Replace(FullWidthSymbol[i], HalfWidthSymbol[i]);
                }
            }

            if (convertTypes == ConvertTypes.All || convertTypes == ConvertTypes.Alphabet)
            {
                for (int i = 0; i < FullWidthAlphabet.Length; i++)
                {
                    convertString = convertString.Replace(FullWidthAlphabet[i], HalfWidthAlphabet[i]);
                }
            }

            if (convertTypes == ConvertTypes.All || convertTypes == ConvertTypes.Katakana)
            {
                for (int i = 0; i < FullWidthKatakana.Length; i++)
                {
                    convertString = convertString.Replace(FullWidthKatakana[i], HalfWidthKatakana[i]);
                }
            }

            return convertString;
        }

        /// <summary>
        /// 半角全角変換
        /// </summary>
        /// <param name="convertString">対象文字列</param>
        /// <param name="convertTypes">変換対象</param>
        /// <returns>変換後文字列</returns>
        public static string ToFull(this string convertString, ConvertTypes convertTypes)
        {
            if (convertTypes == ConvertTypes.All || convertTypes == ConvertTypes.Numeric)
            {
                for (int i = 0; i < HalfWidthNumeric.Length; i++)
                {
                    convertString = convertString.Replace(HalfWidthNumeric[i], FullWidthNumeric[i]);
                }
            }

            if (convertTypes == ConvertTypes.All || convertTypes == ConvertTypes.Symbol)
            {
                for (int i = 0; i < HalfWidthSymbol.Length; i++)
                {
                    convertString = convertString.Replace(HalfWidthSymbol[i], FullWidthSymbol[i]);
                }
            }

            if (convertTypes == ConvertTypes.All || convertTypes == ConvertTypes.Alphabet)
            {
                for (int i = 0; i < HalfWidthAlphabet.Length; i++)
                {
                    convertString = convertString.Replace(HalfWidthAlphabet[i], FullWidthAlphabet[i]);
                }
            }

            if (convertTypes == ConvertTypes.All || convertTypes == ConvertTypes.Katakana)
            {
                for (int i = (HalfWidthKatakana.Length - 1); i >= 0; i--)
                {
                    convertString = convertString.Replace(HalfWidthKatakana[i], FullWidthKatakana[i]);
                }
            }

            return convertString;
        }


    }
}

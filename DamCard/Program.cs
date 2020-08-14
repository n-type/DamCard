using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace DamCard
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Title = CommonConstants.Title;
            WriteTitle(CommonConstants.Title);

            var exeDirectory = AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\');
            var readFilePath = Path.Combine(exeDirectory, CommonConstants.DamCardFileName);
            var dataDirectory = Path.Combine(exeDirectory, CommonConstants.DataDirectory);

            // ドラッグされたファイルを使う
            var files = Environment.GetCommandLineArgs();
            if (files.Length > 1)
            {
                readFilePath = files[1].ToString();
            }

            if (!File.Exists(readFilePath))
            {
                Console.WriteLine("ファイルが存在しません。");
                Finish();
                Environment.Exit(1);
            }

            var extension = Path.GetExtension(readFilePath);
            if (!extension.Equals(".xlsx"))
            {
                Console.WriteLine("Excelファイルではありません。");
                Finish();
                Environment.Exit(2);
            }

            var editFileName = CommonConstants.OutputDamcardFileNamePrefix
                + Path.GetFileNameWithoutExtension(readFilePath);
            var editFilePath = Path.Combine(Path.GetDirectoryName(readFilePath), editFileName + ".xlsx");

            if (File.Exists(editFilePath))
            {
                Console.WriteLine("ダムカード編集Excelが既に存在します。");
                Console.WriteLine("場所：" + editFilePath + Environment.NewLine);
                ConsoleKey consoleKey;
                do
                {
                    Console.Write("削除して続行しますか？(y/n) > ");
                    consoleKey = Console.ReadKey().Key;
                    Console.WriteLine(Environment.NewLine);
                } while (consoleKey != ConsoleKey.Y && consoleKey != ConsoleKey.N);

                if (consoleKey == ConsoleKey.Y)
                {
                    File.Delete(editFilePath);
                }
                else
                {
                    Console.WriteLine("終了します。");
                    Finish();
                    Environment.Exit(0);
                }
            }

            var excel = new DamCardExcel(readFilePath);
            Console.WriteLine("ダムカードExcel読み込み");
            excel.Read(readFilePath);
            var damcards = excel.LoadDamCardList();

            Console.WriteLine("ダム番号検索ファイル読み込み" + Environment.NewLine);
            var damNoData = new DamNoData(dataDirectory);

            Console.WriteLine("ダムカード編集Excel作成");
            var editExcel = new EditDamCardExcel(editFilePath);
            Console.CursorVisible = false;
            Console.Write("処理中...   ");
            int counter = 0;
            foreach (var damCard in damcards)
            {
                counter++;
                Console.Write("{0} / {1}", counter, damcards.Count);
                int no = damNoData.Search(damCard.DamName, damCard.RiverSystem, damCard.DamPrefecture);

                var editDamCard = new EditDamCard();
                editDamCard.No = damCard.No;
                editDamCard.RiverSystem = damCard.RiverSystem;
                editDamCard.River = damCard.River;
                editDamCard.DamName = damCard.DamName;
                editDamCard.Ver = damCard.Ver;
                editDamCard.DamPrefecture = damCard.DamPrefecture;
                editDamCard.DamNo = no;
                int max = CircledNumber.SearchMaxNumber(damCard.Place);
                int i = 0;
                do
                {
                    i++;
                    editDamCard.PlaceNo = i;
                    // 配布先が複数のカードは行分割する
                    if (max == 0)
                    {
                        editDamCard.Place = damCard.Place;
                        editDamCard.DateAndTime = damCard.DateAndTime;
                        editDamCard.Address = damCard.Address;
                        editDamCard.Url = damCard.Url;
                    }
                    else
                    {
                        editDamCard.Place = CircledNumber.ExtractTextByCircledNumber(damCard.Place, i);
                        editDamCard.DateAndTime = CircledNumber.SearchMaxNumber(damCard.DateAndTime) == max
                            ? CircledNumber.ExtractTextByCircledNumber(damCard.DateAndTime, i)
                            : damCard.DateAndTime;
                        editDamCard.Address = CircledNumber.SearchMaxNumber(damCard.Address) == max
                            ? CircledNumber.ExtractTextByCircledNumber(damCard.Address, i)
                            : damCard.Address;
                        editDamCard.Url = CircledNumber.SearchMaxNumber(damCard.Url) == max
                            ? CircledNumber.ExtractTextByCircledNumber(damCard.Url, i)
                            : damCard.Url;
                    }

                    // いい感じに編集
                    // ダム名の改行は外す
                    editDamCard.DamName = editDamCard.DamName.Replace("\n", "")
                                                             .ToHalf(ConvertTypes.Alphabet)
                                                             .ToHalf(ConvertTypes.Symbol);

                    // 配布場所の連続スペースはスペース1個にする
                    string place = editDamCard.Place.Replace("㈱", "(株)")
                                                    .ToFull(ConvertTypes.Katakana)
                                                    .ToHalf(ConvertTypes.Numeric)
                                                    .ToHalf(ConvertTypes.Symbol);
                    editDamCard.Place = Regex.Replace(Regex.Replace(Regex.Replace(place, "[　]+", "　"), "[ |　]+\n", "\n"), "^[ |　]+", "", RegexOptions.Multiline);

                    // 配布日時の行頭以外の連続スペースをスペース1個にする
                    // スマートな方法が思いつかないので、力業
                    string[] dateAndTime = editDamCard.DateAndTime.Replace("~", "～")
                                                                  .Replace("＊", "※")
                                                                  .Replace("②の配付", "配布")
                                                                  .Replace("※②で配布", "配布")
                                                                  .ToFull(ConvertTypes.Katakana)
                                                                  .ToHalf(ConvertTypes.Numeric)
                                                                  .ToHalf(ConvertTypes.Symbol)
                                                                  .Split("\n");
                    string editDateAndTime = string.Empty;
                    foreach (var text in dateAndTime)
                    {
                        if (Regex.IsMatch(text, "^(?![ |　]).+$"))
                        {
                            editDateAndTime += Regex.Replace(Regex.Replace(text, "[ |　]+", " "), "[ ]+$", "");
                        }
                        else
                        {
                            editDateAndTime += Regex.Replace(text, "[ |　]+$", "");
                        }
                        editDateAndTime += "\n";
                    }
                    editDamCard.DateAndTime = editDateAndTime.TrimEnd();

                    editDamCard.Address = editDamCard.Address.Replace("＊", "※")
                                                             .ToFull(ConvertTypes.Katakana)
                                                             .ToHalf(ConvertTypes.Alphabet)
                                                             .ToHalf(ConvertTypes.Numeric)
                                                             .ToHalf(ConvertTypes.Symbol);

                    editExcel.AddDamCardData(editDamCard);
                } while (i < max);

                Console.SetCursorPosition(12, Console.CursorTop);
            }
            Console.Write(Environment.NewLine);
            Console.CursorVisible = true;
            Console.WriteLine("セルの書式設定");
            editExcel.FormatCells();
            Console.WriteLine("ダムカード編集Excel保存中");
            editExcel.Save();
            Console.WriteLine("ダムカード編集Excel作成完了");
            Finish();
        }

        private static void WriteTitle(string title)
        {
            Console.WriteLine(title);
            Console.WriteLine(new string('-', title.Length * 2));
            Console.WriteLine();
        }

        private static void Finish()
        {
            Console.WriteLine(Environment.NewLine + "このウィンドウを閉じるには、任意のキーを押してください...");
            _ = Console.ReadKey();
        }
    }
}

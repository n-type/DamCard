using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace DamCard
{
    class DamCardExcel
    {
        private static readonly int StartRow = 3;
        private static readonly int ColumnNo = 0;
        private static readonly int ColumnRiverSystem = 1;
        private static readonly int ColumnRiver = 2;
        private static readonly int ColumnDamName = 3;
        private static readonly int ColumnVer = 4;
        private static readonly int ColumnPlace = 5;
        private static readonly int ColumnDateAndTime = 6;
        private static readonly int ColumnDamPrefecture = 7;
        private static readonly int ColumnAddress = 8;
        private static readonly int ColumnUrl = 9;

        private IWorkbook readBook;
        private ISheet readSheet;
        private int readSheetLastRow;

        public DamCardExcel()
        {

        }

        public DamCardExcel(string filePath)
        {
            Read(filePath);
        }


        public void Read(string filePath)
        {
            using (var fs = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                readBook = WorkbookFactory.Create(fs);
            }

            readSheet = readBook.GetSheetAt(0);
            readSheetLastRow = readSheet.LastRowNum;
        }

        public List<DamCard> LoadDamCardList()
        {
            var damCards = new List<DamCard>();

            for (int i = StartRow; i < readSheetLastRow; i++)
            {
                IRow row = readSheet.GetRow(i);

                // No未入力のセルになったら終了
                if (row.GetCell(ColumnNo).CellType == CellType.Blank) break;

                var damCard = new DamCard();

                damCard.No = row.GetCell(ColumnNo).NumericCellValue;
                damCard.RiverSystem = NPOIUtility.GetCellValueAsString(row.GetCell(ColumnRiverSystem));
                damCard.River = NPOIUtility.GetCellValueAsString(row.GetCell(ColumnRiver));
                damCard.DamName = NPOIUtility.GetCellValueAsString(row.GetCell(ColumnDamName));
                damCard.Ver = NPOIUtility.GetCellValueAsDouble(row.GetCell(ColumnVer));
                damCard.Place = NPOIUtility.GetCellValueAsString(row.GetCell(ColumnPlace));
                damCard.DateAndTime = NPOIUtility.GetCellValueAsString(row.GetCell(ColumnDateAndTime));
                damCard.DamPrefecture = NPOIUtility.GetCellValueAsString(row.GetCell(ColumnDamPrefecture));
                damCard.Address = NPOIUtility.GetCellValueAsString(row.GetCell(ColumnAddress));
                damCard.Url = NPOIUtility.GetCellValueAsString(row.GetCell(ColumnUrl));

                damCards.Add(damCard);
            }

            return damCards;
        }

    }
}

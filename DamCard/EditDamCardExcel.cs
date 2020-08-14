using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace DamCard
{
    class EditDamCardExcel
    {
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
        private static readonly int ColumnDamNo = 10;
        private static readonly int ColumnPlaceNo = 11;

        private string filePath;

        private IWorkbook writeBook;
        private ISheet editSheet;
        private int currentRow;

        public EditDamCardExcel()
        {
        }

        public EditDamCardExcel(string filePath)
        {
            Create(filePath);
            CreateEditSheet();
            Save();
        }

        public void Create(string filePath)
        {
            this.filePath = filePath;
            writeBook = new XSSFWorkbook();
            using var fs = File.Create(this.filePath);
            writeBook.Write(fs);
        }

        public void Save()
        {
            using var fs = File.Create(filePath);
            writeBook.Write(fs);
        }

        public void CreateEditSheet()
        {
            editSheet = writeBook.CreateSheet("編集");
            WriteHeader();
        }

        private void WriteHeader()
        {
            currentRow = 0;
            var row = editSheet.CreateRow(currentRow);
            row.CreateCell(ColumnNo).SetCellValue("番号");
            row.CreateCell(ColumnRiverSystem).SetCellValue("水系名");
            row.CreateCell(ColumnRiver).SetCellValue("河川名");
            row.CreateCell(ColumnDamName).SetCellValue("ダム名");
            row.CreateCell(ColumnVer).SetCellValue("ver");
            row.CreateCell(ColumnPlace).SetCellValue("配布場所");
            row.CreateCell(ColumnDateAndTime).SetCellValue("配布日時");
            row.CreateCell(ColumnDamPrefecture).SetCellValue("ダム所在県名");
            row.CreateCell(ColumnAddress).SetCellValue("配布場所の住所");
            row.CreateCell(ColumnUrl).SetCellValue("ホームページURL");
            row.CreateCell(ColumnDamNo).SetCellValue("ダム番号");
            row.CreateCell(ColumnPlaceNo).SetCellValue("配布場所番号");
        }

        public void AddDamCardData(EditDamCard card)
        {
            currentRow++;
            var row = editSheet.CreateRow(currentRow);
            row.CreateCell(ColumnNo).SetCellValue(card.No);
            row.CreateCell(ColumnRiverSystem).SetCellValue(card.RiverSystem);
            row.CreateCell(ColumnRiver).SetCellValue(card.River);
            row.CreateCell(ColumnDamName).SetCellValue(card.DamName);
            row.CreateCell(ColumnVer).SetCellValue(card.Ver);
            row.CreateCell(ColumnPlace).SetCellValue(card.Place);
            row.CreateCell(ColumnDateAndTime).SetCellValue(card.DateAndTime);
            row.CreateCell(ColumnDamPrefecture).SetCellValue(card.DamPrefecture);
            row.CreateCell(ColumnAddress).SetCellValue(card.Address);
            row.CreateCell(ColumnUrl).SetCellValue(card.Url);
            row.CreateCell(ColumnDamNo).SetCellValue(card.DamNo);
            row.CreateCell(ColumnPlaceNo).SetCellValue(card.PlaceNo);
        }

        public void FormatCells()
        {
            editSheet.AutoSizeColumn(ColumnNo);
            editSheet.AutoSizeColumn(ColumnRiverSystem);
            editSheet.AutoSizeColumn(ColumnRiver);
            editSheet.AutoSizeColumn(ColumnDamName);
            editSheet.AutoSizeColumn(ColumnVer);
            editSheet.AutoSizeColumn(ColumnPlace);
            editSheet.AutoSizeColumn(ColumnDateAndTime);
            editSheet.AutoSizeColumn(ColumnDamPrefecture);
            editSheet.AutoSizeColumn(ColumnAddress);
            editSheet.AutoSizeColumn(ColumnUrl);
            editSheet.AutoSizeColumn(ColumnDamNo);
            editSheet.AutoSizeColumn(ColumnPlaceNo);

            var style = writeBook.CreateCellStyle();
            style.WrapText = true;
            style.Alignment = HorizontalAlignment.General;
            style.VerticalAlignment = VerticalAlignment.Top;

            for (var i = 0; i <= currentRow; i++)
            {
                foreach (var cell in editSheet.GetRow(i).Cells)
                {
                    cell.CellStyle = style;
                }
            }
        }

    }
}

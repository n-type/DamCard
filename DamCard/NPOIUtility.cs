using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace DamCard
{
    class NPOIUtility
    {
        public static string GetCellValueAsString(ICell cell)
        {
            return cell.CellType switch
            {
                CellType.String => cell.StringCellValue,
                CellType.Numeric => DateUtil.IsCellDateFormatted(cell) ? cell.DateCellValue.ToString() : cell.NumericCellValue.ToString(),
                CellType.Boolean => cell.BooleanCellValue.ToString(),
                CellType.Blank => cell.ToString(),
                CellType.Formula => cell.CachedFormulaResultType switch
                {
                    CellType.String => cell.StringCellValue,
                    CellType.Numeric => DateUtil.IsCellDateFormatted(cell) ? cell.DateCellValue.ToString() : cell.NumericCellValue.ToString(),
                    CellType.Boolean => cell.BooleanCellValue.ToString(),
                    CellType.Blank => cell.ToString(),
                    CellType.Error => cell.ErrorCellValue.ToString(),
                    _ => null,
                },
                CellType.Error => cell.ErrorCellValue.ToString(),
                CellType.Unknown => throw new NotImplementedException(),
                _ => null,
            };
        }

        public static double GetCellValueAsDouble(ICell cell)
        {
            return cell.CellType switch
            {
                CellType.String => Convert.ToDouble(cell.StringCellValue),
                CellType.Numeric => cell.NumericCellValue,
                CellType.Boolean => Convert.ToDouble(cell.BooleanCellValue),
                CellType.Blank => Convert.ToDouble(cell.ToString()),
                CellType.Formula => cell.CachedFormulaResultType switch
                {
                    CellType.String => Convert.ToDouble(cell.StringCellValue),
                    CellType.Numeric => cell.NumericCellValue,
                    CellType.Boolean => Convert.ToDouble(cell.BooleanCellValue),
                    CellType.Blank => Convert.ToDouble(cell.ToString()),
                    CellType.Error => throw new NotImplementedException(),
                    _ => throw new NotImplementedException(),
                },
                CellType.Error => throw new NotImplementedException(),
                CellType.Unknown => throw new NotImplementedException(),
                _ => throw new NotImplementedException(),
            };
        }

    }
}

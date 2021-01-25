using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TransformXLSX
{
    public static class ExtensionsHelp
    {
        public static string GetValue(this ICell cell)
        {
            switch (cell.CellType)
            {
                case NPOI.SS.UserModel.CellType.String:
                    return $"{cell.StringCellValue}";
                case NPOI.SS.UserModel.CellType.Formula:
                    return $"{cell.CellFormula}";
                case NPOI.SS.UserModel.CellType.Boolean:
                    return $"{cell.BooleanCellValue}";
                case NPOI.SS.UserModel.CellType.Numeric:
                    return $"{cell.NumericCellValue}";
                default:
                    return "";
            }
        }
    }
}

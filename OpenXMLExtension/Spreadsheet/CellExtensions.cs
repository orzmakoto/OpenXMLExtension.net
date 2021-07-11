using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OpenXMLExtension.Spreadsheet
{
    public static class CellExtensions
    {
        public static WorksheetPart WriteValue(this WorksheetPart source, string columnName, uint rowIndex, object value)
        {
            var cell = InsertCellInWorksheet(columnName, rowIndex, source, false);

            if (value.GetType() == typeof(string))
            {
                cell.CellValue = new CellValue((string)value);
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
            }
            else if (value.GetType() == typeof(int))
            {
                cell.CellValue = new CellValue((int)value);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            }
            else if (value.GetType() == typeof(DateTime))
            {
                cell.CellValue = new CellValue((DateTime)value);
                cell.DataType = new EnumValue<CellValues>(CellValues.Date);
            }
            else if (value.GetType() == typeof(DateTimeOffset))
            {
                cell.CellValue = new CellValue((DateTimeOffset)value);
                cell.DataType = new EnumValue<CellValues>(CellValues.Date);
            }
            else if (value.GetType() == typeof(bool))
            {
                cell.CellValue = new CellValue((bool)value);
                cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
            }
            else if (value.GetType() == typeof(double))
            {
                cell.CellValue = new CellValue((double)value);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            }
            else if (value.GetType() == typeof(decimal))
            {
                cell.CellValue = new CellValue((decimal)value);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            }
            else
            {
                cell.CellValue = new CellValue(value.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
            }

            return source;
        }
        public static string ReadValue(this WorksheetPart source, string columnName, uint rowIndex)
        {
            return InnerReadValue(source, columnName, rowIndex);
        }
        //public static T ReadValue<T>(this WorksheetPart source, string columnName, uint rowIndex)
        //{
        //    var retType = typeof(T);
        //    object ret;
        //    var value = InnerReadValue(source, columnName, rowIndex);
        //    if (retType == typeof(string))
        //    {
        //        ret = value;
        //        return (T)ret;
        //    }
        //    else if (retType == typeof(bool))
        //    {
        //        switch (value)
        //        {
        //            case "0": ret = false; break;
        //            case "1": ret = true; break;
        //            default: ret = false; break;

        //        }
        //        return (T)ret;
        //    }
        //    return (T)Convert.ChangeType(value, retType);
        //}
        private static string InnerReadValue(this WorksheetPart source, string columnName, uint rowIndex)
        {
            var cell = InsertCellInWorksheet(columnName, rowIndex, source, true);
            if (cell == null)
            {
                return null;
            }

            var retValeu = string.Empty;
            var innerText = cell.InnerText;
            retValeu = innerText;
            switch (cell.DataType.Value)
            {
                case CellValues.SharedString:
                    {
                        var stringTable = source.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                        if (stringTable != null)
                        {
                            retValeu = stringTable.SharedStringTable.ElementAt(int.Parse(innerText)).InnerText;
                        }
                        else
                        {
                            retValeu = innerText;
                        }
                    }; break;
            }

            return retValeu;
        }


        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart, bool readOnly)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex.ToString();

            // 指定された位置のRowオブジェクト
            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                // Rowオブジェクトがまだ存在しないときは作る
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // 指定された位置のCellオブジェクト
            Cell refCell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == cellReference);

            if (refCell != null)
            {
                return refCell; // すでに存在するので、それを返す
            }

            if (readOnly == false)
            {
                // Cellオブジェクトがまだ存在しないときは作って挿入する
                Cell nextCell = row.Elements<Cell>().FirstOrDefault(c => string.Compare(c.CellReference.Value, cellReference, true) > 0);

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, nextCell);

                worksheet.Save();
                return newCell;
            }
            else
            {
                return null;
            }
        }
    }
}

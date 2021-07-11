using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;


namespace OpenXMLExtension.Spreadsheet
{
    public static class SheetExtensions
    {
        /// <summary>
        /// シートの存在を確認します。
        /// </summary>
        /// <param name="source"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static bool HasSheet(this SpreadsheetDocument source, string sheetName)
        {
            if (source.WorkbookPart == null)
            {
                return false;
            }
            var sheets = source.WorkbookPart.Workbook.Sheets;
            if (sheets == null)
            {
                return false;
            }

            foreach (var sheet in sheets.Cast<Sheet>())
            {
                if (sheet.Name == sheetName)
                {
                    return true;
                }
            }

            return false;
        }
        /// <summary>
        /// 名前に一致するシートの取得を行います。
        /// </summary>
        /// <param name="source"></param>
        /// <param name="sheetName"></param>
        /// <param name="makeIfNot"></param>
        /// <returns></returns>
        public static WorksheetPart GetSheet(this SpreadsheetDocument source, string sheetName, bool makeIfNot = false)
        {
            if (source.HasSheet(sheetName) == false && makeIfNot == false)
            {
                throw new Exception("シート名に一致するシートが存在しません。");
            }
            else if (source.HasSheet(sheetName) == false && makeIfNot == false)
            {
                return AddNewSheet(source, sheetName);
            }

            var workbook = source.WorkbookPart.Workbook;
            var sheets = workbook.Sheets.Cast<Sheet>().ToList();

            foreach (var w in source.WorkbookPart.WorksheetParts)
            {
                string partRelationshipId = source.WorkbookPart.GetIdOfPart(w);
                var correspondingSheet = sheets.FirstOrDefault(
                    s => s.Id.HasValue && s.Id.Value == partRelationshipId);

                if (correspondingSheet != null)
                {
                    return w;
                }
            }

            throw new Exception("シート名に一致するシートが存在しません。");
        }
        /// <summary>
        /// 新規にシートを作成して末尾に追加します。
        /// </summary>
        /// <param name="source"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static WorksheetPart AddNewSheet(this SpreadsheetDocument source, string sheetName)
        {
            if (source.WorkbookPart == null)
            {
                WorkbookPart wbpart = source.AddWorkbookPart();
                wbpart.Workbook = new Workbook();
                WorksheetPart newWorksheetPart = wbpart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());
                Sheets sheets = source.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                Sheet sheet = new Sheet()
                {
                    Id = source.WorkbookPart.GetIdOfPart(newWorksheetPart),
                    SheetId = 1,
                    Name = sheetName
                };

                sheets.Append(sheet);

                wbpart.Workbook.Save();
                return newWorksheetPart;
            }
            else
            {
                if (source.HasSheet(sheetName) == true)
                {
                    throw new Exception("既に存在します。");
                }

                var newWorksheetPart = source.WorkbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());

                var sheets = source.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }
                var newSheet = new Sheet()
                {
                    Id = source.WorkbookPart.GetIdOfPart(newWorksheetPart),
                    SheetId = sheetId,
                    Name = sheetName
                };
                sheets.Append(newSheet);
                source.WorkbookPart.Workbook.Save();

                return newWorksheetPart;
            }
        }
        /// <summary>
        /// 指定されたシート名に一致するシートを削除します。
        /// 一致するシートがない場合でもエラーは発生しません。
        /// </summary>
        /// <param name="source"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static bool RemoveSheet(this SpreadsheetDocument source, string sheetName)
        {
            var sheets = source.WorkbookPart.Workbook.Sheets;
            var targetSheet = sheets.Select(i => i as Sheet).FirstOrDefault(i => i.Name == sheetName);

            if (targetSheet != null)
            {
                sheets.RemoveChild(targetSheet);
                source.WorkbookPart.Workbook.Save();

                return true;
            }
            return false;
        }



    }
}

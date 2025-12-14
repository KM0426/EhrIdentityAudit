using System;
using System.IO;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using ExcelDataReader;
using EhrIdentityAudit.Models;

namespace EhrIdentityAudit.Services;

public static class TemplateService
{
    public static async void SaveTemplates(string templateDirectory)
    {
        // エクセルファイルを3種類作成して保存する
        if (!Directory.Exists(templateDirectory))
        {
            Directory.CreateDirectory(templateDirectory);
        }
        // EHR_USERテンプレート作成
        var ehrTemplatePath = Path.Combine(templateDirectory, "カルテ利用者一覧_Template.xlsx");
        // Openize.OpenXML-SDKを使用してエクセルファイルを作成する
        using (var workbook = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create(ehrTemplatePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = workbook.AddWorkbookPart();
            workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
            var worksheetPart = workbookPart.AddNewPart<DocumentFormat.OpenXml.Packaging.WorksheetPart>();
            worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new DocumentFormat.OpenXml.Spreadsheet.SheetData());
            var sheets = workbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());
            var sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "EHR_USER"
            };
            // ヘッダー行を追加
            var sheetData = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
            var headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
            headerRow.Append(
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("RID"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("NAMEN"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("SYOKUKBN"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("LAST_YMD"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("CRT_YMD"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("STP_FLG"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("DEL_FLG"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String }
            );
            sheetData.AppendChild(headerRow);
            sheets.AppendChild(sheet);


            workbookPart.Workbook.Save();
        }

        // SYOKUIN_USERテンプレート作成
        var syokuinTemplatePath = Path.Combine(templateDirectory, "職員一覧_Template.xlsx");
        using (var workbook = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create(syokuinTemplatePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = workbook.AddWorkbookPart();
            workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
            var worksheetPart = workbookPart.AddNewPart<DocumentFormat.OpenXml.Packaging.WorksheetPart>();
            worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new DocumentFormat.OpenXml.Spreadsheet.SheetData());
            var sheets = workbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());
            var sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "SYOKUIN_USER"
            };
            // ヘッダー行を追加
            var sheetData = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
            var headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
            headerRow.Append(
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("職員番号"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("ステータス"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("地区区分"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("所属"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("職名"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("職種"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("氏名"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String }
            );
            sheetData.AppendChild(headerRow);
            sheets.AppendChild(sheet);
            workbookPart.Workbook.Save();
        }

        // HAKENITAKU_USERテンプレート作成
        var hakenitakuTemplatePath = Path.Combine(templateDirectory, "委託派遣一覧_Template.xlsx");
        using (var workbook = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create(hakenitakuTemplatePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = workbook.AddWorkbookPart();
            workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
            var worksheetPart = workbookPart.AddNewPart<DocumentFormat.OpenXml.Packaging.WorksheetPart>();
            worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new DocumentFormat.OpenXml.Spreadsheet.SheetData());
            var sheets = workbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());
            var sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "HAKENITAKU_USER"
            };
            // ヘッダー行を追加
            var sheetData = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
            var headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
            headerRow.Append(
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("No"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("区分"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("会社名"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("業種"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("身分"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("氏名"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("勤務場所"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String },
                new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("所管部署"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String }
            );
            sheetData.AppendChild(headerRow);
            sheets.AppendChild(sheet);
            workbookPart.Workbook.Save();
        }
    }
}
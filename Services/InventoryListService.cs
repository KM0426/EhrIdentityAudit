using System;
using EhrIdentityAudit.Models;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace EhrIdentityAudit.Services;

public static class InventoryListService
{
    public static async Task SaveJoinList(ObservableCollection<JoinUser> joinUsers, string folderPath)
    {
        // JoinUsersを元に全リストを生成して保存する処理を書く
        // Excelファイルとして保存する処理を書く
        await ExcelExportService.ExportInventoryListToExcel(joinUsers.ToList(), folderPath);
    }
    public static async Task SaveInventoryList(ObservableCollection<JoinUser> joinUsers, string folderPath)
    {
        // JoinUsersを元に棚卸リストを生成して保存する処理を書く
        // joinUsers.IsStoppedまたはjoinUsers.IsDeletedがtrueのユーザーは除外する
        // joinUsers.Id_Jinjiがあるものは除外する
        // joinUsers.No_Jinjigaiがあるものは除外する
        // joinUsers.LastLoginDateが実行日から2年以内は除外する
        // joinUsers.CreatedDateが実行日から2年以内は除外する
        // Idの最初に"E"がつくものは、joinUsers.LastLoginDateおよびCreatedDateが実行日から3年以内は除外する
        var inventoryList = joinUsers.Where(j =>
            !(j.IsStopped || j.IsDeleted) &&
            string.IsNullOrEmpty(j.Id_Jinji) &&
            string.IsNullOrEmpty(j.No_Jinjigai) &&
            (j.EHRUserId.StartsWith("E") ?
                ((j.LastLoginDate == null || (DateTime.Now - j.LastLoginDate.Value).TotalDays > 1095) &&
                    (j.CreatedDate == null || (DateTime.Now - j.CreatedDate.Value).TotalDays > 1095))
                :
                ((j.LastLoginDate == null || (DateTime.Now - j.LastLoginDate.Value).TotalDays > 730) &&
                    (j.CreatedDate == null || (DateTime.Now - j.CreatedDate.Value).TotalDays > 730))
            )
        ).ToList();


        // var inventoryList = joinUsers.Where(j =>
        //     !(j.IsStopped || j.IsDeleted) &&
        //     string.IsNullOrEmpty(j.Id_Jinji) &&
        //     string.IsNullOrEmpty(j.No_Jinjigai) &&
        //     (j.LastLoginDate == null || (DateTime.Now - j.LastLoginDate.Value).TotalDays > 730) &&
        //     (j.CreatedDate == null || (DateTime.Now - j.CreatedDate.Value).TotalDays > 730)
        // ).ToList();
        // // Excelファイルとして保存する処理を書く
        await ExcelExportService.ExportInventoryListToExcel(inventoryList, folderPath);
    }
}
public static class ExcelExportService
{
    public static async Task ExportInventoryListToExcel(System.Collections.Generic.List<JoinUser> inventoryList, string folderPath)
    {
        // Excelファイルに書き出す処理を書く
        // ここでは仮にファイル名を "InventoryList_yyyymmddhhMMss.xlsx" とする
        string filePath = System.IO.Path.Combine(folderPath, $"InventoryList_{DateTime.Now:yyyyMMddHHmmss}.xlsx");

        //Openize.OpenXML-SDKを使用してエクセルファイルを作成する
        await Task.Run(() =>
        {
            using (var workbook = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create(filePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
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
                    Name = "InventoryList"
                };
                // データはJoinUserのすべてのプロパティを出力する
                var sheetData = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
                // ヘッダー行を追加
                var headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                var properties = typeof(JoinUser).GetProperties();
                foreach (var prop in properties)
                {
                    headerRow.Append(new DocumentFormat.OpenXml.Spreadsheet.Cell()
                    {
                        CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(prop.Name),
                        DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                    });
                }
                sheetData?.AppendChild(headerRow);
                // データ行を追加
                foreach (var user in inventoryList)
                {
                    var dataRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    foreach (var prop in properties)
                    {
                        var value = prop.GetValue(user)?.ToString() ?? "";
                        dataRow.Append(new DocumentFormat.OpenXml.Spreadsheet.Cell()
                        {
                            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value),
                            DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                        });
                    }
                    sheetData?.AppendChild(dataRow);
                }

                sheets.AppendChild(sheet);
                workbookPart.Workbook.Save();
            }
        });
    }
}
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

public static class ReadFile
{
    public static async Task<ObservableCollection<EHR_USER>> ReadEHRUserFromExcelAsync(string filePath, IProgress<int>? progress = null)
    {
        progress?.Report(10);
        var ehrUsers = new ObservableCollection<EHR_USER>();
        if (!File.Exists(filePath)) return ehrUsers;

        DataSet? ds = null;
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
            {
                ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });
                if (ds.Tables.Count == 0) return ehrUsers;
                if (ds.Tables[0].Rows.Count == 0) return ehrUsers;
                // 
                progress?.Report(20);
                foreach (DataRow item in ds.Tables[0].Rows)
                {
                    var Id = item["RID"]?.ToString() ?? "";
                    // item["LAST_YMD"]は20130614形式の日付文字列
                    // item["LAST_YMD"]をDateTimeに変換する
                    DateTime LastLoginDate = DateTime.TryParseExact(item["LAST_YMD"]?.ToString(), "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out DateTime tempDate) ? tempDate : new DateTime(1900, 1, 1);
                    // CRT_YMDも同様に変換
                    DateTime CreatedDate = DateTime.TryParseExact(item["CRT_YMD"]?.ToString(), "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out DateTime cDate) ? cDate : new DateTime(1900, 1, 1);
                    ehrUsers.Add(new EHR_USER
                    {
                        Id = item["RID"]?.ToString() ?? "",
                        Name = item["NAMEN"]?.ToString() ?? "",
                        ShokusyuCode = item["SYOKUKBN"]?.ToString() ?? "",
                        N_LastLoginDate = !Id.StartsWith("R") || Id.StartsWith("RPA") ? LastLoginDate : new DateTime(1900, 1, 1),
                        R_LastLoginDate = Id.StartsWith("R") && !Id.StartsWith("RPA") ? LastLoginDate : new DateTime(1900, 1, 1),
                        CreatedDate = CreatedDate,
                        // STP_FLG, DEL_FLGは"0"または"1"の文字列なので、boolに変換
                        IsStopped = item["STP_FLG"]?.ToString() == "1",
                        IsDeleted = item["DEL_FLG"]?.ToString() == "1"
                    });

                }
                progress?.Report(50);
                // ehrUsersのIdで、idの頭に"R"を付与したidがあるか検索し、ある場合は、"R"無しのデータのR_LastLoginDateに"R"付きのR_LastLoginDateを設定し、"R"付きのデータを削除する
                var usersToRemove = new List<EHR_USER>();
                foreach (var user in ehrUsers)
                {
                    if (!user.Id.StartsWith("R"))
                    {
                        var rUser = ehrUsers.FirstOrDefault(u => u.Id == "R" + user.Id);
                        if (rUser != null)
                        {
                            user.R_LastLoginDate = rUser.R_LastLoginDate;
                            usersToRemove.Add(rUser);
                        }
                    }
                }
                progress?.Report(90);
                foreach (var userToRemove in usersToRemove)
                {
                    ehrUsers.Remove(userToRemove);
                }


            }
        }
        progress?.Report(100);
        return ehrUsers;
    }
    public static async Task<ObservableCollection<SYOKUIN_USER>> ReadSYOKUINUserFromExcelAsync(string filePath, IProgress<int>? progress = null)
    {
        progress?.Report(10);
        var syokuinUsers = new ObservableCollection<SYOKUIN_USER>();
        if (!File.Exists(filePath)) return syokuinUsers;

        DataSet? ds = null;
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
            {
                ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });
                if (ds.Tables.Count == 0) return syokuinUsers;
                if (ds.Tables[0].Rows.Count == 0) return syokuinUsers;
                progress?.Report(20);
                foreach (DataRow item in ds.Tables[0].Rows)
                {
                    syokuinUsers.Add(new SYOKUIN_USER
                    {
                        Id = item["職員番号"]?.ToString() ?? "",
                        Name = item["氏名"]?.ToString() ?? "",
                        Status = item["ステータス"]?.ToString() ?? "",
                        Chiku = item["地区区分"]?.ToString() ?? "",
                        Syozoku = item["所属"]?.ToString() ?? "",
                        Syokumei = item["職名"]?.ToString() ?? "",
                        ShokuShu = item["職種"]?.ToString() ?? ""
                    });

                }
                progress?.Report(100);
            }
        }
        return syokuinUsers;
    }
    public static async Task<ObservableCollection<HAKENITAKU_USER>> ReadHAKENITAKUUserFromExcelAsync(string filePath, IProgress<int>? progress = null)
    {
        progress?.Report(10);
        var hakenitakuUsers = new ObservableCollection<HAKENITAKU_USER>();
        if (!File.Exists(filePath)) return hakenitakuUsers;

        DataSet? ds = null;
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
            {
                ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });
                if (ds.Tables.Count == 0) return hakenitakuUsers;
                if (ds.Tables[0].Rows.Count == 0) return hakenitakuUsers;
                progress?.Report(20);
                foreach (DataRow item in ds.Tables[0].Rows)
                {
                    hakenitakuUsers.Add(new HAKENITAKU_USER
                    {
                        No = item["No"]?.ToString() ?? "",
                        Kubun = item["区分"]?.ToString() ?? "",
                        Company = item["会社名"]?.ToString() ?? "",
                        Gyousyu = item["業種"]?.ToString() ?? "",
                        Mibun = item["身分"]?.ToString() ?? "",
                        Name = item["氏名"]?.ToString() ?? "",
                        Kinmubasyo = item["勤務場所"]?.ToString() ?? "",
                        Kanribumon = item["所管部署"]?.ToString() ?? ""
                    });

                }
                progress?.Report(100);
            }
        }
        return hakenitakuUsers;
    }
}
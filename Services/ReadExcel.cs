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
    public static async Task<ServiceResult<ObservableCollection<EHR_USER>>> ReadEHRUserFromExcelAsync(string filePath, IProgress<int>? progress = null)
    {
        progress?.Report(0);
        var ehrUsers = new ObservableCollection<EHR_USER>();
        if (!File.Exists(filePath))
        {
            return ServiceResult<ObservableCollection<EHR_USER>>.Fail($"ファイルが存在しません: {filePath}");
        }

        DataSet? ds = null;
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        try
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                progress?.Report(10);
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });
                    if (ds.Tables.Count == 0)
                    {
                        progress?.Report(0);
                        return ServiceResult<ObservableCollection<EHR_USER>>.Fail($"データが存在しません: {filePath}");
                    }
                    if (ds.Tables[0].Rows.Count == 0)
                    {
                        progress?.Report(0);
                        return ServiceResult<ObservableCollection<EHR_USER>>.Fail($"データが存在しません: {filePath}");
                    }
                    //dsのヘッダーに、「RID、NAMEN、SYOKUKBN、LAST_YMD、CRT_YMD、STP_FLG、DEL_FLG」が含まれているかチェック
                    var requiredColumns = new List<string> { "RID", "NAMEN", "SYOKUKBN", "LAST_YMD", "CRT_YMD", "STP_FLG", "DEL_FLG" };
                    foreach (var column in requiredColumns)
                    {
                        if (!ds.Tables[0].Columns.Contains(column))
                        {
                            progress?.Report(0);
                            return ServiceResult<ObservableCollection<EHR_USER>>.Fail($"必要な列が存在しません: {column}\n\r\n\r必要な列: {string.Join(", ", requiredColumns)}");
                        }
                    }


                    foreach (DataRow item in ds.Tables[0].Rows)
                    {
                        // progress を10から最大20まで、Rowsカウントで進める
                        progress?.Report(10 + (10 * ehrUsers.Count / ds.Tables[0].Rows.Count));

                        var Id = item["RID"]?.ToString() ?? "";
                        // item["LAST_YMD"]は20130614形式の日付文字列
                        // item["LAST_YMD"]をDateTimeに変換する
                        DateTime LastLoginDate = DateTime.TryParseExact(item["LAST_YMD"]?.ToString(), "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out DateTime tempDate) ? tempDate : new DateTime(1900, 1, 1);
                        // CRT_YMDも同様に変換
                        DateTime CreatedDate = DateTime.TryParseExact(item["CRT_YMD"]?.ToString(), "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out DateTime cDate) ? cDate : new DateTime(1900, 1, 1);
                        var name = item["NAMEN"]?.ToString()?.Trim() ?? "";
                        // 含まれている場合はスキップ
                        if (name.Contains("部門") || name.Contains("病床") || name.Contains("代行") || name.Contains("テスト") || name.Contains("ＢＹＯＲＩ") || name.Contains("複製") || name.Contains("通治") || name.Contains("ＰＴＣＤ") || name.Contains("共通") || name.Contains("外来") || name.Contains("緩和") || name.Contains("肝胆") || name.Contains("生検") || name.Contains("ＥＲＰ－ＣＴ") || name.Contains("ＥＭＲ") || name.Contains("ＥＲＣＰ") || name.Contains("ＤＵＭＭＹ") || name.Contains("薬剤") || name.Contains("医師") || name.Contains("移行") || name.Contains("検証") || name.Contains("遡及") || name.Contains("紹介") || name.Contains("スーパー") || name.Contains("一括") || name.Contains("レジメン") || name.Contains("診療") || name.Contains("医事") || name.Contains("栄養") || name.Contains("言語") || name.Contains("作業") || name.Contains("理学") || name.Contains("放射") || name.Contains("輸血") || name.Contains("検査") || name.Contains("歯科") || name.Contains("ＴＥＳＴ") || name.Contains("看護") || name.Contains("ＯＳＳＣ") || name.Contains("手術") || name.Contains("栄養") || name.Contains("管栄") || name.Contains("利用者") || name.Contains("ラベル") || name.Contains("権限") || name.Contains("外来") || name.Contains("モニタ") || name.Contains("協力者") || name.Contains("メディカル") || name.Contains("じょぶ") || name.Contains("事前入力") || name.Contains("病院") || name.Contains("スキャナ") || name.Contains("ＩＪＩ") || name.Contains("Ｉｍａｇｅ") || name.Contains("予約") || name.Contains("ＩＢＭ") || name.Contains("アイビーエム") || name.Contains("ＨＩＲＩＭ") || name.Contains("ＨＩＬＡＳ") || name.Contains("ＨＩＴＡＣＨＩ") || name.Contains("研修") || name.Contains("富士通") || name.Contains("ＦＪＪ") || name.Contains("東－") || name.Contains("管理者"))
                        {
                            continue;
                        }

                        ehrUsers.Add(new EHR_USER
                        {
                            Id = item["RID"]?.ToString()?.Trim() ?? "",
                            Name = name,
                            ShokusyuCode = item["SYOKUKBN"]?.ToString() ?? "",
                            N_LastLoginDate = !Id.StartsWith("R") || Id.StartsWith("RPA") ? LastLoginDate : new DateTime(1900, 1, 1),
                            R_LastLoginDate = Id.StartsWith("R") && !Id.StartsWith("RPA") ? LastLoginDate : new DateTime(1900, 1, 1),
                            CreatedDate = CreatedDate,
                            // STP_FLG, DEL_FLGは"0"または"1"の文字列なので、boolに変換
                            IsStopped = item["STP_FLG"]?.ToString() == "1",
                            IsDeleted = item["DEL_FLG"]?.ToString() == "1"
                        });

                    }
                    progress?.Report(20);
                    // ehrUsersのIdで、idの頭に"R"を付与したidがあるか検索し、ある場合は、"R"無しのデータのR_LastLoginDateに"R"付きのR_LastLoginDateを設定し、"R"付きのデータを削除する
                    var rUserDict = ehrUsers.Where(u => u.Id.StartsWith("R") && !u.Id.StartsWith("RPA"))
                                            .ToDictionary(u => u.Id.Substring(1), u => u);
                    var usersToRemove = new List<EHR_USER>();
                    var processedCount = 0;

                    foreach (var user in ehrUsers)
                    {
                        // progress を20から最大95まで、処理済みカウントで進める
                        if (++processedCount % 100 == 0)
                        {
                            progress?.Report(20 + (75 * processedCount / ehrUsers.Count));
                        }

                        if (!user.Id.StartsWith("R") && rUserDict.TryGetValue(user.Id, out var rUser))
                        {
                            user.R_LastLoginDate = rUser.R_LastLoginDate;
                            usersToRemove.Add(rUser);
                        }
                    }
                    progress?.Report(95);
                    foreach (var userToRemove in usersToRemove)
                    {
                        ehrUsers.Remove(userToRemove);
                    }
                }
            }
            progress?.Report(100);

            return ServiceResult<ObservableCollection<EHR_USER>>.Success(ehrUsers);
        }
        catch (Exception ex)
        {
            // ここでは UI は触らない。ログだけ取るイメージ
            // Log(ex);
            return ServiceResult<ObservableCollection<EHR_USER>>.Fail($"取得に失敗しました: {ex.Message}");
        }
    }
    public static async Task<ServiceResult<ObservableCollection<SYOKUIN_USER>>> ReadSYOKUINUserFromExcelAsync(string filePath, IProgress<int>? progress = null)
    {
        progress?.Report(0);
        var syokuinUsers = new ObservableCollection<SYOKUIN_USER>();
        if (!File.Exists(filePath))
        {
            return ServiceResult<ObservableCollection<SYOKUIN_USER>>.Fail($"ファイルが存在しません: {filePath}");
        }

        DataSet? ds = null;
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        try
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                progress?.Report(10);
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });
                    if (ds.Tables.Count == 0)
                    {
                        progress?.Report(0);
                        return ServiceResult<ObservableCollection<SYOKUIN_USER>>.Fail($"データが存在しません: {filePath}");
                    }
                    if (ds.Tables[0].Rows.Count == 0)
                    {
                        progress?.Report(0);
                        return ServiceResult<ObservableCollection<SYOKUIN_USER>>.Fail($"データが存在しません: {filePath}");
                    }
                    // ヘッダーに「職員番号、氏名、ステータス、地区区分、所属、職名、職種」が含まれているかチェック
                    var requiredColumns = new List<string> { "職員番号", "氏名", "ステータス", "地区区分", "所属", "職名", "職種" };
                    foreach (var column in requiredColumns)
                    {
                        if (!ds.Tables[0].Columns.Contains(column))
                        {
                            progress?.Report(0);
                            return ServiceResult<ObservableCollection<SYOKUIN_USER>>.Fail($"必要な列が存在しません: {column}\n\r\n\r必要な列: {string.Join(", ", requiredColumns)}");
                        }
                    }

                    progress?.Report(20);
                    foreach (DataRow item in ds.Tables[0].Rows)
                    {
                        var name = item["氏名"]?.ToString() ?? "";
                        // nameに（または(が含まれている場合、そこから後ろを削除する
                        if (name.Contains("（"))
                        {
                            name = name.Substring(0, name.IndexOf("（")).Trim();
                        }
                        else if (name.Contains("("))
                        {
                            name = name.Substring(0, name.IndexOf("(")).Trim();
                        }
                        syokuinUsers.Add(new SYOKUIN_USER
                        {
                            Id = item["職員番号"]?.ToString() ?? "",
                            Name = name.Trim(),
                            Status = item["ステータス"]?.ToString() ?? "",
                            Chiku = item["地区区分"]?.ToString() ?? "",
                            Syozoku = item["所属"]?.ToString() ?? "",
                            Syokumei = item["職名"]?.ToString() ?? "",
                            ShokuShu = item["職種"]?.ToString() ?? ""
                        });
                    }
                }
            }
            progress?.Report(100);
        }
        catch (Exception ex)
        {
            // ここでは UI は触らない。ログだけ取るイメージ
            // Log(ex);
            return ServiceResult<ObservableCollection<SYOKUIN_USER>>.Fail($"取得に失敗しました: {ex.Message}");
        }
        return ServiceResult<ObservableCollection<SYOKUIN_USER>>.Success(syokuinUsers);

    }
    public static async Task<ServiceResult<ObservableCollection<HAKENITAKU_USER>>> ReadHAKENITAKUUserFromExcelAsync(string filePath, IProgress<int>? progress = null)
    {
        progress?.Report(0);
        var hakenitakuUsers = new ObservableCollection<HAKENITAKU_USER>();
        if (!File.Exists(filePath))
        {
            return ServiceResult<ObservableCollection<HAKENITAKU_USER>>.Fail($"ファイルが存在しません: {filePath}");
        }

        DataSet? ds = null;
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        try
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                progress?.Report(10);
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });
                    if (ds.Tables.Count == 0)
                    {
                        progress?.Report(0);
                        return ServiceResult<ObservableCollection<HAKENITAKU_USER>>.Fail($"データが存在しません: {filePath}");
                    }
                    if (ds.Tables[0].Rows.Count == 0)
                    {
                        progress?.Report(0);
                        return ServiceResult<ObservableCollection<HAKENITAKU_USER>>.Fail($"データが存在しません: {filePath}");
                    }
                    // ヘッダーに「No、区分、会社名、業種、身分、氏名、勤務場所、所管部署」が含まれているかチェック
                    var requiredColumns = new List<string> { "No", "区分", "会社名", "業種", "身分", "氏名", "勤務場所", "所管部署" };
                    foreach (var column in requiredColumns)
                    {
                        if (!ds.Tables[0].Columns.Contains(column))
                        {
                            progress?.Report(0);
                            return ServiceResult<ObservableCollection<HAKENITAKU_USER>>.Fail($"必要な列が存在しません: {column}\n\r\n\r必要な列: {string.Join(", ", requiredColumns)}");
                        }
                    }
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
                            Name = item["氏名"]?.ToString()?.Trim() ?? "",
                            Kinmubasyo = item["勤務場所"]?.ToString() ?? "",
                            Kanribumon = item["所管部署"]?.ToString() ?? ""
                        });
                    }
                }
            }
            progress?.Report(100);
        }
        catch (Exception ex)
        {
            // ここでは UI は触らない。ログだけ取るイメージ
            // Log(ex);
            return ServiceResult<ObservableCollection<HAKENITAKU_USER>>.Fail($"取得に失敗しました: {ex.Message}");
        }
        return ServiceResult<ObservableCollection<HAKENITAKU_USER>>.Success(hakenitakuUsers);
    }
}
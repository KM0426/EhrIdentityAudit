using System;
using EhrIdentityAudit.Models;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
namespace EhrIdentityAudit.Services;

public static class AuditService
{
    public static async Task<ObservableCollection<JoinUser>> ExecuteAudit(ObservableCollection<EHR_USER> ehrUsers, ObservableCollection<SYOKUIN_USER> syokuinUsers, ObservableCollection<HAKENITAKU_USER> hakenitakuUsers, IProgress<int>? progress = null)
    {
        progress?.Report(0);
        // 突合ロジックをここに実装する
        var result = new ObservableCollection<JoinUser>();

        // EHR_USERsを元にして、SYOKUIN_USERsおよびHAKENITAKU_USERsと氏名突合を行う処理を書く
        // 突合結果をJoinUsersに格納する処理を書く
        // 突合時に氏名の標準化を行う　NameStandardization.ConvertNameを使用
        var EHR_List = ehrUsers.ToList();
        var SYOKUIN_List = syokuinUsers.ToList();
        var HAKENITAKU_List = hakenitakuUsers.ToList();
        int total = EHR_List.Count;
        int count = 0;
        // Pre-compute standardized names and create lookup dictionaries for O(1) access
        var syokuinLookup = SYOKUIN_List
            .GroupBy(s => NameStandardization.ConvertName(s.Name))
            .ToDictionary(g => g.Key, g => g.FirstOrDefault());

        var hakenitakuLookup = HAKENITAKU_List
            .GroupBy(h => NameStandardization.ConvertName(h.Name))
            .ToDictionary(g => g.Key, g => g.FirstOrDefault());

        var ehrNameLookup = EHR_List
            .Where(e => !e.Id.StartsWith("R") && !e.Id.StartsWith("E"))
            .GroupBy(e => NameStandardization.ConvertName(e.Name))
            .ToDictionary(g => g.Key, g => g.ToList());

        foreach (var ehr in EHR_List)
        {
            count++;
            progress?.Report((int)((count / (double)total) * 100));
            var standardizedEhrName = NameStandardization.ConvertName(ehr.Name);

            syokuinLookup.TryGetValue(standardizedEhrName, out var matchedSyokuin);
            hakenitakuLookup.TryGetValue(standardizedEhrName, out var matchedHakenitaku);
            bool douseidoumeiCheck = false;
            if (ehr.Id.StartsWith("R") || ehr.Id.StartsWith("E"))
            {
            }
            else
            {
                douseidoumeiCheck = ehrNameLookup.TryGetValue(standardizedEhrName, out var matches) && matches.Any(e => e.Id != ehr.Id && !e.Id.StartsWith("R") && !e.Id.StartsWith("E") && !(ehr.IsStopped || ehr.IsDeleted) && !(e.IsStopped || e.IsDeleted));
            }
            if (douseidoumeiCheck)
            {
                // SYOKUIN_Listに同姓同名がいるかチェックし、いたら全て取得

                if (SYOKUIN_List.FindAll(s => NameStandardization.ConvertName(s.Name) == standardizedEhrName).Count > 1)
                {
                    var syokuinMatches = SYOKUIN_List.FindAll(s => NameStandardization.ConvertName(s.Name) == standardizedEhrName);
                    foreach (var item in syokuinMatches)
                    {
                        matchedSyokuin = item;
                        var joinUserDousei = new JoinUser
                        {
                            EHRUserId = ehr.Id,
                            EHRUserName = ehr.Name,
                            IsDouseiDoumei = douseidoumeiCheck,
                            ShokusyuCode = ehr.ShokusyuCode,
                            N_LastLoginDate = ehr.N_LastLoginDate,
                            R_LastLoginDate = ehr.R_LastLoginDate,
                            CreatedDate = ehr.CreatedDate,
                            IsStopped = ehr.IsStopped,
                            IsDeleted = ehr.IsDeleted,
                            Id_Jinji = matchedSyokuin?.Id ?? "",
                            Name_Jinji = matchedSyokuin?.Name ?? "",
                            Status_Jinji = matchedSyokuin?.Status ?? "",
                            Chiku_Jinji = matchedSyokuin?.Chiku ?? "",
                            Syozoku_Jinji = matchedSyokuin?.Syozoku ?? "",
                            Syokumei_Jinji = matchedSyokuin?.Syokumei ?? "",
                            ShokuShu_Jinji = matchedSyokuin?.ShokuShu ?? "",
                            No_Jinjigai = matchedHakenitaku?.No ?? "",
                            Kubun_Jinjigai = matchedHakenitaku?.Kubun ?? "",
                            Company_Jinjigai = matchedHakenitaku?.Company ?? "",
                            Gyousyu_Jinjigai = matchedHakenitaku?.Gyousyu ?? "",
                            Mibun_Jinjigai = matchedHakenitaku?.Mibun ?? "",
                            Name_itakuhaken_Jinjigai = matchedHakenitaku?.Name ?? "",
                            Kinmubasyo_Jinjigai = matchedHakenitaku?.Kinmubasyo ?? "",
                            Kanribumon_Jinjigai = matchedHakenitaku?.Kanribumon ?? ""
                        };
                        result.Add(joinUserDousei);
                    }
                }
                // hakenitakuLookupに同姓同名がいるかチェックし、いたら全て取得

                if (HAKENITAKU_List.FindAll(h => NameStandardization.ConvertName(h.Name) == standardizedEhrName).Count > 1)
                {
                    var hakenitakuMatches = HAKENITAKU_List.FindAll(h => NameStandardization.ConvertName(h.Name) == standardizedEhrName);
                    foreach (var item in hakenitakuMatches)
                    {
                        matchedHakenitaku = item;
                        var joinUserDousei = new JoinUser
                        {
                            EHRUserId = ehr.Id,
                            EHRUserName = ehr.Name,
                            IsDouseiDoumei = douseidoumeiCheck,
                            ShokusyuCode = ehr.ShokusyuCode,
                            N_LastLoginDate = ehr.N_LastLoginDate,
                            R_LastLoginDate = ehr.R_LastLoginDate,
                            CreatedDate = ehr.CreatedDate,
                            IsStopped = ehr.IsStopped,
                            IsDeleted = ehr.IsDeleted,
                            Id_Jinji = matchedSyokuin?.Id ?? "",
                            Name_Jinji = matchedSyokuin?.Name ?? "",
                            Status_Jinji = matchedSyokuin?.Status ?? "",
                            Chiku_Jinji = matchedSyokuin?.Chiku ?? "",
                            Syozoku_Jinji = matchedSyokuin?.Syozoku ?? "",
                            Syokumei_Jinji = matchedSyokuin?.Syokumei ?? "",
                            ShokuShu_Jinji = matchedSyokuin?.ShokuShu ?? "",
                            No_Jinjigai = matchedHakenitaku?.No ?? "",
                            Kubun_Jinjigai = matchedHakenitaku?.Kubun ?? "",
                            Company_Jinjigai = matchedHakenitaku?.Company ?? "",
                            Gyousyu_Jinjigai = matchedHakenitaku?.Gyousyu ?? "",
                            Mibun_Jinjigai = matchedHakenitaku?.Mibun ?? "",
                            Name_itakuhaken_Jinjigai = matchedHakenitaku?.Name ?? "",
                            Kinmubasyo_Jinjigai = matchedHakenitaku?.Kinmubasyo ?? "",
                            Kanribumon_Jinjigai = matchedHakenitaku?.Kanribumon ?? ""
                        };
                        result.Add(joinUserDousei);
                    }
                }
                continue;
            }

            var joinUser = new JoinUser
            {
                EHRUserId = ehr.Id,
                EHRUserName = ehr.Name,
                IsDouseiDoumei = douseidoumeiCheck,
                ShokusyuCode = ehr.ShokusyuCode,
                N_LastLoginDate = ehr.N_LastLoginDate,
                R_LastLoginDate = ehr.R_LastLoginDate,
                CreatedDate = ehr.CreatedDate,
                IsStopped = ehr.IsStopped,
                IsDeleted = ehr.IsDeleted,
                Id_Jinji = matchedSyokuin?.Id ?? "",
                Name_Jinji = matchedSyokuin?.Name ?? "",
                Status_Jinji = matchedSyokuin?.Status ?? "",
                Chiku_Jinji = matchedSyokuin?.Chiku ?? "",
                Syozoku_Jinji = matchedSyokuin?.Syozoku ?? "",
                Syokumei_Jinji = matchedSyokuin?.Syokumei ?? "",
                ShokuShu_Jinji = matchedSyokuin?.ShokuShu ?? "",
                No_Jinjigai = matchedHakenitaku?.No ?? "",
                Kubun_Jinjigai = matchedHakenitaku?.Kubun ?? "",
                Company_Jinjigai = matchedHakenitaku?.Company ?? "",
                Gyousyu_Jinjigai = matchedHakenitaku?.Gyousyu ?? "",
                Mibun_Jinjigai = matchedHakenitaku?.Mibun ?? "",
                Name_itakuhaken_Jinjigai = matchedHakenitaku?.Name ?? "",
                Kinmubasyo_Jinjigai = matchedHakenitaku?.Kinmubasyo ?? "",
                Kanribumon_Jinjigai = matchedHakenitaku?.Kanribumon ?? ""
            };

            result.Add(joinUser);
        }

        progress?.Report(100);
        return result;
    }
}
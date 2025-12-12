using System;
namespace EhrIdentityAudit.Models;

public class JoinUser
{
    public string EHRUserId { get; set; } = "";
    public string EHRUserName { get; set; } = "";
    public bool IsDouseiDoumei { get; set; } = false;
    public string ShokusyuCode { get; set; } = "";
    public DateTime? LastLoginDate { get { return R_LastLoginDate > N_LastLoginDate ? R_LastLoginDate : N_LastLoginDate; } }
    public DateTime? N_LastLoginDate { get; set; } = new DateTime(1900, 1, 1);
    public DateTime? R_LastLoginDate { get; set; } = new DateTime(1900, 1, 1);
    public DateTime? CreatedDate { get; set; } = new DateTime(1900, 1, 1);
    public bool IsStopped { get; set; } = false;
    public bool IsDeleted { get; set; } = false;

    public string Id_Jinji { get; set; } = "";
    public string Name_Jinji { get; set; } = "";
    public string Status_Jinji { get; set; } = "";
    public string Chiku_Jinji { get; set; } = "";
    public string Syozoku_Jinji { get; set; } = "";
    public string Syokumei_Jinji { get; set; } = "";
    public string ShokuShu_Jinji { get; set; } = "";

    public string No_Jinjigai { get; set; } = "";
    public string Kubun_Jinjigai { get; set; } = "";
    public string Company_Jinjigai { get; set; } = "";
    public string Gyousyu_Jinjigai { get; set; } = "";
    public string Mibun_Jinjigai { get; set; } = "";
    public string Name_itakuhaken_Jinjigai { get; set; } = "";
    public string Kinmubasyo_Jinjigai { get; set; } = "";
    public string Kanribumon_Jinjigai { get; set; } = "";

}

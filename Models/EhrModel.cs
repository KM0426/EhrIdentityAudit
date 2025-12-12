using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace EhrIdentityAudit.Models;

public class EHR_USER
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string ShokusyuCode { get; set; } = "";
    public DateTime? LastLoginDate { get { return R_LastLoginDate > N_LastLoginDate ? R_LastLoginDate : N_LastLoginDate; } }
    public DateTime? N_LastLoginDate { get; set; } = new DateTime(1900, 1, 1);
    public DateTime? R_LastLoginDate { get; set; } = new DateTime(1900, 1, 1);
    public DateTime? CreatedDate { get; set; } = new DateTime(1900, 1, 1);
    public bool IsStopped { get; set; } = false;
    public bool IsDeleted { get; set; } = false;

}




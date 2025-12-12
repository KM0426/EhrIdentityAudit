using EhrIdentityAudit.Models;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
namespace EhrIdentityAudit.Services;

public static class AuditService
{
    public static async Task<ObservableCollection<JoinUser>> ExecuteAudit(ObservableCollection<EHR_USER> ehrUsers, ObservableCollection<SYOKUIN_USER> syokuinUsers, ObservableCollection<HAKENITAKU_USER> hakenitakuUsers)
    {
        // 突合ロジックをここに実装する
        var result = new ObservableCollection<JoinUser>();

        // EHR_USERsを元にして、SYOKUIN_USERsおよびHAKENITAKU_USERsと氏名突合を行う処理を書く
        // 突合結果をJoinUsersに格納する処理を書く
        // 突合時に氏名の標準化を行う

        return result;
    }
}
using System;
using System.Diagnostics;

namespace EhrIdentityAudit.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    public string Greeting { get; } = "Welcome to Avalonia!";
    // File → Read
    public void ReadCommand()
    {
        // TODO: 読み込み処理を書く
        Debug.WriteLine("ReadCommand 実行");
        
        // TODO: ファイル選択ダイアログを開く処理を書く
        // ファイルはエクセルファイルを想定
        // ファイルは三種類のパターンがある
        // 1. 電子カルテ利用者一覧、２．職員一覧、３．委託派遣一覧
        
    }

    // File → Exit
    public void ExitCommand()
    {
        Environment.Exit(0);
    }

    // Config → Config
    public void ShowConfigCommand()
    {
        // TODO: 設定画面を開く処理
        Debug.WriteLine("ShowConfigCommand 実行");
    }
}

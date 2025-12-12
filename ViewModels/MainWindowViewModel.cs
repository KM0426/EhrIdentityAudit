using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using EhrIdentityAudit.Models;
using EhrIdentityAudit.Services;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace EhrIdentityAudit.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{



    // ProgressValue
    private int _auditProgressValue;
    public int AuditProgressValue
    {
        get => _auditProgressValue;
        set
        {
            _auditProgressValue = value;
            OnPropertyChanged(nameof(AuditProgressValue));
        }
    }
    private int _ehrProgressValue;
    public int EhrProgressValue
    {
        get => _ehrProgressValue;
        set
        {
            _ehrProgressValue = value;
            OnPropertyChanged(nameof(EhrProgressValue));
        }
    }
    private int _syokuinProgressValue;
    public int SyokuinProgressValue
    {
        get => _syokuinProgressValue;
        set
        {
            _syokuinProgressValue = value;
            OnPropertyChanged(nameof(SyokuinProgressValue));
        }
    }
    private int _hakenitakuProgressValue;
    public int HakenitakuProgressValue
    {
        get => _hakenitakuProgressValue;
        set
        {
            _hakenitakuProgressValue = value;
            OnPropertyChanged(nameof(HakenitakuProgressValue));
        }
    }
    private ObservableCollection<EHR_USER> _ehrUsers = new ObservableCollection<EHR_USER>();
    public ObservableCollection<EHR_USER> EHR_USERs
    {
        get => _ehrUsers;
        set
        {
            _ehrUsers = value;
            OnPropertyChanged(nameof(EHR_USERs));
        }
    }
    private ObservableCollection<SYOKUIN_USER> _syokuinUsers = new ObservableCollection<SYOKUIN_USER>();
    public ObservableCollection<SYOKUIN_USER> SYOKUIN_USERs
    {
        get => _syokuinUsers;
        set
        {
            _syokuinUsers = value;
            OnPropertyChanged(nameof(SYOKUIN_USERs));
        }
    }
    private ObservableCollection<HAKENITAKU_USER> _hakenitakuUsers = new ObservableCollection<HAKENITAKU_USER>();
    public ObservableCollection<HAKENITAKU_USER> HAKENITAKU_USERs
    {
        get => _hakenitakuUsers;
        set
        {
            _hakenitakuUsers = value;
            OnPropertyChanged(nameof(HAKENITAKU_USERs));
        }
    }
    private ObservableCollection<JoinUser> _joinUsers = new ObservableCollection<JoinUser>();
    public ObservableCollection<JoinUser> JoinUsers
    {
        get => _joinUsers;
        set
        {
            _joinUsers = value;
            OnPropertyChanged(nameof(JoinUsers));
        }
    }

    public MainWindowViewModel()
    {

    }

    public string Greeting { get; } = "Welcome to Avalonia!";
    // File → Read
    public void LoadEHRUserCommand()
    {
        Debug.WriteLine("LoadEHRUserCommand 実行");
        // ファイル選択ダイアログを開く処理を書く
        var filePath = OpenFileDialog("電子カルテ利用者一覧ファイルを選択してください");
        if (!string.IsNullOrEmpty(filePath))
        {
            EHR_USERs.Clear();

            LoadEHRUserAsync(filePath);
        }
    }
    public void LoadSYOKUINUserCommand()
    {
        Debug.WriteLine("LoadSYOKUINUserCommand 実行");
        // ファイル選択ダイアログを開く処理を書く
        var filePath = OpenFileDialog("職員一覧ファイルを選択してください");
        if (!string.IsNullOrEmpty(filePath))
        {
            SYOKUIN_USERs.Clear();

            LoadSYOKUINUserAsync(filePath);
        }
    }
    public void LoadHAKENITAKUUserCommand()
    {
        Debug.WriteLine("LoadHAKENITAKUUserCommand 実行");
        // ファイル選択ダイアログを開く処理を書く
        var filePath = OpenFileDialog("委託派遣一覧ファイルを選択してください");
        if (!string.IsNullOrEmpty(filePath))
        {
            HAKENITAKU_USERs.Clear();

            LoadHAKENITAKUUserAsync(filePath);
        }
    }
    private async void LoadEHRUserAsync(string filePath)
    {
        // ReadEHRUserFromExcelAsync内の処理状況をUIに反映させるために、awaitを使用して非同期で実行
        // 読み込み中はUIがフリーズしないようにする
        // ProgressValueを更新する処理も追加する
        var progress = new Progress<int>(value => EhrProgressValue = value);
        var users = await Task.Run(() => ReadFile.ReadEHRUserFromExcelAsync(filePath, progress));

        EHR_USERs = new ObservableCollection<EHR_USER>(users);
    }
    private async void LoadSYOKUINUserAsync(string filePath)
    {
        var progress = new Progress<int>(value => SyokuinProgressValue = value);
        var users = await Task.Run(() => ReadFile.ReadSYOKUINUserFromExcelAsync(filePath, progress));

        SYOKUIN_USERs = new ObservableCollection<SYOKUIN_USER>(users);
    }
    private async void LoadHAKENITAKUUserAsync(string filePath)
    {
        var progress = new Progress<int>(value => HakenitakuProgressValue = value);
        var users = await Task.Run(() => ReadFile.ReadHAKENITAKUUserFromExcelAsync(filePath, progress));

        HAKENITAKU_USERs = new ObservableCollection<HAKENITAKU_USER>(users);
    }
    private string OpenFileDialog(string Title)
    {
        var window = App.Current.ApplicationLifetime is Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop
            ? desktop.MainWindow
            : null;
        var filePath = string.Empty;
        var storage = window.StorageProvider;
        storage.OpenFilePickerAsync(new Avalonia.Platform.Storage.FilePickerOpenOptions
        {
            Title = Title,
            AllowMultiple = false,
            FileTypeFilter = new List<Avalonia.Platform.Storage.FilePickerFileType>
            {
                new Avalonia.Platform.Storage.FilePickerFileType("Excel Files")
                {
                    Patterns = new List<string> { "*.xlsx", "*.xls" }
                },
                new Avalonia.Platform.Storage.FilePickerFileType("All Files")
                {
                    Patterns = new List<string> { "*.*" }
                }
            }
        }).ContinueWith(t =>
        {
            var files = t.Result;
            if (files != null && files.Count > 0)
            {
                filePath = files[0].Path.LocalPath;
            }
        }).Wait();
        return filePath;
    }

    public async void ExecuteAuditCommand()
    {
        Debug.WriteLine("ExecuteAuditCommand 実行");

        var progress = new Progress<int>(value => AuditProgressValue = value);
        var join = await Task.Run(() => AuditService.ExecuteAudit(EHR_USERs, SYOKUIN_USERs, HAKENITAKU_USERs, progress));
        JoinUsers = new ObservableCollection<JoinUser>(join);
    }
    public void ReadCommand()
    {
        // TODO: 読み込み処理を書く
        Debug.WriteLine("ReadCommand 実行");
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

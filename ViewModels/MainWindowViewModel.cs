using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using EhrIdentityAudit.Models;
using EhrIdentityAudit.Services;
using Microsoft.WindowsAPICodePack.Dialogs;
using Microsoft.VisualBasic;
using Microsoft.WindowsAPICodePack.Shell.Interop;

namespace EhrIdentityAudit.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{


    public bool HasJoinUser
    {
        get => JoinUsers.Count > 0;
    }
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
            OnPropertyChanged(nameof(IsLoadedEHRUser));
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
            OnPropertyChanged(nameof(HasJoinUser));
        }
    }

    public MainWindowViewModel()
    {

    }

    public string Greeting { get; } = "Welcome to Avalonia!";

    public void DownloadJoinListCommand()
    {

        Debug.WriteLine("DownloadJoinListCommand 実行");
        var window = App.Current?.ApplicationLifetime is Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop
            ? desktop.MainWindow
            : null;
        var storage = window?.StorageProvider;
        var folderPath = string.Empty;
        storage?.OpenFolderPickerAsync(new Avalonia.Platform.Storage.FolderPickerOpenOptions
        {
            Title = "保存先フォルダを選択してください"
        }).ContinueWith(t =>
        {
            var folders = t.Result;
            if (folders != null && folders.Count > 0)
            {
                folderPath = folders[0].Path.LocalPath;
            }
        }).Wait();
        if (!string.IsNullOrEmpty(folderPath))
        {
            InventoryListService.SaveJoinList(JoinUsers, folderPath);
            // msgbox表示
            var msgBox = new Avalonia.Controls.Window
            {
                Title = "完了",
                Width = 300,
                Height = 150,
                Content = new Avalonia.Controls.TextBlock
                {
                    Text = "全リストのダウンロードが完了しました。",
                    VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                    HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap
                }
            };
            msgBox.ShowDialog(window);
        }

    }
    public void DownloadInventoryListCommand()
    {

        Debug.WriteLine("DownloadInventoryListCommand 実行");
        var window = App.Current?.ApplicationLifetime is Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop
            ? desktop.MainWindow
            : null;
        var storage = window?.StorageProvider;
        var folderPath = string.Empty;
        storage?.OpenFolderPickerAsync(new Avalonia.Platform.Storage.FolderPickerOpenOptions
        {
            Title = "保存先フォルダを選択してください"
        }).ContinueWith(t =>
        {
            var folders = t.Result;
            if (folders != null && folders.Count > 0)
            {
                folderPath = folders[0].Path.LocalPath;
            }
        }).Wait();
        if (!string.IsNullOrEmpty(folderPath))
        {
            InventoryListService.SaveInventoryList(JoinUsers, folderPath);
            var msgBox = new Avalonia.Controls.Window
            {
                Title = "完了",
                Width = 300,
                Height = 150,
                Content = new Avalonia.Controls.TextBlock
                {
                    Text = "棚卸対象リストのダウンロードが完了しました。",
                    VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                    HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap
                }
            };
            msgBox.ShowDialog(window);
        }

    }
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
    private bool _isLoadingEHRUser = false;
    public bool IsLoadingEHRUser
    {
        get => !_isLoadingEHRUser;
        set
        {
            _isLoadingEHRUser = value;
            OnPropertyChanged(nameof(IsLoadingEHRUser));
        }
    }
    private bool _isLoadingSYOKUINUser = false;
    public bool IsLoadingSYOKUINUser
    {
        get => !_isLoadingSYOKUINUser;
        set
        {
            _isLoadingSYOKUINUser = value;
            OnPropertyChanged(nameof(IsLoadingSYOKUINUser));
        }
    }
    private bool _isLoadingHAKENITAKUUser = false;
    public bool IsLoadingHAKENITAKUUser
    {
        get => !_isLoadingHAKENITAKUUser;
        set
        {
            _isLoadingHAKENITAKUUser = value;
            OnPropertyChanged(nameof(IsLoadingHAKENITAKUUser));
        }
    }
    public bool IsLoadedEHRUser
    {
        get => EHR_USERs.Count > 0;
    }
    private async void LoadEHRUserAsync(string filePath)
    {
        // ReadEHRUserFromExcelAsync内の処理状況をUIに反映させるために、awaitを使用して非同期で実行
        // 読み込み中はUIがフリーズしないようにする
        // ProgressValueを更新する処理も追加する
        var progress = new Progress<int>(value => EhrProgressValue = value);
        IsLoadingEHRUser = true;
        var users = await Task.Run(() => ReadFile.ReadEHRUserFromExcelAsync(filePath, progress));
        IsLoadingEHRUser = false;
        if (users.IsSuccess)
        {
            EHR_USERs = new ObservableCollection<EHR_USER>(users.Value!);
        }
        else
        {
            // エラーメッセージを表示する処理を書く
            Debug.WriteLine($"EHR_USERの読み込みに失敗しました: {users.ErrorMessage}");
            ShowErrorMessage($"EHR_USERの読み込みに失敗しました:\n\r== 内容 ===\n\r {users.ErrorMessage}");
        }
    }
    private void ShowErrorMessage(string message)
    {
        var window = App.Current?.ApplicationLifetime is Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop
            ? desktop.MainWindow
            : null;
        var msgBox = new Avalonia.Controls.Window
        {
            Title = "エラー",
            Width = 400,
            Height = 200,
            Content = new Avalonia.Controls.TextBlock
            {
                Text = message,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                TextWrapping = Avalonia.Media.TextWrapping.Wrap
            }
        };
        msgBox.ShowDialog(window);
    }
    private async void LoadSYOKUINUserAsync(string filePath)
    {
        var progress = new Progress<int>(value => SyokuinProgressValue = value);
        var users = await Task.Run(() => ReadFile.ReadSYOKUINUserFromExcelAsync(filePath, progress));
        if (users.IsSuccess)
        {
            SYOKUIN_USERs = new ObservableCollection<SYOKUIN_USER>(users.Value!);
        }
        else
        {
            // エラーメッセージを表示する処理を書く
            Debug.WriteLine($"SYOKUIN_USERの読み込みに失敗しました: {users.ErrorMessage}");
            // MsgBox表示
            ShowErrorMessage($"SYOKUIN_USERの読み込みに失敗しました:\n\r== 内容 ===\n\r {users.ErrorMessage}");
        }

    }
    private async void LoadHAKENITAKUUserAsync(string filePath)
    {
        var progress = new Progress<int>(value => HakenitakuProgressValue = value);
        var users = await Task.Run(() => ReadFile.ReadHAKENITAKUUserFromExcelAsync(filePath, progress));
        if (!users.IsSuccess)
        {
            // エラーメッセージを表示する処理を書く
            Debug.WriteLine($"HAKENITAKU_USERの読み込みに失敗しました: {users.ErrorMessage}");
            ShowErrorMessage($"HAKENITAKU_USERの読み込みに失敗しました:\n\r== 内容 ===\n\r {users.ErrorMessage}");
            return;
        }
        else
        {
            HAKENITAKU_USERs = new ObservableCollection<HAKENITAKU_USER>(users.Value!);
        }
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
    public void TemplateDownloadCommand()
    {
        Debug.WriteLine("TemplateDownloadCommand 実行");
        // テンプレートダウンロード処理を書く
        // ユーザーに保存場所（ディレクトリ）を選択させる
        var folderPath = string.Empty;
        var window = App.Current.ApplicationLifetime is Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop
            ? desktop.MainWindow
            : null;
        var storage = window.StorageProvider;
        storage.OpenFolderPickerAsync(new Avalonia.Platform.Storage.FolderPickerOpenOptions
        {
            Title = "テンプレートの保存先フォルダを選択してください"
        }).ContinueWith(t =>
        {
            var folders = t.Result;
            if (folders != null && folders.Count > 0)
            {
                folderPath = folders[0].Path.LocalPath;
            }
        }).Wait();
        if (!string.IsNullOrEmpty(folderPath))
        {
            TemplateService.SaveTemplates(folderPath);
        }

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

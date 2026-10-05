using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

namespace SharePointExplorer.WinUI
{
    public sealed partial class ConnectionDialog : ContentDialog
    {
        private readonly Func<SqlConnectionOptions,Task<string>> connectAsync;
        private readonly Func<SqlConnectionOptions,List<string>> discover;
        private bool initialized,discovering,connecting,closed,closeAfterConnect;
        private int settingsGeneration;
        private string discoveryKey;
        public SqlConnectionOptions Options { get; private set; }

        public ConnectionDialog(SqlConnectionOptions defaults,Func<SqlConnectionOptions,Task<string>> connectAsync,
            Func<SqlConnectionOptions,List<string>> discover=null)
        {
            if(defaults==null) throw new ArgumentNullException(nameof(defaults));
            this.connectAsync=connectAsync ?? throw new ArgumentNullException(nameof(connectAsync));
            this.discover=discover ?? SqlDatabaseDiscovery.GetAccessibleDatabases;
            InitializeComponent();
            ServerBox.Text=defaults.Server ?? ""; DatabaseBox.Text=defaults.Database ?? "";
            AuthenticationBox.SelectedIndex=defaults.Authentication==SqlAuthenticationMode.SqlLogin ? 1 : 0;
            UsernameBox.Text=defaults.Username ?? ""; PasswordBox.Password="";
            EncryptBox.IsChecked=defaults.Encrypt; TrustCertificateBox.IsChecked=defaults.TrustServerCertificate;
            CurrentUserText.Text=Environment.UserDomainName+"\\"+Environment.UserName;
            initialized=true; UpdateAuthentication();
        }
        private SqlConnectionOptions CaptureOptions()
        {
            return new SqlConnectionOptions { Server=ServerBox.Text.Trim(), Database=DatabaseBox.Text.Trim(),
                Authentication=AuthenticationBox.SelectedIndex==1 ? SqlAuthenticationMode.SqlLogin : SqlAuthenticationMode.WindowsCurrentUser,
                Username=UsernameBox.Text.Trim(), Password=PasswordBox.Password,
                Encrypt=EncryptBox.IsChecked==true, TrustServerCertificate=TrustCertificateBox.IsChecked==true };
        }
        private void SetFieldsEnabled(bool enabled)
        {
            ServerBox.IsEnabled=enabled; DatabaseBox.IsEnabled=enabled; AuthenticationBox.IsEnabled=enabled;
            UsernameBox.IsEnabled=enabled; PasswordBox.IsEnabled=enabled; EncryptBox.IsEnabled=enabled;
            TrustCertificateBox.IsEnabled=enabled; RefreshDatabasesButton.IsEnabled=enabled && !discovering;
        }
        private void UpdateAuthentication()
        {
            bool sql=AuthenticationBox.SelectedIndex==1;
            CredentialsPanel.Visibility=sql ? Visibility.Visible : Visibility.Collapsed;
            CurrentUserText.Visibility=sql ? Visibility.Collapsed : Visibility.Visible;
        }
        private void SettingsChanged(object sender,RoutedEventArgs args)
        {
            if(!initialized) return;
            settingsGeneration++; discoveryKey=null; UpdateAuthentication();
            string current=DatabaseBox.Text; DatabaseBox.Items.Clear(); DatabaseBox.Text=current;
            DiscoveryText.Text="Connection settings changed. Refresh the database list or enter a name.";
            ConnectionError.IsOpen=false;
        }
        private bool Validate(SqlConnectionOptions options,bool database)
        {
            if(String.IsNullOrWhiteSpace(options.Server)) { ShowError("Enter the SQL Server name or instance."); ServerBox.Focus(FocusState.Programmatic); return false; }
            if(database && String.IsNullOrWhiteSpace(options.Database)) { ShowError("Enter the restored content database name."); DatabaseBox.Focus(FocusState.Programmatic); return false; }
            if(options.Authentication==SqlAuthenticationMode.SqlLogin && String.IsNullOrWhiteSpace(options.Username)) { ShowError("Enter the SQL login username."); UsernameBox.Focus(FocusState.Programmatic); return false; }
            return true;
        }
        private async void ConnectClicked(ContentDialog sender,ContentDialogButtonClickEventArgs args)
        {
            args.Cancel=true;
            SqlConnectionOptions options=CaptureOptions(); if(!Validate(options,true)) return;
            ContentDialogButtonClickDeferral deferral=args.GetDeferral(); connecting=true;
            IsPrimaryButtonEnabled=false; SetFieldsEnabled(false); ConnectingPanel.Visibility=Visibility.Visible; ConnectionError.IsOpen=false;
            try
            {
                string error=await connectAsync(options);
                if(String.IsNullOrWhiteSpace(error)) { Options=options.Clone(); Options.Password=""; args.Cancel=false; }
                else ShowError(Redact(error,options.Password));
            }
            catch(Exception error) { ShowError(Redact(error.GetBaseException().Message,options.Password)); }
            finally
            {
                options.Password=""; connecting=false; IsPrimaryButtonEnabled=true; SetFieldsEnabled(true);
                ConnectingPanel.Visibility=Visibility.Collapsed; deferral.Complete();
                if(closeAfterConnect && !closed) DispatcherQueue.TryEnqueue(()=>{ if(!closed) Hide(); });
            }
        }
        private async void DialogOpened(ContentDialog sender,ContentDialogOpenedEventArgs args)
        {
            if(String.IsNullOrWhiteSpace(ServerBox.Text))
            {
                ServerBox.Focus(FocusState.Programmatic);
                return;
            }
            await DiscoverDatabases(false);
        }
        private async void RefreshDatabasesClicked(object sender,RoutedEventArgs args) { await DiscoverDatabases(true); }
        private async void DatabaseDropDownOpened(object sender,object args) { await DiscoverDatabases(false); }
        private async Task DiscoverDatabases(bool force)
        {
            if(discovering || connecting || closed) return;
            SqlConnectionOptions options=CaptureOptions(); if(!Validate(options,false)) return;
            string key=options.Server+"\0"+options.Authentication+"\0"+options.Username+"\0"+options.Encrypt+"\0"+options.TrustServerCertificate;
            if(!force && discoveryKey==key) return;
            int request=settingsGeneration; discovering=true; RefreshDatabasesButton.IsEnabled=false; DiscoveryText.Text="Loading accessible databases...";
            try
            {
                List<string> names=await Task.Run(()=>discover(options));
                if(closed || connecting) return;
                if(request!=settingsGeneration) { DiscoveryText.Text="Connection settings changed. Refresh the database list."; return; }
                string current=DatabaseBox.Text; DatabaseBox.Items.Clear(); foreach(string name in names) DatabaseBox.Items.Add(name);
                DatabaseBox.Text=current; discoveryKey=key; DiscoveryText.Text=names.Count+" databases available. Choose one or enter a name.";
            }
            catch(Exception error)
            {
                if(!closed && !connecting && request==settingsGeneration) DiscoveryText.Text="You can enter the database name manually. "+Redact(error.GetBaseException().Message,options.Password);
            }
            finally { options.Password=""; discovering=false; if(!closed) RefreshDatabasesButton.IsEnabled=true; }
        }
        public void RequestClose() { if(closed) return; if(connecting) closeAfterConnect=true; else Hide(); }
        private void DialogClosing(ContentDialog sender,ContentDialogClosingEventArgs args) { if(connecting) args.Cancel=true; }
        private void DialogClosed(ContentDialog sender,ContentDialogClosedEventArgs args) { closed=true; settingsGeneration++; PasswordBox.Password=""; }
        private void ShowError(string text) { ConnectionError.Message=Safe(text); ConnectionError.IsOpen=true; }
        private static string Redact(string text,string password) { return Safe(String.IsNullOrEmpty(password) ? text : (text ?? "").Replace(password,"[hidden]")); }
        private static string Safe(string text) { if(text==null) return ""; char[] values=text.ToCharArray(); for(int index=0;index<values.Length;index++) if(Char.IsControl(values[index])) values[index]=' '; return new string(values); }
    }
}




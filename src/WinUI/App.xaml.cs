using System;
using System.Linq;
using Microsoft.UI.Xaml;

namespace SharePointExplorer.WinUI
{
    public partial class App : Application
    {
        private MainWindow window;
        public App() { InitializeComponent(); }
        protected override async void OnLaunched(LaunchActivatedEventArgs args)
        {
            string[] commandLine=Environment.GetCommandLineArgs().Skip(1).ToArray();
            bool diagnostics=commandLine.Any(value=>value=="--ui-check" || value=="--ui-integration");
            window=new MainWindow(null,diagnostics);
            if(diagnostics) window.AppWindow.Move(new Windows.Graphics.PointInt32(-20000,-20000));
            window.Activate();
            if(diagnostics)
            {
                await window.StartupTask;
                await WinUiSmokeChecks.RunAsync(window,commandLine);
            }
        }
    }
}

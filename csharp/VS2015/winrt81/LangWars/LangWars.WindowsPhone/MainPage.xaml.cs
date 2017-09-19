using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Storage;
using Windows.UI.ViewManagement;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

namespace LangWars
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        bool Offline = false;

        public MainPage()
        {
            this.InitializeComponent();

            this.NavigationCacheMode = NavigationCacheMode.Required;
            ResultsWindow.NavigationCompleted += ResultsWindow_NavigationCompleted;
            FileShare.Register();
        }

        private void ResultsWindow_NavigationCompleted(WebView sender, WebViewNavigationCompletedEventArgs args)
        {
            if (!args.IsSuccess)
            {
                SetHTML("<html>" + WebUtility.HtmlEncode("Press the Fight button!") + "</html>");
            }
        }

        private async void StartProgress()
        {
            var Progress = StatusBar.GetForCurrentView().ProgressIndicator;
            Progress.ProgressValue = null;
            await Progress.ShowAsync();
        }

        private async void EndProgress()
        {
            var Progress = StatusBar.GetForCurrentView().ProgressIndicator;
            await Progress.HideAsync();
        }

        private void PhoneApplicationPage_Loaded(object sender, RoutedEventArgs e)
        {
            //OpenLastReport();
        }

        private void OpenLastReport()
        {
            try
            {
                LoadHTML();
            }
            catch (Exception ex)
            {
                SetHTML("<html>" + WebUtility.HtmlEncode(ex.Message) + "</html>");
            }
        }

        private async void appBarFightButton_Click(object sender, RoutedEventArgs e)
        {
            appBar.IsEnabled = false;
            try
            {
                await new ReportGenerator().TryCreateReport(Offline, StartProgress, EndProgress, LoadHTML, SetHTML);
            }
            finally
            {
                appBar.IsEnabled = true;
            }
        }

        void LoadHTML()
        {
            ResultsWindow.Navigate(new System.Uri("ms-appdata:///local/" + ReportGenerator.TempHtmlRelFolder + "/" + ReportGenerator.TempHtmlName));
        }

        void SetHTML(string html)
        {
            ResultsWindow.NavigateToString(html);
        }

        private void ShareButton_Click(object sender, RoutedEventArgs e)
        {
            FileShare.Share();
        }

    }
}

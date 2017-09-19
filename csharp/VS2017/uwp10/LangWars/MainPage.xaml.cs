using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using Windows.Foundation;
using Windows.Foundation.Collections;
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
        bool Offline => !OnlineButton.IsChecked.GetValueOrDefault(false);

        public MainPage()
        {
            this.InitializeComponent();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); //Make all encodings available. Not really necessary as the app will work anyway with other encodings.
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

        private void StartProgress()
        {
            // see http://code.msdn.microsoft.com/windowsapps/How-to-put-a-ProgressRing-a92f2530
            Progress.IsActive = true;
            Progress.Visibility = Windows.UI.Xaml.Visibility.Visible;
            WebViewBrush brush = new WebViewBrush();
            brush.SourceName = "ResultsWindow";
            brush.Redraw();
            MaskRectangle.Fill = brush;
            MaskRectangle.Visibility = Windows.UI.Xaml.Visibility.Visible;
            ResultsWindow.Visibility = Windows.UI.Xaml.Visibility.Collapsed;
        }

        private void EndProgress()
        {
            Progress.IsActive = false;
            Progress.Visibility = Windows.UI.Xaml.Visibility.Collapsed;
            ResultsWindow.Visibility = Windows.UI.Xaml.Visibility.Visible;
            MaskRectangle.Visibility = Windows.UI.Xaml.Visibility.Collapsed;
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

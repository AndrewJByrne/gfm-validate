using Octokit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace GfmValidate
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string _previewTemplate = string.Empty;

        public MainWindow()
        {
            InitializeComponent();

            ContentPreview.Navigated += ContentPreview_Navigated;
            SetSilent(ContentPreview, true);

            // register for both handled and unhandled Executed events

            MarkDownText.AddHandler(CommandManager.ExecutedEvent,

                new RoutedEventHandler(CommandExecuted), true);

            MarkDownText.PreviewDragOver += MarkDownText_PreviewDragOver;

            Loaded += MainWindow_Loaded;
        }

        void MarkDownText_PreviewDragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;

            e.Handled = true;
            
        }

        void ContentPreview_Navigated(object sender, NavigationEventArgs e)
        {
            SetSilent(ContentPreview, true);
        }

        async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            using (StreamReader reader = new StreamReader("Preview.html"))
            {
                _previewTemplate = await reader.ReadToEndAsync();
            }
        }

        private async void CommandExecuted(object sender, RoutedEventArgs e)
        {

            if ((e as ExecutedRoutedEventArgs).Command

                == ApplicationCommands.Paste)
            {

                // verify that the textbox handled the paste command

                if (e.Handled)
                {
                    string html = string.Empty;
                    if (String.IsNullOrEmpty(MarkDownText.Text))
                    {
                        await PreviewMarkDownAsync();
                    }
                }

            }

        }

        private GitHubClient _gitHubClient = null;

        private async Task PreviewMarkDownAsync()
        {
            if (_gitHubClient == null)
                _gitHubClient = new GitHubClient(new ProductHeaderValue("AndrewJByrne", "pepsi101"));

            string html = await _gitHubClient.Miscellaneous.RenderRawMarkdown(MarkDownText.Text);

            // Could trigger a validation of the text right now!
           // MessageBox.Show(html);
            ContentPreview.NavigateToString(string.Format(_previewTemplate, html));
        }

        public static void SetSilent(WebBrowser browser, bool silent)
        {
            if (browser == null)
                throw new ArgumentNullException("browser");

            // get an IWebBrowser2 from the document
            IOleServiceProvider sp = browser.Document as IOleServiceProvider;
            if (sp != null)
            {
                Guid IID_IWebBrowserApp = new Guid("0002DF05-0000-0000-C000-000000000046");
                Guid IID_IWebBrowser2 = new Guid("D30C1661-CDAF-11d0-8A3E-00C04FC9E26E");

                object webBrowser;
                sp.QueryService(ref IID_IWebBrowserApp, ref IID_IWebBrowser2, out webBrowser);
                if (webBrowser != null)
                {
                    webBrowser.GetType().InvokeMember("Silent", BindingFlags.Instance | BindingFlags.Public | BindingFlags.PutDispProperty, null, webBrowser, new object[] { silent });
                    webBrowser.GetType().InvokeMember("RegisterAsDropTarget", BindingFlags.Instance | BindingFlags.Public | BindingFlags.PutDispProperty, null, webBrowser, new object[] { false });
                }
            }
        }


        [ComImport, Guid("6D5140C1-7436-11CE-8034-00AA006009FA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleServiceProvider
        {
            [PreserveSig]
            int QueryService([In] ref Guid guidService, [In] ref Guid riid, [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);
        }

        private void MarkDownText_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;

        }

        private async void MarkDownText_Drop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Length != 0)
            {
                var filePath = files[0];
                using (StreamReader reader = new StreamReader(filePath))
                {
                    MarkDownText.Text = await reader.ReadToEndAsync();
                    await PreviewMarkDownAsync();
                }
            }
        }

        private void MarkDownText_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;
        }

    }
}

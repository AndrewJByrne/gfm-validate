using Microsoft.Win32;
using Octokit;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
        public MainViewModel MainViewModel = new MainViewModel();

        public MainWindow()
        {
            InitializeComponent();

            ContentPreview.Navigated += ContentPreview_Navigated;
            SetSilent(ContentPreview, true);

            // I want to preview whatever the user pastes into the markdown textbox
            // To do this, I listen for the Paste event.
            MarkDownText.AddHandler(CommandManager.ExecutedEvent,

                new RoutedEventHandler(CommandExecuted), true);

            MarkDownText.PreviewDragOver += MarkDownText_PreviewDragOver;

            Loaded += MainWindow_Loaded;
            Closing += MainWindow_Closing;

            this.DataContext = MainViewModel;
        }

        void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Properties.Settings.Default.Save();
        }

        // For some reason, DragOver is never fired, so I do the work here instead. 
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
            GunText.Text = MainViewModel.Gun;
            PunText.Password = MainViewModel.Pun;
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
            try
            {
                if (_gitHubClient == null)
                    _gitHubClient = new GitHubClient(new ProductHeaderValue(MainViewModel.Gun, MainViewModel.Pun));

                string html = await _gitHubClient.Miscellaneous.RenderRawMarkdown(MarkDownText.Text);

                // Could trigger a validation of the text right now!
                // MessageBox.Show(html);
                ContentPreview.NavigateToString(string.Format(_previewTemplate, html));
            }
            catch(ArgumentException aex)
            {
                MessageBox.Show(aex.Message);
            }
        }

        // I do two things here - I suppress warnings and I make sure the browser is not a drop target.
        // This is done using reflection, setting properties through IWebBrowser2
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

        // NEVER FIRES - can remove
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
            if (files != null && files.Length == 1)
            {
                var filePath = files[0];
                await LoadFileAsync(filePath);
                await PreviewMarkDownAsync();
            }
        }

        private async Task LoadFileAsync(string path)
        {
            using (StreamReader reader = new StreamReader(path))
            {
                MarkDownText.Text =  await reader.ReadToEndAsync();
            }

        }

        private void MarkDownText_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;
        }

        private void SetCredentials_Click(object sender, RoutedEventArgs e)
        {
            // Password on PasswordBox is not bindabe since it isn't a DependencyObject.
            // This is done for security reasons, since a string-based version of the password would be stored in the DO tree
            // However, I'm not concerned about that for this internal tool right now, but I still have to update the field manually.
            MainViewModel.Gun = GunText.Text;
            MainViewModel.Pun = PunText.Password;
        }

        private async void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false;
            ofd.Filter = "markdown files (*.md)|*.md|txt files (*.txt)|*.txt|All files (*.*)|*.*";
            ofd.Title = "Select a MarkDown file";
            bool? response = ofd.ShowDialog();
            if (response.HasValue)
            {
                Debug.WriteLine(ofd.FileName);
                await LoadFileAsync(ofd.FileName);
                await PreviewMarkDownAsync();
            }
        }

        private async void Validate_Click(object sender, RoutedEventArgs e)
        {
            await PreviewMarkDownAsync();
        }

    }
}

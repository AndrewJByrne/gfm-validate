using Microsoft.Win32;
using Octokit;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Navigation;

namespace Andrew.J.Byrne.GfmValidate
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // String to hold the preview (HTML) template. 
        // I inject the rendered markdwon into this template. 
        private string _previewTemplate = string.Empty;

        private GitHubClient _gitHubClient = null;
        public MainViewModel _vM = new MainViewModel();

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

            this.DataContext = _vM;
        }

        void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Properties.Settings.Default.Save();
        }

        // For some reason, DragOver is never fired, so I do the work here instead. 
        void MarkDownText_PreviewDragOver(object sender, DragEventArgs e)
        {
            // Check if the user is trying to drop a file
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                // The copy cursor indicates to the user that this action is allowed. 
                e.Effects = DragDropEffects.Copy;
            else
                // Indicate that whatever the user is trying to drop is not allowed. 
                e.Effects = DragDropEffects.None;

            e.Handled = true;
            
        }

        void ContentPreview_Navigated(object sender, NavigationEventArgs e)
        {
            SetSilent(ContentPreview, true);
        }

        async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // I should do this with data binding. However, the Password field on 
            // the PasswordBox is not bindable for security reasons. 
            GitHubUsernameTextBox.Text = _vM.GitHubUsername;
            GitHubPasswordBox.Password = _vM.GitHubPassword;

            // Load the preview template
            using (StreamReader reader = new StreamReader("Preview.html"))
            {
                _previewTemplate = await reader.ReadToEndAsync();
            }
        }

        private bool CheckCredentials()
        {
            if (String.IsNullOrEmpty(_vM.GitHubUsername) || String.IsNullOrEmpty(_vM.GitHubPassword))
            {
                MessageBox.Show("You need to enter GitHub credentials to use this app.");
                return false;
            }
            return true;
        }

        private async void CommandExecuted(object sender, RoutedEventArgs e)
        {
            if ((e as ExecutedRoutedEventArgs).Command == ApplicationCommands.Paste)
            {
                // verify that the textbox handled the paste command
                if (e.Handled)
                {
                    if (String.IsNullOrEmpty(MarkDownText.Text))
                    {
                        await PreviewMarkDownAsync();
                    }
                }

            }

        }

        // Use the Octokit client library to render the markdown into HTML
        private async Task PreviewMarkDownAsync()
        {
            try
            {
                if (CheckCredentials())
                {
                    string html = "";

                    // Lasy-load the client. I could refactor this to a lazy-loaded property
                    if (_gitHubClient == null)
                    {
                        var productHeaderValue = new ProductHeaderValue(_vM.GitHubUsername, _vM.GitHubPassword);
                        _gitHubClient = new GitHubClient(productHeaderValue);
                        html = await _gitHubClient.Miscellaneous.RenderRawMarkdown(MarkDownText.Text);
                    }

                    // Could trigger a validation of the text right now!
                    // MessageBox.Show(html);
                    ContentPreview.NavigateToString(string.Format(_previewTemplate, html));
                }
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
                    webBrowser.GetType().InvokeMember("Silent"
                                                        , BindingFlags.Instance 
                                                        | BindingFlags.Public 
                                                        | BindingFlags.PutDispProperty
                                                        , null, webBrowser, new object[] { silent });

                    webBrowser.GetType().InvokeMember("RegisterAsDropTarget"
                                                      , BindingFlags.Instance 
                                                      | BindingFlags.Public 
                                                      | BindingFlags.PutDispProperty
                                                      , null, webBrowser, new object[] { false });
                }
            }
        }


        [ComImport, Guid("6D5140C1-7436-11CE-8034-00AA006009FA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleServiceProvider
        {
            [PreserveSig]
            int QueryService([In] ref Guid guidService, [In] ref Guid riid, [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);
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
            // Password on PasswordBox is not bindable since it isn't a DependencyObject.
            // This is done for security reasons, since a string-based version of the password would be stored in the DO tree
            // However, I'm not concerned about that for this internal tool right now, but I still have to update the field manually.
            _vM.GitHubUsername = GitHubUsernameTextBox.Text;
            _vM.GitHubPassword = GitHubPasswordBox.Password;
            
        }

        private async void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false;
            ofd.Filter = "markdown files (*.md)|*.md|txt files (*.txt)|*.txt|All files (*.*)|*.*";
            ofd.Title = "Select a MarkDown file";
            bool? response = ofd.ShowDialog();
            if (response.HasValue && !String.IsNullOrEmpty(ofd.FileName))
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

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        private async void NewFile_Click(object sender, RoutedEventArgs e)
        {
            MarkDownText.Text = string.Empty;
            await PreviewMarkDownAsync();
        }

    }
}

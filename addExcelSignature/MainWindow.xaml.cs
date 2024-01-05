using Microsoft.Win32;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.IO;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.DirectoryServices.AccountManagement;
using System.Windows.Input;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace addExcelSignature
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string SourceFilePathName = string.Empty;
        string DestinationFilePathName = string.Empty;
        public static App CurrentApp => (App)Application.Current;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            InitializeWindow();
        }

        private void Window_ContentRendered(object sender, EventArgs e)
        {
            Process();
        }

        private async void InitializeWindow()
        {

            // Command line argument is present
            if (CurrentApp.SourceFilePathName != string.Empty)
            {
                this.SourceFilePathName = CurrentApp.SourceFilePathName;
            }
            TextBoxSourceFilePathName.Text = this.SourceFilePathName;

            TextBoxDefaultSignaturePath.Text = Properties.Settings.Default.DefaultSignatureSavePath;

            this.DestinationFilePathName = await this.GetDestinationFile();
            TextBoxDestinationFilePathName.Text = this.DestinationFilePathName;

            if (this.SourceFilePathName == string.Empty || this.DestinationFilePathName == string.Empty)
            {
                StatusText.Content = "Selecting a file to sign is required!";
                return;
            }

        }

        private async void ChangeSourceFile(object sender, RoutedEventArgs e)
        {
            // Command line argument is not present
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (.xlsx)|*.xlsx";

            // Show open file dialog box
            bool? result = openFileDialog.ShowDialog();
            if (result == true)
            {
                this.SourceFilePathName = openFileDialog.FileName;
                TextBoxSourceFilePathName.Text = this.SourceFilePathName;

                this.DestinationFilePathName = await this.GetDestinationFile();
                TextBoxDestinationFilePathName.Text = this.DestinationFilePathName;


                StatusText.Content = "Ready";
            }
        }

        private async void ChangeDefaultSignaturePath(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = Properties.Settings.Default.DefaultSignatureSavePath;
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                Properties.Settings.Default.DefaultSignatureSavePath = dialog.FileName;
                Properties.Settings.Default.Save();
            }
        }

        private async void Process(object sender, RoutedEventArgs e)
        {
            Process();
        }

        private async void Process()
        {
            if(this.SourceFilePathName==string.Empty || this.DestinationFilePathName==string.Empty) {
                return;
            }

            try
            {
                StatusText.Content = "Copying file";
                this.CopySourceFileToDestination();
                StatusText.Content = "Generating signature";
                this.AddSignatureToDestinationFile();
                StatusText.Content = "Creating signature document succeeded";
            }
            catch (Exception e)
            {
                StatusText.Content = "Creating signature document failed: " + e.Message;
                this.DeleteDestinationFileOnFail();
            }
        }

        private void CopySourceFileToDestination()
        {
            File.Copy(this.SourceFilePathName, this.DestinationFilePathName, true);
        }

        private void DeleteDestinationFileOnFail()
        {
            try
            {
                File.Delete(this.DestinationFilePathName);
            }
            catch (Exception e)
            {
                StatusText.Content = e.Message;
            }
        }

        private async Task<string> GetDestinationFile()        {
            if(this.SourceFilePathName==string.Empty)
            {
                return "";
            }

            string sourceName = Path.GetFileNameWithoutExtension(this.SourceFilePathName);
            string sourcePath = Path.GetFullPath(this.SourceFilePathName).Replace(Path.GetFileName(this.SourceFilePathName), "");
            string now = DateTime.Now.ToString("yyyy-MM-dd");
            string destPath = Properties.Settings.Default.DefaultSignatureSavePath == string.Empty ? sourcePath : Properties.Settings.Default.DefaultSignatureSavePath;
            return (destPath + "\\" + sourceName + "-signature-" + now + ".xlsx").Replace("\\\\", "\\");
        }

        private void AddSignatureToDestinationFile()
        {
            string CurrentUserName = string.Empty;
            string CurrentUserEmail = string.Empty;
            try
            {
                CurrentUserName = UserPrincipal.Current.GivenName + " " + UserPrincipal.Current.Surname;
                CurrentUserEmail = UserPrincipal.Current.EmailAddress;
            }
            catch (InvalidOperationException e)
            {

            }

            Excel.Application excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            Excel.Workbook workbook = excelApp.Workbooks.Open(this.DestinationFilePathName);
            Excel.Worksheet worksheet = workbook.Sheets.Add();
            worksheet.Name = "Approval Signature";
            SignatureSet signatureSet = workbook.Signatures;
            Signature signature = signatureSet.AddSignatureLine();
            SignatureSetup signatureSetup = signature.Setup;
            signatureSetup.ShowSignDate = true;
            signatureSetup.SuggestedSigner = CurrentUserName;
            signatureSetup.SigningInstructions = "I agree that I have reviewed this file, made notes as appropriate, and addressed all issues presented.";
            signatureSetup.SuggestedSignerEmail = CurrentUserEmail;
            workbook.Save();
            excelApp.Visible = true;
            /*workbook.Close();
            excelApp.Quit();*/
        }

        private async void ClearDestinationPath(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.DefaultSignatureSavePath = string.Empty;
            Properties.Settings.Default.Save();

            this.DestinationFilePathName = await this.GetDestinationFile();
            TextBoxDestinationFilePathName.Text = this.DestinationFilePathName;
        }

    }
}

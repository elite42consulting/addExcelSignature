using System.IO;
using System.Windows;

namespace addExcelSignature
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public string SourceFilePathName = string.Empty;
        protected override void OnStartup(StartupEventArgs e)
        {
            foreach (string arg in e.Args)
            {
                if (File.Exists(arg))
                {
                    this.SourceFilePathName = arg;
                }
            }
            base.OnStartup(e);
        }
    }
}

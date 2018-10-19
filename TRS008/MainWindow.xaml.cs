using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.IO;

namespace TRS008
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void TargetSearch(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            var result = openFileDialog.ShowDialog();
            if(result == true)
            {
                this.targetPath.Text = openFileDialog.FileName;
            }
        }

        private void ToDo(object sender, RoutedEventArgs e)
        {
            var targetPathValue = targetPath.Text;
            var begin = dateBegin.Text;
            var end = dateEnd.Text;

            object oMissing = System.Reflection.Missing.Value;
             
        }
    }
}

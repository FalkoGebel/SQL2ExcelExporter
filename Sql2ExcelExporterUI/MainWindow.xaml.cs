using ExporterLogicLibrary;
using System.Windows;

namespace Sql2ExcelExporterUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void DatabaseAssistButton_Click(object sender, RoutedEventArgs e)
        {
            OpenDatabaseChoiceWindow();
        }

        private void OpenDatabaseChoiceWindow()
        {
            try
            {
                List<string> databases = SqlLogic.GetDatabasesFromServer(ServerTextBox.Text);
                DatabaseChoiceWindow dbcw = new(databases);
                dbcw.ShowDialog();
                DatabaseTextBox.Text = dbcw.GetDatabase();
            }
            catch (Exception e)
            {
                ShowError(e.Message);
            }
        }

        private static void ShowError(string msg)
        {
            MessageBox.Show(msg, Properties.Resources.ERROR_TITLE, MessageBoxButton.OK, MessageBoxImage.Error);
        }

    }
}
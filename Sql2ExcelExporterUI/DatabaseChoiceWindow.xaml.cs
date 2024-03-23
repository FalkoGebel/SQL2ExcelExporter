using System.Windows;

namespace Sql2ExcelExporterUI
{
    /// <summary>
    /// Interaktionslogik für DatabaseChoiceWindow.xaml
    /// </summary>
    public partial class DatabaseChoiceWindow : Window
    {
        private string _database = "";
        private readonly List<string> _databases = [];

        public DatabaseChoiceWindow(List<string> databases)
        {
            InitializeComponent();
            _databases = databases;
            DatabasesListView.ItemsSource = _databases;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            SetDatabaseAndCloseWindow();
        }

        private void SetDatabaseAndCloseWindow()
        {
            _database = (string)DatabasesListView.SelectedValue;
            Close();
        }

        public string GetDatabase()
        {
            return _database;
        }
    }
}

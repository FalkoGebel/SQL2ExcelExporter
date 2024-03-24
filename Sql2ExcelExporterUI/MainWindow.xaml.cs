using ExporterLogicLibrary;
using ExporterLogicLibrary.Models;
using Sql2ExcelExporterUI.Models;
using System.Windows;

namespace Sql2ExcelExporterUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<ColumnsListViewModel> _columns = [];

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
                ChoiceWindow dbcw = new(Properties.Resources.DBCW_TITLE, databases);
                dbcw.ShowDialog();
                DatabaseTextBox.Text = dbcw.GetChoice();
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

        private void UpdateColumnsListView(bool fromDatabase)
        {
            if (fromDatabase)
            {
                _columns = [];

                if (ServerTextBox.Text == string.Empty || DatabaseTextBox.Text == string.Empty)
                    return;

                foreach (ColumnModel col in SqlLogic.GetColumnsForTable(ServerTextBox.Text, DatabaseTextBox.Text, TableTextBox.Text).OrderBy(cm => cm.Name))
                    _columns.Add(new ColumnsListViewModel() { Name = col.Name, Selected = true, Type = col.Type });
            }

            ColumnsListView.ItemsSource = null;
            ColumnsListView.ItemsSource = _columns;
        }

        private void TableAssistButton_Click(object sender, RoutedEventArgs e)
        {
            OpenTableChoiceWindow();
        }

        private void OpenTableChoiceWindow()
        {
            try
            {
                List<string> tables = SqlLogic.GetTablesForDatabase(ServerTextBox.Text, DatabaseTextBox.Text);
                ChoiceWindow dbcw = new(Properties.Resources.TCW_TITLE, tables);
                dbcw.ShowDialog();
                TableTextBox.Text = dbcw.GetChoice();
                UpdateColumnsListView(true);
            }
            catch (Exception e)
            {
                ShowError(e.Message);
            }
        }

        private void SelectAllColumnsButton_Click(object sender, RoutedEventArgs e)
        {
            SelectAllListViewColumns();
        }

        private void SelectAllListViewColumns()
        {
            foreach (var column in _columns)
                column.Selected = true;

            UpdateColumnsListView(false);
        }

        private void SelectNoneColumnsButton_Click(object sender, RoutedEventArgs e)
        {
            SelectNoneListViewColumns();
        }

        private void SelectNoneListViewColumns()
        {
            foreach (var column in _columns)
                column.Selected = false;

            UpdateColumnsListView(false);
        }
    }
}
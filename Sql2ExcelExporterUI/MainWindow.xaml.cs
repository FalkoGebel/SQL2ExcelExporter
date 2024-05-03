using DocumentFormat.OpenXml.Packaging;
using ExporterLogicLibrary;
using ExporterLogicLibrary.Models;
using Sql2ExcelExporterUI.Models;
using System.Windows;
using System.Windows.Forms;

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
                TableTextBox.Text = string.Empty;
            }
            catch (Exception e)
            {
                ShowError(e.Message);
            }
        }

        private static void ShowError(string msg)
        {
            System.Windows.MessageBox.Show(msg, Properties.Resources.ERROR_TITLE, MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private static void ShowInformation(string msg)
        {
            System.Windows.MessageBox.Show(msg, Properties.Resources.MW_INFO_TITLE, MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void UpdateColumnsListView(bool fromDatabase)
        {
            if (fromDatabase)
            {
                _columns = [];

                if (ServerTextBox.Text != string.Empty && DatabaseTextBox.Text != string.Empty && TableTextBox.Text != string.Empty)
                {
                    foreach (ColumnModel col in SqlLogic.GetColumnsForTable(ServerTextBox.Text, DatabaseTextBox.Text, TableTextBox.Text).OrderBy(cm => cm.Name))
                        _columns.Add(new ColumnsListViewModel() { Supported = col.Type.FormatCode() != string.Empty, Selected = col.Type.FormatCode() != string.Empty, Name = col.Name, Type = col.Type });
                }
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
                ChoiceWindow dbcw = new(Properties.Resources.TCW_TITLE, [.. tables.OrderBy(t => t)]);
                bool? result = dbcw.ShowDialog();
                if (result == null || !(bool)result)
                    return;
                TableTextBox.Text = dbcw.GetChoice();
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
                column.Selected = column.Supported;

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

        private void DirectoryAssistButton_Click(object sender, RoutedEventArgs e)
        {
            ChooseDirectory();
        }

        private void ChooseDirectory()
        {
            using FolderBrowserDialog fbd = new();
            DialogResult result = fbd.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                DirectoryTextBox.Text = fbd.SelectedPath;
        }

        private void CreateExcelFileButton_Click(object sender, RoutedEventArgs e)
        {
            CreateExcelFile();
        }

        private void CreateExcelFile()
        {
            if (TableTextBox.Text == string.Empty)
            {
                ShowError(Properties.Resources.MW_ERROR_MISSING_TABLE);
                return;
            }

            if (DirectoryTextBox.Text == string.Empty)
            {
                ShowError(Properties.Resources.MW_ERROR_MISSING_DIRECTORY);
                return;
            }

            List<ColumnModel> selectedColumns = _columns.Where(col => col.Selected).Select(clvw => new ColumnModel() { Name = clvw.Name, Type = clvw.Type }).ToList();
            if (selectedColumns.Count == 0)
            {
                ShowError(Properties.Resources.MW_ERROR_NO_COLUMNS_SELECTED);
                return;
            }

            // Get the data for the selected columns
            List<List<CellModel>> dataLines = SqlLogic.GetContentForTable(ServerTextBox.Text, DatabaseTextBox.Text, TableTextBox.Text,
                selectedColumns);

            string filePath = $"{DirectoryTextBox.Text}\\{TableTextBox.Text}.xlsx";

            SpreadsheetDocument s = ExcelLogic.CreateSpreadsheetDocument(filePath, TableTextBox.Text);
            ExcelLogic.InsertHeaderLine(s, TableTextBox.Text, selectedColumns.Select(cm => cm.Name).ToList());
            ExcelLogic.InsertDataLines(s, TableTextBox.Text, dataLines);
            s.SaveAndClose();

            ShowInformation(Properties.Resources.MW_INFO_FILE_CREATED.Replace("{FILE_PATH}", filePath));
        }

        private void TableTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdateColumnsListView(true);
        }
    }
}
using DocumentFormat.OpenXml.Packaging;
using ExporterLogicLibrary;
using ExporterLogicLibrary.Models;
using Sql2ExcelExporterUI.Models;
using System.Drawing.Text;
using System.Windows;
using System.Windows.Forms;

namespace Sql2ExcelExporterUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly string _defaultFontName = "Arial";
        private readonly int _defaultHeaderFontSize = 12;
        private readonly int _defaultDataFontSize = 10;
        private readonly System.Drawing.Color _defaultFontColor = System.Drawing.Color.Black;
        private readonly System.Drawing.Color _defaultFillColor = System.Drawing.Color.White;
        private readonly System.Drawing.Color _defaultBorderColor = System.Drawing.Color.White;
        private List<ColumnsListViewModel> _columns = [];

        public MainWindow()
        {
            InitializeComponent();
            InitHeaderStyleFontSizeTextBox();
            InitHeaderStyleFontColorPickers();
            InitHeaderStyleFontNameComboBoxes();
            InitDataStyleFontSizeTextBox();
        }

        private void InitHeaderStyleFontNameComboBoxes()
        {
            HeaderStyleFontNameComboBox.Items.Clear();
            DataStyleFontNameComboBox.Items.Clear();
            using InstalledFontCollection col = new();
            foreach (System.Drawing.FontFamily fa in col.Families)
            {
                HeaderStyleFontNameComboBox.Items.Add(fa.Name);
                DataStyleFontNameComboBox.Items.Add(fa.Name);
            }

            if (HeaderStyleFontNameComboBox.Items.Contains(_defaultFontName))
                HeaderStyleFontNameComboBox.Text = _defaultFontName;
            else
                HeaderStyleFontNameComboBox.SelectedIndex = 0;

            if (DataStyleFontNameComboBox.Items.Contains(_defaultFontName))
                DataStyleFontNameComboBox.Text = _defaultFontName;
            else
                DataStyleFontNameComboBox.SelectedIndex = 0;
        }

        private void InitHeaderStyleFontColorPickers()
        {
            HeaderStyleFontColorPicker.SelectedColor = new System.Windows.Media.Color()
            {
                A = _defaultFontColor.A,
                R = _defaultFontColor.R,
                G = _defaultFontColor.G,
                B = _defaultFontColor.B
            };

            HeaderStyleFillColorPicker.SelectedColor = new System.Windows.Media.Color()
            {
                A = _defaultFillColor.A,
                R = _defaultFillColor.R,
                G = _defaultFillColor.G,
                B = _defaultFillColor.B
            };

            HeaderStyleBorderColorPicker.SelectedColor = new System.Windows.Media.Color()
            {
                A = _defaultBorderColor.A,
                R = _defaultBorderColor.R,
                G = _defaultBorderColor.G,
                B = _defaultBorderColor.B
            };

            DataStyleFontColorPicker.SelectedColor = new System.Windows.Media.Color()
            {
                A = _defaultFontColor.A,
                R = _defaultFontColor.R,
                G = _defaultFontColor.G,
                B = _defaultFontColor.B
            };

            DataStyleFillColorPicker.SelectedColor = new System.Windows.Media.Color()
            {
                A = _defaultFillColor.A,
                R = _defaultFillColor.R,
                G = _defaultFillColor.G,
                B = _defaultFillColor.B
            };

            DataStyleBorderColorPicker.SelectedColor = new System.Windows.Media.Color()
            {
                A = _defaultBorderColor.A,
                R = _defaultBorderColor.R,
                G = _defaultBorderColor.G,
                B = _defaultBorderColor.B
            };
        }

        private void InitHeaderStyleFontSizeTextBox()
        {
            HeaderStyleFontSizeTextBox.Text = _defaultHeaderFontSize.ToString();
        }

        private void InitDataStyleFontSizeTextBox()
        {
            DataStyleFontSizeTextBox.Text = _defaultDataFontSize.ToString();
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

            CellFormatDefinition cellFormatDefinitionHeader = new()
            {
                FontName = HeaderStyleFontNameComboBox.Text ?? _defaultFontName,
                FontSize = HeaderStyleFontSizeTextBox.Text != string.Empty ? int.Parse(HeaderStyleFontSizeTextBox.Text) : _defaultHeaderFontSize,
                FontColor = GetSystemDrawingColorFromColorPicker(HeaderStyleFontColorPicker) ?? _defaultFontColor,
                FillColor = GetSystemDrawingColorFromColorPicker(HeaderStyleFillColorPicker) ?? _defaultFillColor,
                BorderColor = GetSystemDrawingColorFromColorPicker(HeaderStyleBorderColorPicker) ?? _defaultBorderColor,
                BorderThick = HeaderStyleBorderThickCheckBox.IsChecked ?? false,
                Bold = HeaderStyleBoldCheckBox.IsChecked ?? false,
                Italic = HeaderStyleItalicCheckBox.IsChecked ?? false,
                Underline = HeaderStyleUnderlineCheckBox.IsChecked ?? false
            };

            ExcelLogic.InsertHeaderLine(s, TableTextBox.Text, selectedColumns.Select(cm => cm.Name).ToList(), cellFormatDefinitionHeader);

            CellFormatDefinition cellFormatDefinitionData = new()
            {
                FontName = DataStyleFontNameComboBox.Text ?? _defaultFontName,
                FontSize = DataStyleFontSizeTextBox.Text != string.Empty ? int.Parse(DataStyleFontSizeTextBox.Text) : _defaultDataFontSize,
                FontColor = GetSystemDrawingColorFromColorPicker(DataStyleFontColorPicker) ?? _defaultFontColor,
                FillColor = GetSystemDrawingColorFromColorPicker(DataStyleFillColorPicker) ?? _defaultFillColor,
                BorderColor = GetSystemDrawingColorFromColorPicker(DataStyleBorderColorPicker) ?? _defaultBorderColor,
                BorderThick = DataStyleBorderThickCheckBox.IsChecked ?? false,
                Bold = DataStyleBoldCheckBox.IsChecked ?? false,
                Italic = DataStyleItalicCheckBox.IsChecked ?? false,
                Underline = DataStyleUnderlineCheckBox.IsChecked ?? false
            };

            foreach (var dataLine in dataLines)
                foreach (CellModel cm in dataLine)
                    cm.FormatDefinition = cellFormatDefinitionData;

            ExcelLogic.InsertDataLines(s, TableTextBox.Text, dataLines);
            s.SaveAndClose();

            ShowInformation(Properties.Resources.MW_INFO_FILE_CREATED.Replace("{FILE_PATH}", filePath));
        }

        private void TableTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdateColumnsListView(true);
        }

        private void HeaderStyleFontSizeTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            string text = ((System.Windows.Controls.TextBox)e.OriginalSource).Text;
            if (text != string.Empty)
            {
                if (!int.TryParse(text, out int fontSize) || fontSize <= 0 || fontSize > 100)
                    InitHeaderStyleFontSizeTextBox();
            }
        }

        private System.Drawing.Color? GetSystemDrawingColorFromColorPicker(Xceed.Wpf.Toolkit.ColorPicker colorPicker)
        {
            if (colorPicker.SelectedColor == null)
                return null;

            return System.Drawing.Color.FromArgb(
                colorPicker.SelectedColor.Value.A,
                colorPicker.SelectedColor.Value.R,
                colorPicker.SelectedColor.Value.G,
                colorPicker.SelectedColor.Value.B
            );
        }

        private void DataStyleFontSizeTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            string text = ((System.Windows.Controls.TextBox)e.OriginalSource).Text;
            if (text != string.Empty)
            {
                if (!int.TryParse(text, out int fontSize) || fontSize <= 0 || fontSize > 100)
                    InitDataStyleFontSizeTextBox();
            }
        }
    }
}
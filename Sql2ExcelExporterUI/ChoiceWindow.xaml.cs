using System.Windows;

namespace Sql2ExcelExporterUI
{
    /// <summary>
    /// Interaktionslogik für DatabaseChoiceWindow.xaml
    /// </summary>
    public partial class ChoiceWindow : Window
    {
        //private readonly string _title = "";
        private string _database = "";
        private readonly List<string> _elements = [];

        public ChoiceWindow(string title, List<string> elements)
        {
            InitializeComponent();
            this.Title = title;
            _elements = elements;
            ElementsListView.ItemsSource = _elements;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            SetChoiceAndCloseWindow();
        }

        private void SetChoiceAndCloseWindow()
        {
            _database = (string)ElementsListView.SelectedValue;
            Close();
        }

        public string GetChoice()
        {
            return _database;
        }
    }
}

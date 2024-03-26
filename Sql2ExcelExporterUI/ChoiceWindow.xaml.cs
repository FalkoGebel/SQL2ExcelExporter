using System.Windows;

namespace Sql2ExcelExporterUI
{
    /// <summary>
    /// Interaktionslogik für DatabaseChoiceWindow.xaml
    /// </summary>
    public partial class ChoiceWindow : Window
    {
        private string _choice = "";
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
            _choice = (string)ElementsListView.SelectedValue;
            if (_choice == null)
                _choice = string.Empty;
            DialogResult = true;
            Close();
        }

        public string GetChoice()
        {
            return _choice;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}

using System.Windows;
using Autodesk.Revit.DB;

namespace ColumnDesign.UI
{
    public partial class Warning : Window
    {
        public bool ContinueOperation { get; set; }
        public Warning(string message)
        {
            InitializeComponent();
            Message.Text = message;
        }

        private void ButtonOk_OnClick(object sender, RoutedEventArgs e)
        {
            ContinueOperation = true;
            DialogResult = true;
        }

        private void ButtonCancel_OnClick(object sender, RoutedEventArgs e)
        {
            ContinueOperation = false;
            DialogResult = true;
        }
    }
}
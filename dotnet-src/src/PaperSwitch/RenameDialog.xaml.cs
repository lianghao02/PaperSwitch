using System.Windows;

namespace PaperSwitch
{
    public partial class RenameDialog : Window
    {
        public string ProposedName => NameTextBox.Text.Trim();

        public RenameDialog(string currentName, int selectedCount)
        {
            InitializeComponent();
            NameTextBox.Text = currentName;
            NameTextBox.SelectAll();
            NameTextBox.Focus();
        }

        private void Confirm_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}

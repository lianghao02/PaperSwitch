using System.Windows;

namespace PaperSwitch
{
    public partial class RenameDialog : Window
    {
        public string ProposedName => NameTextBox.Text.Trim();

        public RenameDialog(string currentName, int selectedCount, string? heading = null, string? explanation = null, string? applyText = null)
        {
            InitializeComponent();
            if (heading != null) { Title = heading; Heading.Text = heading; }
            if (explanation != null) Explanation.Text = explanation;
            if (applyText != null) ApplyButton.Content = applyText;
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

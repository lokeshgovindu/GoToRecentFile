using System.Globalization;
using System.Windows;

namespace GoToRecentFile.Options
{
    public partial class SettingEditorDialog : Window
    {
        private readonly string _type;

        public string SettingValue => ValueBox.Text;

        public SettingEditorDialog(string name, string currentValue, string type)
        {
            InitializeComponent();
            _type = type;
            NameBox.Text = name;
            ValueBox.Text = currentValue;
            ValueBox.Focus();
            ValueBox.SelectAll();
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            string error = Validate(_type, ValueBox.Text);
            if (error != null)
            {
                ErrorText.Text = error;
                ErrorText.Visibility = Visibility.Visible;
                ValueBox.Focus();
                ValueBox.SelectAll();
                return;
            }

            DialogResult = true;
            Close();
        }

        private static string Validate(string type, string value)
        {
            switch (type.ToLowerInvariant())
            {
                case "integer":
                    if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                        return "Value must be a valid integer (e.g., 100, 760, 12).";
                    return null;
                case "boolean":
                    if (!string.Equals(value, "True", System.StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(value, "False", System.StringComparison.OrdinalIgnoreCase))
                        return "Value must be either 'True' or 'False'.";
                    return null;
                case "string":
                    return null;
                default:
                    return null;
            }
        }
    }
}

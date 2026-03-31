using System.Windows;

namespace PDnDocuments.Wpf;

public partial class InputDialog : Window
{
    public InputDialog(string title, string prompt, string initialValue = "")
    {
        InitializeComponent();
        TitleTextBlock.Text = title;
        PromptTextBlock.Text = prompt;
        ValueTextBox.Text = initialValue;
        Loaded += (_, _) =>
        {
            ValueTextBox.Focus();
            ValueTextBox.SelectAll();
            ValueTextBox.CaretIndex = ValueTextBox.Text.Length;
        };
    }

    public string Value => ValueTextBox.Text;

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}

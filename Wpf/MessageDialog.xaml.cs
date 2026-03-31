using System.Windows;
using System.Windows.Media;

namespace PDnDocuments.Wpf;

public partial class MessageDialog : Window
{
    public MessageDialog(
        string badgeText,
        string title,
        string message,
        string confirmText,
        string? cancelText = null,
        bool useDangerAccent = false)
    {
        InitializeComponent();

        BadgeTextBlock.Text = badgeText;
        TitleTextBlock.Text = title;
        MessageTextBlock.Text = message;
        ConfirmButton.Content = confirmText;

        CancelButton.Visibility = string.IsNullOrWhiteSpace(cancelText) ? Visibility.Collapsed : Visibility.Visible;
        if (!string.IsNullOrWhiteSpace(cancelText))
        {
            CancelButton.Content = cancelText;
        }

        if (useDangerAccent)
        {
            BadgeTextBlock.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFE8B7C6"));
            HeroBorder.Background = (Brush)FindResource("DangerBrush");
            HeroBorder.BorderBrush = (Brush)FindResource("DangerBorderBrush");
            HeroGlyphTextBlock.Foreground = Brushes.White;
            HeroGlyphTextBlock.Text = "!";
            ConfirmButton.Style = (Style)FindResource("DangerButtonStyle");
            MessagePanelBorder.BorderBrush = (Brush)FindResource("DangerBorderBrush");
        }
        else
        {
            HeroGlyphTextBlock.Text = "?";
        }

        Loaded += (_, _) => ConfirmButton.Focus();
    }

    public static void ShowMessage(Window? owner, string title, string message, string badgeText = "Сообщение", bool useDangerAccent = false)
    {
        var dialog = new MessageDialog(badgeText, title, message, "Ок", null, useDangerAccent)
        {
            Owner = owner,
        };
        dialog.ShowDialog();
    }

    public static bool ShowConfirm(
        Window? owner,
        string title,
        string message,
        string confirmText = "Да",
        string cancelText = "Нет",
        bool useDangerAccent = false)
    {
        var dialog = new MessageDialog("Подтверждение", title, message, confirmText, cancelText, useDangerAccent)
        {
            Owner = owner,
        };
        return dialog.ShowDialog() == true;
    }

    private void ConfirmButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}

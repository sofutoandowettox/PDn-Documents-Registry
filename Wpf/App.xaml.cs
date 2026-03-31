using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;

namespace PDnDocuments.Wpf;

public partial class App : Application
{
    public static string CurrentUserLabel => $"{Environment.UserName}@{Environment.MachineName}";

    protected override void OnStartup(StartupEventArgs e)
    {
        try
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            CultureInfo.DefaultThreadCurrentCulture = new CultureInfo("ru-RU");
            CultureInfo.DefaultThreadCurrentUICulture = new CultureInfo("ru-RU");

            if (e.Args.Any(arg => string.Equals(arg, "--smoke-test", StringComparison.OrdinalIgnoreCase)))
            {
                var store = new WorkspaceStore(AppContext.BaseDirectory, CurrentUserLabel);
                var state = store.LoadState();
                store.SaveState(state);
                Shutdown(0);
                return;
            }

            base.OnStartup(e);
            var window = new MainWindow();
            MainWindow = window;
            window.Show();
        }
        catch (Exception ex)
        {
            TryWriteStartupLog(ex);
            try
            {
                MessageDialog.ShowMessage(
                    owner: null,
                    title: "Ошибка запуска",
                    message: $"Приложение не смогло запуститься.{Environment.NewLine}{Environment.NewLine}{ex.Message}",
                    badgeText: "Ошибка",
                    useDangerAccent: true);
            }
            catch
            {
                MessageBox.Show(
                    $"Приложение не смогло запуститься.{Environment.NewLine}{Environment.NewLine}{ex.Message}",
                    "Ошибка запуска",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            Shutdown(1);
        }
    }

    private static void TryWriteStartupLog(Exception ex)
    {
        try
        {
            var logPath = Path.Combine(AppContext.BaseDirectory, "startup-error.log");
            File.WriteAllText(logPath, $"{DateTimeOffset.Now:O}{Environment.NewLine}{ex}", Encoding.UTF8);
        }
        catch
        {
            // Ignore secondary logging failures.
        }
    }
}

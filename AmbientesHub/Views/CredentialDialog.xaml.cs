using System.Windows;
using System.Windows.Input;
using Wpf.Ui.Controls;

namespace AmbientesHub.Views;

public partial class CredentialDialog : FluentWindow
{
    public string Username { get; private set; } = string.Empty;
    public string Password { get; private set; } = string.Empty;

    public CredentialDialog(string defaultUsername)
    {
        InitializeComponent();
        TxtUser.Text = defaultUsername;
        Loaded += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(TxtUser.Text))
                TxtUser.Focus();
            else
                TxtPassword.Focus();
        };
    }

    private void BtnSave_Click(object sender, RoutedEventArgs e)
    {
        string user = TxtUser.Text.Trim();
        string pass = TxtPassword.Password;

        if (string.IsNullOrWhiteSpace(user))
        {
            Services.AppLoggerService.Warning("Tentativa de salvar credencial sem informar o nome do usuário.", nameof(CredentialDialog));
            ShowError("Informe o nome do usuário.");
            TxtUser.Focus();
            return;
        }

        if (string.IsNullOrEmpty(pass))
        {
            Services.AppLoggerService.Warning($"Tentativa de salvar credencial com senha vazia para o usuário '{user}'.", nameof(CredentialDialog));
            ShowError("A senha não pode estar vazia.");
            TxtPassword.Focus();
            return;
        }

        Username = user;
        Password = pass;
        Services.AppLoggerService.Info($"Credencial confirmada no diálogo para usuário '{user}' (senha omitida).", nameof(CredentialDialog));
        DialogResult = true;
        Close();
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        Services.AppLoggerService.Info("Diálogo de credencial cancelado pelo usuário.", nameof(CredentialDialog));
        DialogResult = false;
        Close();
    }

    private void TxtPassword_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            BtnSave_Click(sender, e);
        }
    }

    private void ShowError(string message)
    {
        TxtError.Text = message;
        TxtError.Visibility = Visibility.Visible;
    }
}


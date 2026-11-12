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
            ShowError("Informe o nome do usuário.");
            TxtUser.Focus();
            return;
        }

        if (string.IsNullOrEmpty(pass))
        {
            ShowError("A senha não pode estar vazia.");
            TxtPassword.Focus();
            return;
        }

        Username = user;
        Password = pass;
        DialogResult = true;
        Close();
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
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


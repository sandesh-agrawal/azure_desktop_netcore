using Microsoft.Identity.Client;
using System;
using System.ComponentModel;
using System.Configuration;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows.Input;

namespace AzureADApp
{
    public class MainWindowsViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private const string _graphUrl = "https://graph.microsoft.com/v1.0/me";
        private string[] scopes = new string[] { "user.read" };
        private IPublicClientApplication app;

        public bool UseDeviceCode { get; set; }
        public ICommand LoginCommand { get; private set; }
        public ICommand LogoutCommand { get; private set; }

        private readonly string _clientId = "4a1aa1d5-c567-49d0-ad0b-cd957a47f842";

        private bool _isLogged = false;

        private string _tenandId;
        public string TenantId
        {
            get { return _tenandId; }
            set
            {
                if (value != null && _tenandId != value)
                {
                    _tenandId = value;
                    LoginEnabled = !(String.IsNullOrWhiteSpace(_userName) || String.IsNullOrWhiteSpace(_tenandId));
                    app = PublicClientApplicationBuilder
                                            .Create(_clientId)
                                            .WithAuthority($"https://login.microsoftonline.com/{_tenandId}")
                                            .WithDefaultRedirectUri()
                                            .Build();
                    RaisePropertyChanged();
                }
            }
        }

        private string _userName;
        public string UserName
        {
            get
            {
                return _userName;
            }
            set
            {
                if (value != null && _userName != value)
                {
                    _userName = value;
                    LoginEnabled = !(String.IsNullOrWhiteSpace(_userName) || String.IsNullOrWhiteSpace(_tenandId));
                    RaisePropertyChanged();
                }
            }
        }

        private bool _canLogin;

        public bool LoginEnabled
        {
            get { return _canLogin; }
            private set
            {
                if (value != _canLogin)
                {
                    _canLogin = value;
                    RaisePropertyChanged();
                }
            }
        }

        private bool _canLogout;

        public bool LogoutEnabled
        {
            get { return _canLogout; }
            set
            {
                if (value != _canLogout)
                {
                    _canLogout = value;
                    RaisePropertyChanged();
                }
            }
        }

        private string _response;
        public string Response
        {
            get
            {
                return _response;
            }
            private set
            {
                _response = value;
                RaisePropertyChanged();
            }
        }

        public MainWindowsViewModel()
        {
            TenantId = ConfigurationManager.AppSettings?.Get("TenantID") ?? "common";
            UserName = ConfigurationManager.AppSettings?.Get("ClientId") ?? "";
            LoginCommand = new RelayCommand(Login, CanLogin);
            LogoutCommand = new RelayCommand(Logout, CanLogout);
            Response = "";

        }

        private async void Login(object obj)
        {
            try
            {
                LoginEnabled = false;
                AuthenticationResult authResult = null;
                if (!UseDeviceCode)
                {
                    authResult = await app.AcquireTokenInteractive(scopes)
                                    .WithLoginHint(_userName)
                                    .WithPrompt(Prompt.SelectAccount)
                                    .ExecuteAsync();
                }
                else
                    authResult = await app.AcquireTokenWithDeviceCode(scopes, DeviceCodeFunction).ExecuteAsync();


                var httpClient = new HttpClient();
                HttpResponseMessage response;
                var request = new HttpRequestMessage(HttpMethod.Get, _graphUrl);
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                response = await httpClient.SendAsync(request);
                Response = await response.Content.ReadAsStringAsync();

                _isLogged = true;
                LogoutEnabled = true;


            }
            catch (Exception ex)
            {

                Response = $"Failed with exception {ex.Message}";
            }
        }

        private async Task DeviceCodeFunction(DeviceCodeResult codeResult)
        {
            Response = codeResult.Message;
            await Task.Delay(1000);
        }

        private bool CanLogin(object param)
        {
            return !_isLogged && !(String.IsNullOrWhiteSpace(_userName) || String.IsNullOrWhiteSpace(_tenandId)); ;
        }

        private bool CanLogout(object param)
        {
            return _isLogged;

        }

        private async void Logout(object param)
        {
            try
            {

                var accounts = await app.GetAccountsAsync();
                foreach (var account in accounts)
                {
                    await app.RemoveAsync(account);
                }
                Response = "Successfully logged out";
                _isLogged = false;
                LoginEnabled = true;
            }
            catch (MsalException ex)
            {
                Response = $"Error signing-out user: {ex.Message}";
            }
        }

        private void RaisePropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        }
    }

    public class RelayCommand : ICommand
    {

        readonly Action<object> _execute;
        readonly Predicate<object> _canExecute;

        public RelayCommand(Action<object> execute)
        : this(execute, null)
        {
        }

        public RelayCommand(Action<object> execute, Predicate<object> canExecute)
        {
            if (execute == null)
            {
                throw new ArgumentNullException("execute");
            }

            _execute = execute;
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            return _canExecute == null || _canExecute(parameter);
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public void Execute(object parameter)
        {
            _execute(parameter);
        }

    }
}

using Gma.System.MouseKeyHook;
using System.Configuration;
using System.Windows;
using System.Windows.Forms;
using SpotifyAPI.Web; //Base Namespace
using SpotifyAPI.Web.Enums; //Enums
using SpotifyAPI.Web.Models; //Models for the JSON-responses
using SpotifyAPI.Web.Auth;
using System.Threading.Tasks;
using System;

namespace SpotifyVolumeController.UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private IKeyboardMouseEvents GlobalHook;

        private readonly int DeltaOffset = -120;

        private bool IsDebug { get => bool.Parse(ConfigurationManager.AppSettings["Debug"]); }

        private string SpotifyClientId { get => ConfigurationManager.AppSettings["SpotifyClientId"]; }

        private string SpotifyClientSecret { get => ConfigurationManager.AppSettings["SpotifyClientSecret"];  }

        private SpotifyWebAPI SpotifyWebAPI;
        ImplicitGrantAuth ImplicitGrantAuth;

        public MainWindow()
        {
            InitializeComponent();

            DebugBox.IsEnabled = IsDebug;

            var golbalHookSuccessful = SubscribeGlobalHook();
            if (IsDebug)
            {
                DebugBox.Items.Add($"Hook bound: {golbalHookSuccessful}");
            }

        }

        #region Keyboard Hook


        public bool SubscribeGlobalHook()
        {
            try
            {
                GlobalHook = Hook.GlobalEvents();
                GlobalHook.MouseHWheel += GlobalMouseHWheel;
                return true;
            }
            catch
            {
                Unsubscribe();
                return false;
            }
        }

        private async void GlobalMouseHWheel(object sender, MouseEventArgs mouseEventArgs)
        {
            var delta = mouseEventArgs.Delta / DeltaOffset;

            if (IsDebug)
            {
                DebugBox.Items.Add(delta);
            }

            var playbackContext = await SpotifyWebAPI.GetPlaybackAsync();

            if (playbackContext != null && playbackContext.IsPlaying && !playbackContext.HasError())
            {
                var playbackPercent = playbackContext.Device.VolumePercent;

                var response = SpotifyWebAPI.SetVolumeAsync(playbackPercent + delta);
            }

        }

        public void Unsubscribe()
        {
            if (GlobalHook != null)
            {
                GlobalHook.MouseHWheel -= GlobalMouseHWheel;

                GlobalHook.Dispose();
                if (IsDebug)
                {
                    DebugBox.Items.Add($"Hook unbound: true");
                }
            }
        }

        #endregion Keyboard Hook

        #region UI Event Handlers


        private void Window_Unloaded(object sender, RoutedEventArgs e)
        {
            Unsubscribe();
        }

        #endregion UI Event Handlers

        private void Auth_Click(object sender, RoutedEventArgs e)
        {
            //https://johnnycrazy.github.io/SpotifyAPI-NET/docs/5.1.1/auth/implicit_grant
            ImplicitGrantAuth = new ImplicitGrantAuth(
                SpotifyClientId,
                "http://localhost:4002",
                "http://localhost:4002",
                Scope.UserReadPrivate
            );
            ImplicitGrantAuth.AuthReceived += async (senders, payload) =>
            {
                ImplicitGrantAuth.Stop(); // `sender` is also the auth instance
                SpotifyWebAPI = new SpotifyWebAPI()
                {
                    TokenType = payload.TokenType,
                    AccessToken = payload.AccessToken
                };
            };
            ImplicitGrantAuth.Start(); // Starts an internal HTTP Server

            ImplicitGrantAuth.OpenBrowser();
        }
    }
}

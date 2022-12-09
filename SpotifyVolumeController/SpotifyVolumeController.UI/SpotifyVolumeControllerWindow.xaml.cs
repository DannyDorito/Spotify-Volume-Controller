using Gma.System.MouseKeyHook;
using System.Configuration;
using System.Windows;
using System.Windows.Forms;
using SpotifyAPI.Web;
using SpotifyAPI.Web.Enums;
using SpotifyAPI.Web.Auth;

namespace SpotifyVolumeController.UI
{
    /// <summary>
    /// Interaction logic for SpotifyVolumeControllerWindow.xaml
    /// </summary>
    public partial class SpotifyVolumeControllerWindow : Window
    {
        #region Mouse Wheel Hook Variables

        private IKeyboardMouseEvents GlobalHook;

        /// <summary>
        /// Mousewheel delta offset for the mousewheel movement
        /// </summary>
        private int DeltaOffset { get => int.Parse(ConfigurationManager.AppSettings["DeltaOffset"]); }

        /// <summary>
        /// Minimum volume clamp value
        /// </summary>
        private int ClampMin { get => 0; }

        /// <summary>
        /// Maximum volume clamp value
        /// </summary>
        private int ClampMax { get => 100; }

        #endregion Mouse Wheel Hook Variables

        #region Debug Variables

        /// <summary>
        /// Debug mode
        /// </summary>
        private bool IsDebug { get => bool.Parse(ConfigurationManager.AppSettings["IsDebug"]); }

        #endregion Debug Variables

        #region Spotify API Variables

        /// <summary>
        /// Client id for the Spotify API
        /// </summary>
        private string ClientId { get => ConfigurationManager.AppSettings["ClientId"]; }

        /// <summary>
        /// Authorisation <seealso cref="Scope" for Spotify API/>
        /// </summary>
        private Scope AuthScope { get => Scope.UserReadPlaybackState | Scope.UserReadCurrentlyPlaying | Scope.UserModifyPlaybackState; }

        /// <summary>
        /// Local redirect uri for authorisation server
        /// </summary>
        public string RedirectURL { get => ConfigurationManager.AppSettings["RedirectURL"]; }

        /// <summary>
        /// Local server uri for authorisation
        /// </summary>
        public string ServerURL { get => ConfigurationManager.AppSettings["ServerURL"]; }

        /// <summary>
        /// Spotify web API
        /// </summary>
        private SpotifyWebAPI SpotifyWebAPI;

        /// <summary>
        /// Authorisation for the API, set in <see cref="SpotifyVolumeControllerWindow"/>
        /// </summary>
        private readonly ImplicitGrantAuth ImplicitGrantAuth;

        #endregion Spotify API Variables

        public SpotifyVolumeControllerWindow()
        {
            InitializeComponent();

            DebugListBox.IsEnabled = IsDebug;

            DebugLog($"Hook bound: {SubscribeGlobalHook()}");

            ImplicitGrantAuth = new ImplicitGrantAuth(
              ClientId,
              RedirectURL,
              ServerURL,
              AuthScope
            );
        }

        #region Keyboard Hook

        /// <summary>
        /// Set up the hook
        /// </summary>
        /// <returns></returns>
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
                UnsubscribeGlobalHook();
                return false;
            }
        }

        /// <summary>
        /// Dispose of the hook
        /// </summary>
        public void UnsubscribeGlobalHook()
        {
            if (GlobalHook != null)
            {
                GlobalHook.MouseHWheel -= GlobalMouseHWheel;

                GlobalHook.Dispose();
                DebugLog("Hook unbound");
            }
        }

        /// <summary>
        /// Set the volume with the <see cref="SpotifyAPI"/> using the mousewheel delta - <see cref="DeltaOffset"/>
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="mouseEventArgs"></param>
        private async void GlobalMouseHWheel(object sender, MouseEventArgs mouseEventArgs)
        {
            if (mouseEventArgs != null)
            {
                var delta = mouseEventArgs.Delta / DeltaOffset;

                DebugLog($"Delta : {delta}");

                // get the playback context from the API
                var playbackContext = await SpotifyWebAPI.GetPlaybackAsync();

                if (playbackContext != null && !playbackContext.HasError())
                {
                    // TODO: save previous volume and negate that from current volume to prevent volume jumps
                    var currentVolume = playbackContext.Device.VolumePercent;
                    // clamp the desired volume between 0 and 100
                    var desiredVolume = Clamp(currentVolume + delta, ClampMin, ClampMax);

                    // prevent unnecessary API calls if the value is the same e.g. volume may have been clamped
                    if (currentVolume != desiredVolume)
                    {
                        var response = await SpotifyWebAPI.SetVolumeAsync(desiredVolume);

                        // only log if there is an error
                        if (response != null && response.HasError())
                        {
                            DebugLog(response.Error.Message);
                        }
                    }
                }
            }
        }

        #endregion Keyboard Hook

        #region UI Event Handlers

        /// <summary>
        /// Unsubscribe from hook when exiting the window
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Unloaded(object sender, RoutedEventArgs e)
        {
            UnsubscribeGlobalHook();
        }

        /// <summary>
        /// https://johnnycrazy.github.io/SpotifyAPI-NET/docs/5.1.1/auth/implicit_grant
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Auth_Click(object sender, RoutedEventArgs e)
        {
            ImplicitGrantAuth.AuthReceived += (_, payload) =>
            {
                ImplicitGrantAuth.Stop();
                SpotifyWebAPI = new SpotifyWebAPI()
                {
                    TokenType = payload.TokenType,
                    AccessToken = payload.AccessToken
                };
            };

            ImplicitGrantAuth.Start();
            ImplicitGrantAuth.OpenBrowser();
            DebugLog("Opened Auth Browser");
        }

        #endregion UI Event Handlers

        #region Helper Methods
        
        /// <summary>
        /// Log a message to <see cref="DebugBox"/> if <see cref="IsDebug"/> is tue
        /// </summary>
        /// <param name="message">Message to log to <see cref="DebugBox"/></param>
        private void DebugLog(object message)
        {
            if (IsDebug)
            {
                DebugListBox.Items.Add(message as string);
            }
        }

        /// <summary>
        /// https://stackoverflow.com/questions/3176602/how-to-force-a-number-to-be-in-a-range-in-c
        /// </summary>
        /// <param name="value">Value to be clamed</param>
        /// <param name="min">Minimum clamp value <seealso cref="ClampMin"/></param>
        /// <param name="max">Maximum clamp value <seealso cref="ClampMax"/></param>
        /// <returns>Clamped <see cref="int"/> <paramref name="value"/> between <paramref name="min"/> and <paramref name="max"/></returns>
        private int Clamp(int value, int min, int max)
        {
            return (value < min) ? min : (value > max) ? max : value;
        }

        #endregion Helper Methods
    }
}

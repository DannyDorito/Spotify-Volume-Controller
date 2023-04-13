﻿using Gma.System.MouseKeyHook;
using System.Configuration;
using System.Windows;
using System.Windows.Forms;
using System.Diagnostics;
using System;
using SpotifyAPI.Web.Auth;
using SpotifyAPI.Web;
using System.Collections.Generic;
using System.Threading.Tasks;

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
        private static int DeltaOffset { get => int.Parse(ConfigurationManager.AppSettings["DeltaOffset"]); }

        /// <summary>
        /// Minimum volume clamp value
        /// </summary>
        private static int ClampMin { get => 0; }

        /// <summary>
        /// Maximum volume clamp value
        /// </summary>
        private static int ClampMax { get => 100; }

        /// <summary>
        /// Previous volume of the client
        /// </summary>
        private int PreviousVolume { get; set; }

        #endregion Mouse Wheel Hook Variables

        #region Debug Variables

        /// <summary>
        /// Debug mode
        /// </summary>
        private static bool IsDebug { get => bool.Parse(ConfigurationManager.AppSettings["IsDebug"]); }

        #endregion Debug Variables

        #region Spotify API Variables

        /// <summary>
        /// Client id for the Spotify API
        /// </summary>
        private static string ClientId { get => ConfigurationManager.AppSettings["ClientId"]; }

        /// <summary>
        /// Authorisation <seealso cref="Scope" for Spotify API/>
        /// </summary>
        private static List<string> Scope => new() { Scopes.UserReadPlaybackState, Scopes.UserReadCurrentlyPlaying, Scopes.UserModifyPlaybackState };

        /// <summary>
        /// Local redirect uri for authorisation server
        /// </summary>
        public static string RedirectURL { get => ConfigurationManager.AppSettings["RedirectURL"]; }

        /// <summary>
        /// Local server uri for authorisation
        /// </summary>
        public static string ServerURL { get => ConfigurationManager.AppSettings["ServerURL"]; }

        /// <summary>
        /// Spotify web API
        /// </summary>
        private SpotifyClient spotifyClient;

        private static EmbedIOAuthServer _server;


        #endregion Spotify API Variables

        public SpotifyVolumeControllerWindow()
        {
            InitializeComponent();
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

                DebugLog($"Delta: {delta}");

                // get the playback context from the API
                var playbackContext = spotifyClient.Player;

                if (playbackContext is null)
                {
                    return;
                }
                var devices = await spotifyClient.Player.GetAvailableDevices();

                if (playbackContext != null && devices.Devices.Count >= 0)
                {
                    var device = devices.Devices[0];
                    var currentVolume = device.VolumePercent;
                    // clamp the desired volume between 0 and 100
                    var desiredVolume = Clamp(currentVolume.Value + delta, ClampMin, ClampMax);
                    // get change in value
                    var volumeChange = Math.Abs(desiredVolume - PreviousVolume);
                    if (volumeChange > 0)
                    {
                        // prevent unnecessary API calls if the value is the same e.g. volume may have been clamped
                        if (currentVolume != desiredVolume)
                        {
                            PreviousVolume = desiredVolume;
                            device.VolumePercent = desiredVolume;
                        }
                    }
                }
                else
                {
                    DebugLog("Playback context cannot be found");
                }
            }
            else
            {
                DebugLog("Mouse Event was not bound");
            }
        }

        private async Task OnImplicitGrantReceived(object sender, ImplictGrantResponse response)
        {
            await _server.Stop();
            spotifyClient = new SpotifyClient(response.AccessToken);
        }

        private async Task OnErrorReceived(object sender, string error, string state)
        {
            DebugLog($"Aborting authorization, error received: {error}");
            await _server.Stop();
        }

        #endregion Keyboard Hook

        #region UI Event Handlers

        /// <summary>
        /// Set debug and auth when window is fully loaded
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            DebugListBox.IsEnabled = IsDebug;

            DebugLog($"Hook bound: {SubscribeGlobalHook()}");
        }

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
        private async void Auth_Click(object sender, RoutedEventArgs e)
        {
            _server = new EmbedIOAuthServer(new Uri("http://localhost:5000/callback"), 5000);
            await _server.Start();

            _server.ImplictGrantReceived += OnImplicitGrantReceived;
            _server.ErrorReceived += OnErrorReceived;

            var request = new LoginRequest(_server.BaseUri, "ClientId", LoginRequest.ResponseType.Token)
            {
                Scope = new List<string> { Scopes.UserReadEmail }
            };
            BrowserUtil.Open(request.ToUri());
            DebugLog("Opened Auth Browser");
        }

        /// <summary>
        /// Opens settings window
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsDebug)
            {
                var settingsWindow = new SettingsWindow();
                settingsWindow.ShowDialog();
            }
        }

        #endregion UI Event Handlers

        #region Helper Methods

        /// <summary>
        /// Log a message to <see cref="DebugBox"/> if <see cref="IsDebug"/> is tue
        /// </summary>
        /// <param name="message">Message to log to <see cref="DebugBox"/></param>
        private void DebugLog(object message)
        {
            var messageString = message as string;

            Debug.WriteLine(messageString);

            if (IsDebug)
            {
                DebugListBox.Items.Add(messageString);
            }
        }

        /// <summary>
        /// https://stackoverflow.com/questions/3176602/how-to-force-a-number-to-be-in-a-range-in-c
        /// </summary>
        /// <param name="value">Value to be clamed</param>
        /// <param name="min">Minimum clamp value <seealso cref="ClampMin"/></param>
        /// <param name="max">Maximum clamp value <seealso cref="ClampMax"/></param>
        /// <returns>Clamped <see cref="int"/> <paramref name="value"/> between <paramref name="min"/> and <paramref name="max"/></returns>
        private static int Clamp(int value, int min, int max)
        {
            return (value < min) ? min : (value > max) ? max : value;
        }

        #endregion Helper Methods
    }
}

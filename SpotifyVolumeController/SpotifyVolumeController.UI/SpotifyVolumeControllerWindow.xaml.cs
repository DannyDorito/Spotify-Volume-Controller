using Gma.System.MouseKeyHook;
using System.Configuration;
using System.Windows;
using System.Windows.Forms;
using System.Diagnostics;
using System;
using SpotifyAPI.Web.Auth;
using SpotifyAPI.Web;
using System.Collections.Generic;
using System.Runtime.Versioning;
using System.Globalization;

namespace SpotifyVolumeController.UI;

/// <summary>
/// Interaction logic for SpotifyVolumeControllerWindow.xaml
/// </summary>
[System.Diagnostics.CodeAnalysis.SuppressMessage("Maintainability", "CA1515:Consider making public types internal", Justification = "Required")]
public partial class SpotifyVolumeControllerWindow : Window
{
    #region Mouse Wheel Hook Variables

    private IKeyboardMouseEvents GlobalHook;

    /// <summary>
    /// Mousewheel delta offset for the mousewheel movement
    /// </summary>
    private static int DeltaOffset { get => int.Parse(ConfigurationManager.AppSettings["DeltaOffset"], NumberStyles.Integer, CultureInfo.InvariantCulture); }

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
    /// Client Id for the Spotify API
    /// </summary>
    private static string ClientId { get => ConfigurationManager.AppSettings["ClientId"]; }

    /// <summary>
    /// Client Secret for the Spotify API
    /// </summary>
    private static string ClientSecret { get => ConfigurationManager.AppSettings["ClientSecret"]; }

    /// <summary>
    /// Authorisation <seealso cref="Scope" for Spotify API/>
    /// </summary>
    private static List<string> Scope => [Scopes.UserReadPlaybackState, Scopes.UserReadCurrentlyPlaying, Scopes.UserModifyPlaybackState];

    /// <summary>
    /// Local redirect uri for authorisation server
    /// </summary>
    public static Uri RedirectURL { get => new(ConfigurationManager.AppSettings["RedirectURL"]); }

    /// <summary>
    /// Local server uri for authorisation
    /// </summary>
    public static Uri ServerURL { get => new(ConfigurationManager.AppSettings["ServerURL"]); }

    /// <summary>
    /// Local server port for authorisation
    /// </summary>
    public static int ServerPort { get => int.Parse(ConfigurationManager.AppSettings["ServerPort"], NumberStyles.Integer, CultureInfo.InvariantCulture); }

    /// <summary>
    /// Spotify web API
    /// </summary>
    private SpotifyClient spotifyClient;


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
    [SupportedOSPlatform("windows7.0")]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "Justified catch-all")]
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
    [SupportedOSPlatform("windows7.0")]
    public void UnsubscribeGlobalHook()
    {
        if (GlobalHook is not null)
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
    [SupportedOSPlatform("windows7.0")]
    private async void GlobalMouseHWheel(object sender, MouseEventArgs mouseEventArgs)
    {
        if (mouseEventArgs is not null)
        {
            var delta = mouseEventArgs.Delta / DeltaOffset;

            DebugLog($"Delta: {delta}");

            if (spotifyClient is null)
            {
                DebugLog("Spotify Client cannot be found");
                return;
            }

            // get the playback context from the API
            var playbackContext = spotifyClient.Player;
            var devices = await spotifyClient.Player.GetAvailableDevices().ConfigureAwait(true);

            if (playbackContext is not null && devices.Devices.Count > 0)
            {
                foreach (var device in devices.Devices)
                {
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

    #endregion Keyboard Hook

    #region UI Event Handlers

    /// <summary>
    /// Set debug and auth when window is fully loaded
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    [SupportedOSPlatform("windows7.0")]
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
    [SupportedOSPlatform("windows7.0")]
    private void Window_Unloaded(object sender, RoutedEventArgs e)
    {
        UnsubscribeGlobalHook();
    }

    /// <summary>
    /// https://johnnycrazy.github.io/SpotifyAPI-NET/docs/5.1.1/auth/implicit_grant
    /// https://johnnycrazy.github.io/SpotifyAPI-NET/docs/5_to_6
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private async void Auth_Click(object sender, RoutedEventArgs e)
    {
        var config = SpotifyClientConfig.CreateDefault();
        using (var server = new EmbedIOAuthServer(RedirectURL, ServerPort))
        {
            server.AuthorizationCodeReceived += async (sender, response) =>
                    {
                        await server.Stop().ConfigureAwait(true);
                        var tokenResponse = await new OAuthClient(config).RequestToken(new AuthorizationCodeTokenRequest(
                          ClientId, ClientSecret, response.Code, server.BaseUri
                        )).ConfigureAwait(true);

                        AuthButton.Content = "Authorised";
                        AuthButton.IsEnabled = false;

                        spotifyClient = new SpotifyClient(config.WithToken(tokenResponse.AccessToken));
                        DebugLog("Authorization Code Received");
                    };
            server.ErrorReceived += async (sender, error, state) =>
            {
                DebugLog($"Aborting authorization, error received: {error}");
                await server.Stop().ConfigureAwait(true);
            };

            await server.Start().ConfigureAwait(true);
        }

        var request = new LoginRequest(RedirectURL, ClientId, LoginRequest.ResponseType.Code)
        {
            Scope = Scope
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

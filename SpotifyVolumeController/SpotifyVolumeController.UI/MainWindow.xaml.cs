using Gma.System.MouseKeyHook;
using System.Configuration;
using System.Windows;
using System.Windows.Forms;

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

        private void GlobalMouseHWheel(object sender, MouseEventArgs mouseEventArgs)
        {
            if (IsDebug)
            {
                DebugBox.Items.Add($"{mouseEventArgs.Delta / DeltaOffset} {mouseEventArgs.X},{mouseEventArgs.Y} {mouseEventArgs.Location}");
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

        private void Auth_Click(object sender, RoutedEventArgs e)
        {

        }

        #endregion UI Event Handlers
    }
}

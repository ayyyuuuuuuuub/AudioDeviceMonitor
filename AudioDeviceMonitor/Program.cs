using Microsoft.Toolkit.Uwp.Notifications;
using Microsoft.Win32;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using Windows.Media.Control;

namespace AudioDeviceMonitor;

static class Program
{
    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}

public class MainForm : Form
{
    #region UI Components
    private NotifyIcon _trayIcon = null!;
    private ContextMenuStrip _trayMenu = null!;
    private CheckBox _startupCheckBox = null!;
    private CheckBox _toastCheckBox = null!;
    private Label _infoLabel = null!;
    private Label _themeLabel = null!;
    private ComboBox _themeComboBox = null!;
    private Label _languageLabel = null!;
    private ComboBox _languageComboBox = null!;
    private FlowLayoutPanel _mainLayout = null!;
    #endregion

    #region Logic and State
    // This resolves the CS8618 warning by assuring the compiler it will be initialized.
    private readonly AudioDeviceChangeMonitor _monitor = null!;
    private const string AppName = "AudioDeviceMonitor";
    private const string RegistryKeyPath = @"SOFTWARE\AudioDeviceMonitor";
    private readonly Dictionary<string, Dictionary<string, string>> _uiStrings = new();
    #endregion

    public MainForm()
    {
        try
        {
            _monitor = new AudioDeviceChangeMonitor(this.OnMediaPaused);

            InitializeLocalization();
            InitializeUI();
            InitializeSystemTray();

            this.Load += MainForm_Load;
            this.FormClosing += MainForm_FormClosing;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"A critical error occurred on startup: {ex.Message}\n\n{ex.StackTrace}", "Application Startup Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Application.Exit();
        }
    }

    private void InitializeUI()
    {
        this.Text = AppName;
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.MaximizeBox = false;
        this.MinimizeBox = true;

        _mainLayout = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(10), WrapContents = false };
        this.Controls.Add(_mainLayout);

        _infoLabel = new Label { AutoSize = true, Margin = new Padding(3, 0, 3, 10) };
        _startupCheckBox = new CheckBox { AutoSize = true, Margin = new Padding(3, 3, 3, 3) };
        _toastCheckBox = new CheckBox { AutoSize = true, Margin = new Padding(3, 3, 3, 10) };

        var langPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(0) };
        _languageLabel = new Label { AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 6, 3, 3) };
        _languageComboBox = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 120, Margin = new Padding(3, 3, 3, 3) };
        _languageComboBox.Items.AddRange(new object[] { "English", "Suomi" });
        langPanel.Controls.AddRange(new Control[] { _languageLabel, _languageComboBox });

        var themePanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Margin = new Padding(0) };
        _themeLabel = new Label { AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 6, 3, 3) };
        _themeComboBox = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 120, Margin = new Padding(3, 3, 3, 3) };
        _themeComboBox.Items.AddRange(new object[] { "Light", "Dark" });
        themePanel.Controls.AddRange(new Control[] { _themeLabel, _themeComboBox });

        _mainLayout.Controls.AddRange(new Control[] { _infoLabel, _startupCheckBox, _toastCheckBox, langPanel, themePanel });

        _startupCheckBox.CheckedChanged += (s, e) => SaveSettings();
        _toastCheckBox.CheckedChanged += (s, e) => SaveSettings();
        _languageComboBox.SelectedIndexChanged += (s, e) => { UpdateUIText(); SaveSettings(); };
        _themeComboBox.SelectedIndexChanged += (s, e) => { ApplyTheme(); SaveSettings(); };
    }

    private void InitializeSystemTray()
    {
        _trayMenu = new ContextMenuStrip();
        this.Icon = CreateAppIcon();
        _trayIcon = new NotifyIcon { Text = AppName, Icon = this.Icon, Visible = true, ContextMenuStrip = _trayMenu };
        _trayIcon.DoubleClick += (s, e) => this.Show();
    }

    private Icon CreateAppIcon()
    {
        using (var bmp = new Bitmap(32, 32))
        using (var g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var p = new Pen(Color.DodgerBlue, 4))
            {
                g.DrawArc(p, 8, 5, 16, 18, 180, 180);
            }
            g.FillEllipse(Brushes.DodgerBlue, 4, 13, 10, 12);
            g.FillEllipse(Brushes.DodgerBlue, 18, 13, 10, 12);
            return Icon.FromHandle(bmp.GetHicon());
        }
    }

    #region Settings, Theming, and Localization
    private void LoadSettings()
    {
        // Suspend layout to prevent flickering and ensure all settings apply at once.
        this.SuspendLayout();

        using (RegistryKey? rk = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
        {
            _languageComboBox.SelectedIndex = Convert.ToInt32(rk?.GetValue("Language", 0));
            _themeComboBox.SelectedIndex = Convert.ToInt32(rk?.GetValue("Theme", 0));
            _toastCheckBox.Checked = Convert.ToBoolean(rk?.GetValue("ShowNotifications", true));
        }
        using (RegistryKey? rk = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", false))
        {
            _startupCheckBox.Checked = (rk?.GetValue(AppName) != null);
        }

        // Apply the loaded settings
        UpdateUIText();
        ApplyTheme();

        this.ResumeLayout();
    }

    private void SaveSettings()
    {
        using (RegistryKey rk = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
        {
            rk.SetValue("ShowNotifications", _toastCheckBox.Checked);
            rk.SetValue("Language", _languageComboBox.SelectedIndex);
            rk.SetValue("Theme", _themeComboBox.SelectedIndex);
        }
        using (RegistryKey rk = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true)!)
        {
            if (_startupCheckBox.Checked) rk.SetValue(AppName, Application.ExecutablePath);
            else rk.DeleteValue(AppName, false);
        }
    }

    private void InitializeLocalization()
    {
        _uiStrings["English"] = new Dictionary<string, string> {
            { "info", "The application is running in the system tray." }, { "startup", "Start with Windows" },
            { "toast", "Show notification when media is paused" }, { "language", "Language" },
            { "theme", "Theme" }, { "show", "Show" }, { "exit", "Exit" },
            { "paused_title", "Media Paused" }, { "paused_text", "Paused: {0} by {1}" }
        };
        _uiStrings["Suomi"] = new Dictionary<string, string> {
            { "info", "Sovellus on käynnissä ilmaisinalueella." }, { "startup", "Käynnistä Windowsin kanssa" },
            { "toast", "Näytä ilmoitus, kun media on keskeytetty" }, { "language", "Kieli" },
            { "theme", "Teema" }, { "show", "Näytä" }, { "exit", "Poistu" },
            { "paused_title", "Media Keskeytetty" }, { "paused_text", "Keskeytetty: {0} artistilta {1}" }
        };
    }

    private void UpdateUIText()
    {
        string lang = _languageComboBox.SelectedItem?.ToString() ?? "English";
        var d = _uiStrings[lang];
        _infoLabel.Text = d["info"]; _startupCheckBox.Text = d["startup"];
        _toastCheckBox.Text = d["toast"]; _languageLabel.Text = d["language"];
        _themeLabel.Text = d["theme"];
        _trayMenu.Items.Clear();
        _trayMenu.Items.Add(d["show"], null, (s, e) => this.Show());
        _trayMenu.Items.Add(d["exit"], null, OnExit);
    }

    private void ApplyTheme()
    {
        string theme = _themeComboBox.SelectedItem?.ToString() ?? "Light";
        var backColor = theme == "Dark" ? Color.FromArgb(45, 45, 48) : SystemColors.Control;
        var foreColor = theme == "Dark" ? Color.White : SystemColors.ControlText;
        this.BackColor = backColor; _mainLayout.BackColor = backColor;
        foreach (Control c in _mainLayout.Controls) c.ForeColor = foreColor;
    }
    #endregion

    #region Event Handlers
    private void MainForm_Load(object? sender, EventArgs e)
    {
        LoadSettings();
        this.ClientSize = new Size(_mainLayout.PreferredSize.Width + 20, _mainLayout.PreferredSize.Height + 20);
        this.BeginInvoke(new Action(() => this.Hide()));
    }

    private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
    {
        if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; this.Hide(); }
        else { SaveSettings(); }
    }

    public void OnMediaPaused(string title, string artist)
    {
        if (!_toastCheckBox.Checked) return;
        string lang = _languageComboBox.SelectedItem?.ToString() ?? "English";
        var d = _uiStrings[lang];
        new ToastContentBuilder().AddText(d["paused_title"]).AddText(string.Format(d["paused_text"], title, artist)).Show();
    }

    private void OnExit(object? sender, EventArgs e)
    {
        _trayIcon.Visible = false; _monitor.Dispose(); Application.Exit();
    }
    #endregion
}

public class AudioDeviceChangeMonitor : IDisposable, IMMNotificationClient
{
    private readonly MMDeviceEnumerator _deviceEnumerator;
    private string? _defaultDeviceId;
    private readonly Action<string, string> _onMediaPausedCallback;

    public AudioDeviceChangeMonitor(Action<string, string> onMediaPaused)
    {
        _onMediaPausedCallback = onMediaPaused;
        _deviceEnumerator = new MMDeviceEnumerator();
        try
        {
            var defaultDevice = _deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Console);
            _defaultDeviceId = defaultDevice.ID;
        }
        catch (COMException) { _defaultDeviceId = null; }
        _deviceEnumerator.RegisterEndpointNotificationCallback(this);
    }

    public void OnDeviceStateChanged(string pwstrDeviceId, DeviceState dwNewState)
    {
        if (pwstrDeviceId == _defaultDeviceId && dwNewState == DeviceState.Unplugged)
        {
            PauseMediaPlayback().GetAwaiter().GetResult();
        }
    }

    public void OnDefaultDeviceChanged(DataFlow flow, Role role, string? pwstrDefaultDeviceId)
    {
        if (flow == DataFlow.Render && (role == Role.Console || role == Role.Multimedia))
        {
            var previousDefaultDeviceId = _defaultDeviceId; _defaultDeviceId = pwstrDefaultDeviceId;
            if (previousDefaultDeviceId != null && previousDefaultDeviceId != _defaultDeviceId)
            {
                try
                {
                    if (_deviceEnumerator.GetDevice(previousDefaultDeviceId).State == DeviceState.Unplugged)
                    {
                        PauseMediaPlayback().GetAwaiter().GetResult();
                    }
                }
                catch (COMException) { PauseMediaPlayback().GetAwaiter().GetResult(); }
            }
        }
    }

    private async Task PauseMediaPlayback()
    {
        try
        {
            var sessions = (await GlobalSystemMediaTransportControlsSessionManager.RequestAsync()).GetSessions();
            foreach (var session in sessions)
            {
                if (session.GetPlaybackInfo().PlaybackStatus == GlobalSystemMediaTransportControlsSessionPlaybackStatus.Playing)
                {
                    if (await session.TryPauseAsync())
                    {
                        var props = await session.TryGetMediaPropertiesAsync();
                        _onMediaPausedCallback(props.Title, props.Artist);
                    }
                }
            }
        }
        catch (Exception) { /* Fails silently */ }
    }

    public void OnDeviceAdded(string pwstrDeviceId) { }
    public void OnDeviceRemoved(string pwstrDeviceId) { }
    public void OnPropertyValueChanged(string pwstrDeviceId, PropertyKey key) { }
    public void Dispose() => _deviceEnumerator.UnregisterEndpointNotificationCallback(this);
}

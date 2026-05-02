using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace SoundDeviceSwitcher
{
    internal sealed class MainForm : Form
    {
        private readonly CoreAudioController _audioController = new CoreAudioController();
        private readonly AppSettings _settings = new AppSettings();
        private readonly VoicemeeterRemoteClient _voicemeeter = new VoicemeeterRemoteClient();

        private TabControl _tabs;
        private TabPage _systemPage;
        private TabPage _voicemeeterPage;
        private TabPage _languagePage;
        private Button _systemRefreshButton;
        private Button _systemVmDefaultsButton;
        private Button _vmRefreshButton;
        private Button _vmRestartButton;
        private GroupBox _systemOutputGroup;
        private GroupBox _systemInputGroup;
        private GroupBox _vmInputGroup;
        private GroupBox _vmOutputGroup;
        private ListView _systemOutputList;
        private ListView _systemInputList;
        private ListView _vmInputList;
        private ListView _vmOutputList;
        private SplitContainer _systemSplit;
        private SplitContainer _systemVolumeSplit;
        private SplitContainer _vmSplit;
        private GroupBox _systemInputVolumeGroup;
        private GroupBox _systemOutputVolumeGroup;
        private Panel _systemInputVolumePanel;
        private Panel _systemOutputVolumePanel;
        private Label _systemStatus;
        private Label _vmStatus;
        private Label _languageTitle;
        private Label _languageDescription;
        private ComboBox _languageCombo;
        private ContextMenuStrip _languageMenu;
        private RadioButton _vmPhysicalOnly;
        private RadioButton _vmShowAll;

        private string _languageCode;
        private IList<AudioDevice> _systemOutputs = new List<AudioDevice>();
        private IList<AudioDevice> _systemInputs = new List<AudioDevice>();
        private bool _loadingSystem;
        private bool _loadingVoicemeeter;
        private bool _updatingLanguage;
        private bool _updatingVolumeControls;
        private bool _showingLanguageMenu;

        public MainForm()
        {
            Text = "Sound Device Switcher";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(900, 560);
            Size = new Size(1040, 680);
            Font = new Font("Segoe UI", 9F);
            _languageCode = _settings.LanguageCode;

            BuildInterface();
            ApplyLanguage();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            SetSplitEven(_systemSplit);
            SetSplitEven(_systemVolumeSplit);
            SetSplitEven(_vmSplit);
            RefreshSystemDevices();
            ApplyStartupDefaultsIfNeeded();
            RefreshSystemDevices();
            RefreshVoicemeeterDevices();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _voicemeeter.Dispose();
            base.OnFormClosed(e);
        }

        private void BuildInterface()
        {
            _tabs = new TabControl
            {
                Dock = DockStyle.Fill
            };

            _systemPage = BuildSystemTab();
            _voicemeeterPage = BuildVoicemeeterTab();
            _languagePage = BuildLanguageTab();

            _tabs.TabPages.Add(_systemPage);
            _tabs.TabPages.Add(_voicemeeterPage);
            _tabs.TabPages.Add(_languagePage);
            _tabs.Selecting += TabsSelecting;
            _tabs.SelectedTab = _voicemeeterPage;
            Controls.Add(_tabs);
        }

        private TabPage BuildSystemTab()
        {
            var page = new TabPage();

            var topPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 48,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(10, 9, 10, 8),
                WrapContents = false
            };

            _systemRefreshButton = new Button
            {
                Width = 116,
                Height = 28
            };
            _systemRefreshButton.Click += delegate { RefreshSystemDevices(); };

            _systemVmDefaultsButton = new Button
            {
                Width = 240,
                Height = 28
            };
            _systemVmDefaultsButton.Click += delegate { ApplyVoicemeeterDefaultsFromButton(); };

            _systemStatus = new Label
            {
                AutoSize = false,
                Width = 690,
                Height = 28,
                TextAlign = ContentAlignment.MiddleLeft
            };

            topPanel.Controls.Add(_systemRefreshButton);
            topPanel.Controls.Add(_systemVmDefaultsButton);
            topPanel.Controls.Add(_systemStatus);

            _systemSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical
            };
            _systemSplit.Resize += delegate { SetSplitEven(_systemSplit); };

            _systemOutputList = CreateListView();
            _systemOutputList.Columns.Add("", 34);
            _systemOutputList.Columns.Add("", 235);
            _systemOutputList.Columns.Add("", 150);
            _systemOutputList.Columns.Add("", 80);
            _systemOutputList.ItemSelectionChanged += delegate(object sender, ListViewItemSelectionChangedEventArgs e) { ApplySystemDeviceFromSelection(_systemOutputList, e); };
            _systemOutputList.KeyDown += delegate(object sender, KeyEventArgs e) { ApplySystemDeviceFromKey(_systemOutputList, e); };

            _systemInputList = CreateListView();
            _systemInputList.Columns.Add("", 34);
            _systemInputList.Columns.Add("", 235);
            _systemInputList.Columns.Add("", 150);
            _systemInputList.Columns.Add("", 80);
            _systemInputList.ItemSelectionChanged += delegate(object sender, ListViewItemSelectionChangedEventArgs e) { ApplySystemDeviceFromSelection(_systemInputList, e); };
            _systemInputList.KeyDown += delegate(object sender, KeyEventArgs e) { ApplySystemDeviceFromKey(_systemInputList, e); };

            _systemInputGroup = WrapInGroup(_systemInputList);
            _systemOutputGroup = WrapInGroup(_systemOutputList);
            _systemSplit.Panel1.Controls.Add(_systemInputGroup);
            _systemSplit.Panel2.Controls.Add(_systemOutputGroup);

            _systemVolumeSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical
            };
            _systemVolumeSplit.Resize += delegate { SetSplitEven(_systemVolumeSplit); };

            _systemInputVolumePanel = CreateVolumePanel();
            _systemOutputVolumePanel = CreateVolumePanel();
            _systemInputVolumeGroup = WrapInGroup(_systemInputVolumePanel);
            _systemOutputVolumeGroup = WrapInGroup(_systemOutputVolumePanel);
            _systemVolumeSplit.Panel1.Controls.Add(_systemInputVolumeGroup);
            _systemVolumeSplit.Panel2.Controls.Add(_systemOutputVolumeGroup);

            var contentPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2
            };
            contentPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            contentPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            contentPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 220F));
            contentPanel.Controls.Add(_systemSplit, 0, 0);
            contentPanel.Controls.Add(_systemVolumeSplit, 0, 1);

            page.Controls.Add(contentPanel);
            page.Controls.Add(topPanel);
            return page;
        }

        private TabPage BuildVoicemeeterTab()
        {
            var page = new TabPage();

            var topPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 54,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(10, 9, 10, 8),
                WrapContents = false
            };

            _vmPhysicalOnly = new RadioButton
            {
                AutoSize = true,
                Checked = true,
                Padding = new Padding(0, 5, 0, 0)
            };
            _vmPhysicalOnly.CheckedChanged += delegate { if (!_loadingVoicemeeter) RefreshVoicemeeterDevices(); };

            _vmShowAll = new RadioButton
            {
                AutoSize = true,
                Padding = new Padding(14, 5, 0, 0)
            };
            _vmShowAll.CheckedChanged += delegate { if (!_loadingVoicemeeter) RefreshVoicemeeterDevices(); };

            _vmRefreshButton = new Button
            {
                Width = 116,
                Height = 28
            };
            _vmRefreshButton.Click += delegate { RefreshVoicemeeterDevices(); };

            _vmRestartButton = new Button
            {
                Width = 150,
                Height = 28
            };
            _vmRestartButton.Click += delegate { RestartVoicemeeterEngine(); };

            _vmStatus = new Label
            {
                AutoSize = false,
                Width = 470,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft
            };

            topPanel.Controls.Add(_vmPhysicalOnly);
            topPanel.Controls.Add(_vmShowAll);
            topPanel.Controls.Add(_vmRefreshButton);
            topPanel.Controls.Add(_vmRestartButton);
            topPanel.Controls.Add(_vmStatus);

            _vmSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical
            };
            _vmSplit.Resize += delegate { SetSplitEven(_vmSplit); };

            _vmInputList = CreateListView();
            _vmInputList.Columns.Add("", 34);
            _vmInputList.Columns.Add("", 64);
            _vmInputList.Columns.Add("", 300);
            _vmInputList.Columns.Add("", 150);
            _vmInputList.ItemSelectionChanged += delegate(object sender, ListViewItemSelectionChangedEventArgs e) { ApplyVoicemeeterDeviceFromSelection(_vmInputList, e); };
            _vmInputList.KeyDown += delegate(object sender, KeyEventArgs e) { ApplyVoicemeeterDeviceFromKey(_vmInputList, e); };

            _vmOutputList = CreateListView();
            _vmOutputList.Columns.Add("", 34);
            _vmOutputList.Columns.Add("", 64);
            _vmOutputList.Columns.Add("", 300);
            _vmOutputList.Columns.Add("", 150);
            _vmOutputList.ItemSelectionChanged += delegate(object sender, ListViewItemSelectionChangedEventArgs e) { ApplyVoicemeeterDeviceFromSelection(_vmOutputList, e); };
            _vmOutputList.KeyDown += delegate(object sender, KeyEventArgs e) { ApplyVoicemeeterDeviceFromKey(_vmOutputList, e); };

            _vmInputGroup = WrapInGroup(_vmInputList);
            _vmOutputGroup = WrapInGroup(_vmOutputList);
            _vmSplit.Panel1.Controls.Add(_vmInputGroup);
            _vmSplit.Panel2.Controls.Add(_vmOutputGroup);

            page.Controls.Add(_vmSplit);
            page.Controls.Add(topPanel);
            return page;
        }

        private TabPage BuildLanguageTab()
        {
            var page = new TabPage();
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 150,
                ColumnCount = 2,
                RowCount = 3,
                Padding = new Padding(22, 24, 22, 12)
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));

            _languageTitle = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };

            _languageCombo = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 220
            };
            _languageCombo.SelectedIndexChanged += delegate { ChangeLanguageFromCombo(); };

            _languageDescription = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };

            panel.Controls.Add(_languageTitle, 0, 0);
            panel.Controls.Add(_languageCombo, 1, 0);
            panel.Controls.Add(_languageDescription, 1, 1);
            page.Controls.Add(panel);
            return page;
        }

        private static ListView CreateListView()
        {
            return new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false,
                MultiSelect = false,
                BorderStyle = BorderStyle.None
            };
        }

        private static Panel CreateVolumePanel()
        {
            return new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = SystemColors.Window
            };
        }

        private static void SetSplitEven(SplitContainer split)
        {
            if (split == null || split.ClientSize.Width <= 0)
            {
                return;
            }

            int distance = Math.Max(split.Panel1MinSize, (split.ClientSize.Width - split.SplitterWidth) / 2);
            int maxDistance = split.ClientSize.Width - split.SplitterWidth - split.Panel2MinSize;
            if (maxDistance > 0)
            {
                distance = Math.Min(distance, maxDistance);
            }

            if (distance > 0 && split.SplitterDistance != distance)
            {
                split.SplitterDistance = distance;
            }
        }

        private static GroupBox WrapInGroup(Control child)
        {
            var group = new GroupBox
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10, 20, 10, 10)
            };
            group.Controls.Add(child);
            return group;
        }

        private string L(string key)
        {
            return Localizer.T(_languageCode, key);
        }

        private string LF(string key, params object[] args)
        {
            return string.Format(L(key), args);
        }

        private void ApplyLanguage()
        {
            if (_systemPage != null)
            {
                _systemPage.Text = L("tab.system");
            }

            if (_voicemeeterPage != null)
            {
                _voicemeeterPage.Text = L("tab.voicemeeter");
            }

            if (_languagePage != null)
            {
                _languagePage.Text = GetLanguageTabCaption();
            }

            if (_systemRefreshButton != null)
            {
                _systemRefreshButton.Text = L("button.refresh");
            }

            if (_systemVmDefaultsButton != null)
            {
                _systemVmDefaultsButton.Text = L("button.vmDefaults");
            }

            if (_vmRefreshButton != null)
            {
                _vmRefreshButton.Text = L("button.refresh");
            }

            if (_vmRestartButton != null)
            {
                _vmRestartButton.Text = L("button.restart");
            }

            if (_vmPhysicalOnly != null)
            {
                _vmPhysicalOnly.Text = L("radio.physicalOnly");
            }

            if (_vmShowAll != null)
            {
                _vmShowAll.Text = L("radio.showAll");
            }

            if (_systemInputGroup != null)
            {
                _systemInputGroup.Text = L("group.input");
            }

            if (_systemOutputGroup != null)
            {
                _systemOutputGroup.Text = L("group.output");
            }

            if (_systemInputVolumeGroup != null)
            {
                _systemInputVolumeGroup.Text = L("group.volumeInput");
            }

            if (_systemOutputVolumeGroup != null)
            {
                _systemOutputVolumeGroup.Text = L("group.volumeOutput");
            }

            if (_vmInputGroup != null)
            {
                _vmInputGroup.Text = L("group.vmInput");
            }

            if (_vmOutputGroup != null)
            {
                _vmOutputGroup.Text = L("group.vmOutput");
            }

            SetSystemColumns(_systemInputList, L("column.input"));
            SetSystemColumns(_systemOutputList, L("column.output"));
            SetVoicemeeterColumns(_vmInputList, L("column.input"));
            SetVoicemeeterColumns(_vmOutputList, L("column.output"));

            if (_languageTitle != null)
            {
                _languageTitle.Text = L("language.title");
            }

            if (_languageDescription != null)
            {
                _languageDescription.Text = L("language.description");
            }

            PopulateLanguageCombo();
            RebuildLanguageMenu();
            RefreshSystemVolumeControls();
        }

        private string GetLanguageTabCaption()
        {
            if (string.Equals(_languageCode, "zh", StringComparison.OrdinalIgnoreCase))
            {
                return "Language ▼";
            }

            if (string.Equals(_languageCode, "en", StringComparison.OrdinalIgnoreCase))
            {
                return "语言 ▼";
            }

            return L("tab.language") + " ▼";
        }

        private void TabsSelecting(object sender, TabControlCancelEventArgs e)
        {
            if (e.TabPage != _languagePage)
            {
                return;
            }

            e.Cancel = true;
            ShowLanguageMenu();
        }

        private void ShowLanguageMenu()
        {
            if (_tabs == null || _languagePage == null || _showingLanguageMenu)
            {
                return;
            }

            RebuildLanguageMenu();
            int index = _tabs.TabPages.IndexOf(_languagePage);
            if (index < 0)
            {
                return;
            }

            Rectangle rect = _tabs.GetTabRect(index);
            _showingLanguageMenu = true;
            _languageMenu.Closed += LanguageMenuClosedOnce;
            _languageMenu.Show(_tabs, new Point(rect.Left, rect.Bottom + 1));
        }

        private void LanguageMenuClosedOnce(object sender, ToolStripDropDownClosedEventArgs e)
        {
            if (_languageMenu != null)
            {
                _languageMenu.Closed -= LanguageMenuClosedOnce;
            }

            _showingLanguageMenu = false;
        }

        private void RebuildLanguageMenu()
        {
            if (_languageMenu == null)
            {
                _languageMenu = new ContextMenuStrip();
            }

            _languageMenu.Items.Clear();
            foreach (LanguageOption option in Localizer.LanguageOptions)
            {
                var item = new ToolStripMenuItem(option.DisplayName)
                {
                    Tag = option,
                    Checked = string.Equals(option.Code, _languageCode, StringComparison.OrdinalIgnoreCase)
                };
                item.Click += delegate(object sender, EventArgs e)
                {
                    var menuItem = sender as ToolStripMenuItem;
                    var languageOption = menuItem == null ? null : menuItem.Tag as LanguageOption;
                    if (languageOption != null)
                    {
                        SetLanguage(languageOption);
                    }
                };

                _languageMenu.Items.Add(item);
            }
        }

        private void PopulateLanguageCombo()
        {
            if (_languageCombo == null)
            {
                return;
            }

            _updatingLanguage = true;
            try
            {
                if (_languageCombo.Items.Count == 0)
                {
                    foreach (LanguageOption option in Localizer.LanguageOptions)
                    {
                        _languageCombo.Items.Add(option);
                    }
                }

                LanguageOption selected = Localizer.GetOption(_languageCode);
                foreach (object item in _languageCombo.Items)
                {
                    LanguageOption option = item as LanguageOption;
                    if (option != null && string.Equals(option.Code, selected.Code, StringComparison.OrdinalIgnoreCase))
                    {
                        _languageCombo.SelectedItem = item;
                        break;
                    }
                }
            }
            finally
            {
                _updatingLanguage = false;
            }
        }

        private void ChangeLanguageFromCombo()
        {
            if (_updatingLanguage || _languageCombo == null)
            {
                return;
            }

            LanguageOption option = _languageCombo.SelectedItem as LanguageOption;
            if (option == null)
            {
                return;
            }

            SetLanguage(option);
        }

        private void SetLanguage(LanguageOption option)
        {
            if (option == null)
            {
                return;
            }

            _languageCode = option.Code;
            _settings.LanguageCode = option.Code;
            ApplyLanguage();
            string status = LF("language.saved", option.DisplayName);
            if (_tabs != null && _tabs.SelectedTab == _systemPage && _systemStatus != null)
            {
                _systemStatus.Text = status;
            }
            else if (_vmStatus != null)
            {
                _vmStatus.Text = status;
            }
        }

        private void SetSystemColumns(ListView list, string deviceHeader)
        {
            if (list == null || list.Columns.Count < 4)
            {
                return;
            }

            list.Columns[1].Text = deviceHeader;
            list.Columns[2].Text = L("column.driverInterface");
            list.Columns[3].Text = L("column.status");
        }

        private void SetVoicemeeterColumns(ListView list, string deviceHeader)
        {
            if (list == null || list.Columns.Count < 4)
            {
                return;
            }

            list.Columns[1].Text = L("column.driver");
            list.Columns[2].Text = deviceHeader;
            list.Columns[3].Text = L("column.hardwareId");
        }

        private void RefreshSystemDevices()
        {
            if (_loadingSystem)
            {
                return;
            }

            try
            {
                _loadingSystem = true;
                _systemOutputs = MarkSavedDefaultIfNeeded(_audioController.GetDevices(SoundDeviceFlow.Output), _settings.SystemOutputDeviceId);
                _systemInputs = MarkSavedDefaultIfNeeded(_audioController.GetDevices(SoundDeviceFlow.Input), _settings.SystemInputDeviceId);

                FillSystemList(_systemOutputList, _systemOutputs);
                FillSystemList(_systemInputList, _systemInputs);
                RefreshSystemVolumeControls();
                _systemStatus.Text = LF("status.systemRead", _systemOutputs.Count, _systemInputs.Count);
            }
            catch (Exception ex)
            {
                _systemStatus.Text = LF("status.systemReadFailed", ex.Message);
            }
            finally
            {
                _loadingSystem = false;
            }
        }

        private void FillSystemList(ListView list, IList<AudioDevice> devices)
        {
            list.BeginUpdate();
            list.Items.Clear();
            foreach (AudioDevice device in devices)
            {
                var item = new ListViewItem(device.IsDefault ? "✓" : string.Empty);
                item.SubItems.Add(device.Name);
                item.SubItems.Add(device.InterfaceName);
                item.SubItems.Add(device.StateText);
                item.Tag = device;

                if (device.IsDefault)
                {
                    item.BackColor = Color.FromArgb(226, 240, 255);
                    item.Font = new Font(list.Font, FontStyle.Bold);
                }

                list.Items.Add(item);
            }
            list.EndUpdate();
        }

        private void RefreshSystemVolumeControls()
        {
            if (_systemInputVolumePanel == null || _systemOutputVolumePanel == null)
            {
                return;
            }

            FillVolumePanel(_systemInputVolumePanel, GetSystemVolumeDevices(_systemInputs));
            FillVolumePanel(_systemOutputVolumePanel, GetSystemVolumeDevices(_systemOutputs));
        }

        private void FillVolumePanel(Panel panel, IList<AudioDevice> devices)
        {
            panel.SuspendLayout();
            try
            {
                panel.Controls.Clear();

                var table = new TableLayoutPanel
                {
                    Dock = DockStyle.Top,
                    AutoSize = true,
                    ColumnCount = 3,
                    GrowStyle = TableLayoutPanelGrowStyle.AddRows,
                    Padding = new Padding(4, 4, 4, 4)
                };
                table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
                table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190F));
                table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 58F));

                panel.Controls.Add(table);

                if (devices.Count == 0)
                {
                    table.RowStyles.Add(new RowStyle(SizeType.Absolute, 32F));
                    var empty = new Label
                    {
                        Dock = DockStyle.Fill,
                        TextAlign = ContentAlignment.MiddleLeft,
                        Text = L("volume.none"),
                        ForeColor = SystemColors.GrayText
                    };
                    table.Controls.Add(empty, 0, 0);
                    table.SetColumnSpan(empty, 3);
                    return;
                }

                int row = 0;
                foreach (AudioDevice device in devices)
                {
                    int volume;
                    bool hasVolume = _audioController.TryGetEndpointVolumePercent(device, out volume);
                    AudioDevice rowDevice = device;

                    table.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));

                    var nameLabel = new Label
                    {
                        Dock = DockStyle.Fill,
                        AutoEllipsis = true,
                        TextAlign = ContentAlignment.MiddleLeft,
                        Text = FormatVolumeDeviceName(device) + (hasVolume ? string.Empty : " (" + L("volume.unavailable") + ")"),
                        Margin = new Padding(4, 3, 8, 3)
                    };

                    var slider = new TrackBar
                    {
                        Dock = DockStyle.Fill,
                        AutoSize = false,
                        Height = 28,
                        Minimum = 0,
                        Maximum = 100,
                        TickStyle = TickStyle.None,
                        SmallChange = 1,
                        LargeChange = 5,
                        Value = hasVolume ? volume : 0,
                        Enabled = hasVolume,
                        Margin = new Padding(0, 2, 8, 2)
                    };

                    var number = new NumericUpDown
                    {
                        Minimum = 0,
                        Maximum = 100,
                        Increment = 1,
                        Width = 54,
                        TextAlign = HorizontalAlignment.Right,
                        Value = hasVolume ? volume : 0,
                        Enabled = hasVolume,
                        Margin = new Padding(0, 5, 0, 3)
                    };

                    TrackBar rowSlider = slider;
                    NumericUpDown rowNumber = number;
                    slider.ValueChanged += delegate
                    {
                        if (_updatingVolumeControls)
                        {
                            return;
                        }

                        int newValue = rowSlider.Value;
                        _updatingVolumeControls = true;
                        try
                        {
                            rowNumber.Value = newValue;
                        }
                        finally
                        {
                            _updatingVolumeControls = false;
                        }

                        ApplySystemVolume(rowDevice, newValue);
                    };

                    number.ValueChanged += delegate
                    {
                        if (_updatingVolumeControls)
                        {
                            return;
                        }

                        int newValue = (int)rowNumber.Value;
                        _updatingVolumeControls = true;
                        try
                        {
                            rowSlider.Value = newValue;
                        }
                        finally
                        {
                            _updatingVolumeControls = false;
                        }

                        ApplySystemVolume(rowDevice, newValue);
                    };

                    table.Controls.Add(nameLabel, 0, row);
                    table.Controls.Add(slider, 1, row);
                    table.Controls.Add(number, 2, row);
                    row++;
                }
            }
            finally
            {
                panel.ResumeLayout();
            }
        }

        private IList<AudioDevice> GetSystemVolumeDevices(IList<AudioDevice> devices)
        {
            var result = new List<AudioDevice>();
            foreach (AudioDevice device in devices)
            {
                if (IsPrimaryVoicemeeterVolumeDevice(device) || IsPhysicalVolumeDevice(device))
                {
                    result.Add(device);
                }
            }

            return result;
        }

        private static bool IsPhysicalVolumeDevice(AudioDevice device)
        {
            return device != null && !DeviceFilters.LooksVirtual(device.FullName);
        }

        private static bool IsPrimaryVoicemeeterVolumeDevice(AudioDevice device)
        {
            if (device == null)
            {
                return false;
            }

            string name = NormalizeText(device.Name);
            string value = NormalizeText(device.FullName);
            if (value.Contains("aux"))
            {
                return false;
            }

            if (device.Flow == SoundDeviceFlow.Output)
            {
                return name.Contains("voicemeeter input") || name.Contains("voicemeter input");
            }

            return name.Contains("voicemeeter output") || name.Contains("voicemeter output") ||
                name.Contains("voicemeeter b1") || name.Contains("voicemeter b1");
        }

        private static string NormalizeText(string value)
        {
            return (value ?? string.Empty).ToLowerInvariant();
        }

        private static string FormatVolumeDeviceName(AudioDevice device)
        {
            if (device == null)
            {
                return string.Empty;
            }

            if (string.IsNullOrWhiteSpace(device.InterfaceName) || string.Equals(device.Name, device.InterfaceName, StringComparison.OrdinalIgnoreCase))
            {
                return device.Name;
            }

            return device.Name + " (" + device.InterfaceName + ")";
        }

        private void ApplySystemVolume(AudioDevice device, int volumePercent)
        {
            try
            {
                _audioController.SetEndpointVolumePercent(device, volumePercent);
                if (_systemStatus != null)
                {
                    _systemStatus.Text = LF("status.volumeSet", device.Name, volumePercent);
                }
            }
            catch (Exception ex)
            {
                if (_systemStatus != null)
                {
                    _systemStatus.Text = LF("status.volumeSetFailed", device.Name, ex.Message);
                }
            }
        }

        private static IList<AudioDevice> MarkSavedDefaultIfNeeded(IList<AudioDevice> devices, string savedDeviceId)
        {
            if (string.IsNullOrWhiteSpace(savedDeviceId))
            {
                return devices;
            }

            foreach (AudioDevice device in devices)
            {
                if (device.IsDefault)
                {
                    return devices;
                }
            }

            var marked = new List<AudioDevice>();
            foreach (AudioDevice device in devices)
            {
                bool isSaved = string.Equals(device.Id, savedDeviceId, StringComparison.OrdinalIgnoreCase);
                marked.Add(new AudioDevice(device.Id, device.Name, device.InterfaceName, device.Flow, device.State, isSaved));
            }

            return marked;
        }

        private void ApplyStartupDefaultsIfNeeded()
        {
            if (_settings.HasUserSystemSelection)
            {
                return;
            }

            try
            {
                AudioDevice output;
                AudioDevice input;
                bool changed = _audioController.TryApplyVoicemeeterDefaults(_systemOutputs, _systemInputs, out output, out input);
                string message = BuildVoicemeeterDefaultMessage(output, input);
                if (changed)
                {
                    _systemStatus.Text = LF("status.startupDefaultsApplied", message);
                    Thread.Sleep(250);
                }
                else
                {
                    _systemStatus.Text = LF("status.startupDefaultsMissing", message);
                }
            }
            catch (Exception ex)
            {
                _systemStatus.Text = LF("status.vmDefaultsFailed", ex.Message);
            }
        }

        private void ApplyVoicemeeterDefaultsFromButton()
        {
            try
            {
                AudioDevice output;
                AudioDevice input;
                _systemOutputs = _audioController.GetDevices(SoundDeviceFlow.Output);
                _systemInputs = _audioController.GetDevices(SoundDeviceFlow.Input);
                _audioController.TryApplyVoicemeeterDefaults(_systemOutputs, _systemInputs, out output, out input);
                string message = BuildVoicemeeterDefaultMessage(output, input);
                _systemStatus.Text = message;
                Thread.Sleep(250);
                RefreshSystemDevices();
            }
            catch (Exception ex)
            {
                _systemStatus.Text = LF("status.vmDefaultsFailed", ex.Message);
            }
        }

        private string BuildVoicemeeterDefaultMessage(AudioDevice output, AudioDevice input)
        {
            string outputPart = output != null ? LF("status.vmDefaultOutput", output.Name) : L("status.vmDefaultOutputMissing");
            string inputPart = input != null ? LF("status.vmDefaultInput", input.Name) : L("status.vmDefaultInputMissing");
            return outputPart + "; " + inputPart;
        }

        private void ApplySystemDeviceFromSelection(ListView list, ListViewItemSelectionChangedEventArgs e)
        {
            if (_loadingSystem || !e.IsSelected)
            {
                return;
            }

            ApplySystemDevice(list, e.Item);
        }

        private void ApplySystemDeviceFromKey(ListView list, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && list.SelectedItems.Count == 1)
            {
                ApplySystemDevice(list, list.SelectedItems[0]);
                e.Handled = true;
            }
        }

        private void ApplySystemDevice(ListView list, ListViewItem item)
        {
            AudioDevice device = item.Tag as AudioDevice;
            if (device == null)
            {
                return;
            }

            try
            {
                _audioController.SetDefaultDevice(device);
                if (ReferenceEquals(list, _systemOutputList))
                {
                    _settings.SystemOutputDeviceId = device.Id;
                    SetSystemListCheck(_systemOutputList, device.Id);
                    _systemStatus.Text = LF("status.defaultOutputSet", device.Name);
                }
                else
                {
                    _settings.SystemInputDeviceId = device.Id;
                    SetSystemListCheck(_systemInputList, device.Id);
                    _systemStatus.Text = LF("status.defaultInputSet", device.Name);
                }

                Thread.Sleep(250);
                RefreshSystemDevices();
            }
            catch (Exception ex)
            {
                _systemStatus.Text = LF("status.systemSetFailed", ex.Message);
                MessageBox.Show(this, ex.Message, L("message.systemSetFailed"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                RefreshSystemDevices();
            }
        }

        private static void SetSystemListCheck(ListView list, string selectedDeviceId)
        {
            list.BeginUpdate();
            foreach (ListViewItem row in list.Items)
            {
                AudioDevice rowDevice = row.Tag as AudioDevice;
                bool selected = rowDevice != null &&
                    string.Equals(rowDevice.Id, selectedDeviceId, StringComparison.OrdinalIgnoreCase);
                row.Text = selected ? "✓" : string.Empty;
                row.BackColor = selected ? Color.FromArgb(226, 240, 255) : SystemColors.Window;
                row.Font = selected ? new Font(list.Font, FontStyle.Bold) : list.Font;
            }
            list.EndUpdate();
        }

        private void RefreshVoicemeeterDevices()
        {
            if (_loadingVoicemeeter)
            {
                return;
            }

            try
            {
                _loadingVoicemeeter = true;
                string typeText = _voicemeeter.EnsureConnected();
                string currentInput = _voicemeeter.GetCurrentHardwareInput1();
                string currentOutput = _voicemeeter.GetCurrentA1Output();

                IList<VoicemeeterDevice> inputs = _voicemeeter.GetInputDevices();
                IList<VoicemeeterDevice> outputs = _voicemeeter.GetOutputDevices();

                FillVoicemeeterList(_vmInputList, inputs, currentInput, _settings.VoicemeeterInputDeviceKey);
                FillVoicemeeterList(_vmOutputList, outputs, currentOutput, _settings.VoicemeeterOutputDeviceKey);

                string mode = _vmShowAll.Checked ? L("status.modeAll") : L("status.modePhysical");
                _vmStatus.Text = LF("status.vmConnected", typeText, mode, _voicemeeter.DllPath);
            }
            catch (Exception ex)
            {
                _vmStatus.Text = LF("status.vmReadFailed", ex.Message);
                _vmInputList.Items.Clear();
                _vmOutputList.Items.Clear();
            }
            finally
            {
                _loadingVoicemeeter = false;
            }
        }

        private void FillVoicemeeterList(ListView list, IList<VoicemeeterDevice> devices, string currentDeviceName, string savedDeviceKey)
        {
            bool showAll = _vmShowAll.Checked;
            string selectedKey = ChooseVoicemeeterSelectedKey(devices, currentDeviceName, savedDeviceKey, showAll);

            list.BeginUpdate();
            list.Items.Clear();
            foreach (VoicemeeterDevice device in devices)
            {
                if (!showAll && device.LooksVirtual)
                {
                    continue;
                }

                bool selected = string.Equals(device.Key, selectedKey, StringComparison.OrdinalIgnoreCase);
                var item = new ListViewItem(selected ? "✓" : string.Empty);
                item.SubItems.Add(device.Driver);
                item.SubItems.Add(device.Name);
                item.SubItems.Add(device.HardwareId);
                item.Tag = device;

                if (selected)
                {
                    item.BackColor = Color.FromArgb(226, 240, 255);
                    item.Font = new Font(list.Font, FontStyle.Bold);
                }

                list.Items.Add(item);
            }
            list.EndUpdate();
        }

        private static string ChooseVoicemeeterSelectedKey(IList<VoicemeeterDevice> devices, string currentDeviceName, string savedDeviceKey, bool showAll)
        {
            if (string.IsNullOrWhiteSpace(currentDeviceName))
            {
                return string.Empty;
            }

            VoicemeeterDevice firstVisibleMatch = null;
            VoicemeeterDevice preferredMatch = null;

            foreach (VoicemeeterDevice device in devices)
            {
                if (!showAll && device.LooksVirtual)
                {
                    continue;
                }

                if (!string.Equals(device.Name, currentDeviceName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(savedDeviceKey) &&
                    string.Equals(device.Key, savedDeviceKey, StringComparison.OrdinalIgnoreCase))
                {
                    return device.Key;
                }

                if (firstVisibleMatch == null)
                {
                    firstVisibleMatch = device;
                }

                if (preferredMatch == null || DriverRank(device.Driver) < DriverRank(preferredMatch.Driver))
                {
                    preferredMatch = device;
                }
            }

            return preferredMatch != null ? preferredMatch.Key : (firstVisibleMatch != null ? firstVisibleMatch.Key : string.Empty);
        }

        private static int DriverRank(string driver)
        {
            switch (driver)
            {
                case "WDM":
                    return 0;
                case "KS":
                    return 1;
                case "MME":
                    return 2;
                case "ASIO":
                    return 3;
                default:
                    return 99;
            }
        }

        private void ApplyVoicemeeterDeviceFromSelection(ListView list, ListViewItemSelectionChangedEventArgs e)
        {
            if (_loadingVoicemeeter || !e.IsSelected)
            {
                return;
            }

            ApplyVoicemeeterDevice(list, e.Item);
        }

        private void ApplyVoicemeeterDeviceFromKey(ListView list, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && list.SelectedItems.Count == 1)
            {
                ApplyVoicemeeterDevice(list, list.SelectedItems[0]);
                e.Handled = true;
            }
        }

        private void ApplyVoicemeeterDevice(ListView list, ListViewItem item)
        {
            VoicemeeterDevice device = item.Tag as VoicemeeterDevice;
            if (device == null)
            {
                return;
            }

            try
            {
                if (ReferenceEquals(list, _vmInputList))
                {
                    _voicemeeter.SetHardwareInput1(device);
                    _settings.VoicemeeterInputDeviceKey = device.Key;
                    SetVoicemeeterListCheck(_vmInputList, device.Key);
                    _vmStatus.Text = LF("status.vmInputSet", device.Driver, device.Name);
                }
                else
                {
                    _voicemeeter.SetA1Output(device);
                    _settings.VoicemeeterOutputDeviceKey = device.Key;
                    SetVoicemeeterListCheck(_vmOutputList, device.Key);
                    _vmStatus.Text = LF("status.vmOutputSet", device.Driver, device.Name);
                }

                Thread.Sleep(300);
                RefreshVoicemeeterDevices();
            }
            catch (Exception ex)
            {
                _vmStatus.Text = LF("status.vmSetFailed", ex.Message);
                MessageBox.Show(this, ex.Message, L("message.vmSetFailed"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                RefreshVoicemeeterDevices();
            }
        }

        private static void SetVoicemeeterListCheck(ListView list, string selectedDeviceKey)
        {
            list.BeginUpdate();
            foreach (ListViewItem row in list.Items)
            {
                VoicemeeterDevice rowDevice = row.Tag as VoicemeeterDevice;
                bool selected = rowDevice != null &&
                    string.Equals(rowDevice.Key, selectedDeviceKey, StringComparison.OrdinalIgnoreCase);
                row.Text = selected ? "✓" : string.Empty;
                row.BackColor = selected ? Color.FromArgb(226, 240, 255) : SystemColors.Window;
                row.Font = selected ? new Font(list.Font, FontStyle.Bold) : list.Font;
            }
            list.EndUpdate();
        }

        private void RestartVoicemeeterEngine()
        {
            try
            {
                _voicemeeter.RestartAudioEngine();
                _vmStatus.Text = L("status.vmRestarted");
            }
            catch (Exception ex)
            {
                _vmStatus.Text = LF("status.vmRestartFailed", ex.Message);
            }
        }
    }
}

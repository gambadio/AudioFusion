using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using NAudio.CoreAudioApi;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

namespace AudioFusion
{
    public partial class MainWindow : Window
    {
        // Core components
        private MMDeviceEnumerator _deviceEnumerator;
        private List<MMDevice> _outputDevices = new List<MMDevice>();
        private List<MMDevice> _inputDevices = new List<MMDevice>();
        private MMDevice _defaultOutputDevice;
        private MMDevice _defaultInputDevice;
        
        // Output Fusion Components
        private WasapiLoopbackCapture _audioSourceCapture;
        private WasapiOut _secondaryHeadsetPlayer;
        private BufferedWaveProvider _outputFusionBuffer;
        private bool _isOutputFusionRunning = false;

        // Microphone Switching
        private bool _isMicSwitchingEnabled = false;
        private MMDevice _micOne;
        private MMDevice _micTwo;

        public MainWindow()
        {
            InitializeComponent();
            
            // Initialize slider to disabled state
            MicSelectionSlider.IsEnabled = false;
            
            LoadAudioDevices();
        }

        private void LoadAudioDevices()
        {
            try
            {
                // Clear existing devices
                _outputDevices.Clear();
                _inputDevices.Clear();
                
                // Create enumerator if needed
                if (_deviceEnumerator == null)
                    _deviceEnumerator = new MMDeviceEnumerator();
                
                // Save selections before clearing
                string selectedSecondaryHeadset = SecondaryHeadsetComboBox.SelectedItem?.ToString();
                string selectedMicOne = MicOneComboBox.SelectedItem?.ToString();
                string selectedMicTwo = MicTwoComboBox.SelectedItem?.ToString();
                
                // Get default output device
                _defaultOutputDevice = _deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                DefaultOutputDeviceText.Text = $"{_defaultOutputDevice.FriendlyName} (System Default)";
                
                // Get default input device - try both roles
                try 
                {
                    // Try communications role first
                    _defaultInputDevice = _deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications);
                }
                catch 
                {
                    try
                    {
                        // Fall back to multimedia role
                        _defaultInputDevice = _deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Multimedia);
                    }
                    catch
                    {
                        DefaultMicrophoneText.Text = "No default microphone found";
                        _defaultInputDevice = null;
                    }
                }

                if (_defaultInputDevice != null)
                {
                    DefaultMicrophoneText.Text = $"{_defaultInputDevice.FriendlyName} (System Default)";
                }
                
                // Get output devices
                var outputs = _deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active).ToList();
                _outputDevices.AddRange(outputs);
                
                // Get input devices
                var inputs = _deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active).ToList();
                _inputDevices.AddRange(inputs);
                
                // Clear and repopulate combo boxes
                PopulateComboBoxes();
                
                // Restore selections or select defaults, ensuring we pick the actual default device first
                if (_defaultInputDevice != null)
                {
                    // Find index of default device
                    int defaultIndex = -1;
                    for (int i = 0; i < MicOneComboBox.Items.Count; i++)
                    {
                        if (MicOneComboBox.Items[i].ToString() == _defaultInputDevice.FriendlyName)
                        {
                            defaultIndex = i;
                            break;
                        }
                    }

                    if (defaultIndex != -1)
                    {
                        MicOneComboBox.SelectedIndex = defaultIndex;
                    }
                }

                RestoreSelection(SecondaryHeadsetComboBox, selectedSecondaryHeadset);
                RestoreSelection(MicTwoComboBox, selectedMicTwo);
                
                // Ensure different microphones selected
                if (MicOneComboBox.SelectedIndex == MicTwoComboBox.SelectedIndex && MicTwoComboBox.Items.Count > 1)
                {
                    MicTwoComboBox.SelectedIndex = (MicOneComboBox.SelectedIndex + 1) % MicTwoComboBox.Items.Count;
                }
                
                // Update mic name labels
                UpdateMicNameLabels();
                
                // Update slider position based on default mic
                UpdateSliderPosition();
                
                StatusTextBlock.Text = "Audio devices loaded. Ready.";
                RefreshDevicesButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading audio devices: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusTextBlock.Text = "Error loading audio devices.";
            }
        }

        private void PopulateComboBoxes()
        {
            // SECONDARY HEADSET
            SecondaryHeadsetComboBox.Items.Clear();
            foreach (var device in _outputDevices)
            {
                SecondaryHeadsetComboBox.Items.Add(device.FriendlyName);
            }
            
            // MICROPHONES
            MicOneComboBox.Items.Clear();
            MicTwoComboBox.Items.Clear();
            foreach (var device in _inputDevices)
            {
                MicOneComboBox.Items.Add(device.FriendlyName);
                MicTwoComboBox.Items.Add(device.FriendlyName);
            }
        }
        
        private void UpdateSliderPosition()
        {
            // If mic switching is disabled, don't update the slider
            if (!_isMicSwitchingEnabled) return;
            
            // Update slider based on which mic is currently the default
            if (_defaultInputDevice != null)
            {
                if (_micOne != null && _defaultInputDevice.ID == _micOne.ID)
                {
                    MicSelectionSlider.Value = 0;
                }
                else if (_micTwo != null && _defaultInputDevice.ID == _micTwo.ID)
                {
                    MicSelectionSlider.Value = 1;
                }
            }
        }
        
        private void UpdateMicNameLabels()
        {
            MicOneName.Text = MicOneComboBox.SelectedItem?.ToString() ?? "None";
            MicTwoName.Text = MicTwoComboBox.SelectedItem?.ToString() ?? "None";
            
            // Also update references to the actual devices
            _micOne = FindMicDevice(MicOneComboBox.SelectedItem?.ToString());
            _micTwo = FindMicDevice(MicTwoComboBox.SelectedItem?.ToString());
        }
        
        private MMDevice FindMicDevice(string friendlyName)
        {
            if (string.IsNullOrEmpty(friendlyName))
                return null;
                
            return _inputDevices.FirstOrDefault(d => d.FriendlyName == friendlyName);
        }
        
        private void RestoreSelection(ComboBox comboBox, string previousValue)
        {
            if (comboBox.Items.Count == 0) return;
            
            if (!string.IsNullOrEmpty(previousValue))
            {
                for (int i = 0; i < comboBox.Items.Count; i++)
                {
                    if (comboBox.Items[i].ToString() == previousValue)
                    {
                        comboBox.SelectedIndex = i;
                        return;
                    }
                }
            }
            
            // If no match or no previous, select first item
            comboBox.SelectedIndex = 0;
        }

        private void RefreshDevicesButton_Click(object sender, RoutedEventArgs e)
        {
            LoadAudioDevices();
        }

        private void OutputFusionButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isOutputFusionRunning) StopOutputFusion();
            else StartOutputFusion();
        }

        private void StartOutputFusion()
        {
            if (SecondaryHeadsetComboBox.SelectedIndex == -1)
            {
                MessageBox.Show("Please select a secondary headset device.", "Selection Missing", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            
            string selectedDeviceName = SecondaryHeadsetComboBox.SelectedItem.ToString();
            MMDevice secondaryDevice = null;
            
            // Find the selected device
            foreach (var device in _outputDevices)
            {
                if (device.FriendlyName == selectedDeviceName)
                {
                    secondaryDevice = device;
                    break;
                }
            }
            
            if (secondaryDevice == null)
            {
                MessageBox.Show("Cannot find selected output device.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            
            if (_defaultOutputDevice.ID == secondaryDevice.ID)
            {
                MessageBox.Show("The secondary headset cannot be the same as the system default output for this function.", "Device Conflict", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                _audioSourceCapture = new WasapiLoopbackCapture(_defaultOutputDevice);
                _outputFusionBuffer = new BufferedWaveProvider(_audioSourceCapture.WaveFormat)
                { BufferDuration = TimeSpan.FromMilliseconds(200), DiscardOnBufferOverflow = true };

                _audioSourceCapture.DataAvailable += (s, args) => _outputFusionBuffer?.AddSamples(args.Buffer, 0, args.BytesRecorded);
                
                _audioSourceCapture.RecordingStopped += (s, args) =>
                {
                    if (_isOutputFusionRunning) Dispatcher.Invoke(() => {
                        StopOutputFusion();
                        StatusTextBlock.Text = "Output fusion stopped unexpectedly.";
                    });
                };

                _secondaryHeadsetPlayer = new WasapiOut(secondaryDevice, AudioClientShareMode.Shared, true, 100);
                _secondaryHeadsetPlayer.Init(_outputFusionBuffer);

                _secondaryHeadsetPlayer.PlaybackStopped += (s, args) =>
                {
                    if (args.Exception != null && _isOutputFusionRunning)
                    {
                        Dispatcher.Invoke(() => {
                            StatusTextBlock.Text = $"Secondary headset playback error: {args.Exception.Message}";
                            StopOutputFusion();
                        });
                    }
                };

                _audioSourceCapture.StartRecording();
                _secondaryHeadsetPlayer.Play();

                _isOutputFusionRunning = true;
                OutputFusionButton.Content = "Stop Output Fusion";
                SecondaryHeadsetComboBox.IsEnabled = false;
                StatusTextBlock.Text = $"Output Fusion: Mirroring '{_defaultOutputDevice.FriendlyName}' to '{secondaryDevice.FriendlyName}'";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error starting output fusion: {ex.Message}", "Output Fusion Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StopOutputFusion();
            }
        }

        private void StopOutputFusion()
        {
            try
            {
                if (_audioSourceCapture != null)
                {
                    _audioSourceCapture.StopRecording();
                    _audioSourceCapture.Dispose();
                    _audioSourceCapture = null;
                }
                
                if (_secondaryHeadsetPlayer != null)
                {
                    _secondaryHeadsetPlayer.Stop();
                    _secondaryHeadsetPlayer.Dispose();
                    _secondaryHeadsetPlayer = null;
                }
                
                _outputFusionBuffer = null;
                _isOutputFusionRunning = false;
                OutputFusionButton.Content = "Start Output Fusion";
                SecondaryHeadsetComboBox.IsEnabled = true;
                
                if (!StatusTextBlock.Text.Contains("Error") && !StatusTextBlock.Text.Contains("unexpectedly"))
                    StatusTextBlock.Text = "Output fusion stopped.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error stopping output fusion: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MicSwitchingButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_isMicSwitchingEnabled)
                {
                    // Disable mic switching
                    _isMicSwitchingEnabled = false;
                    MicSwitchingButton.Content = "Enable Mic Switching";
                    MicSelectionSlider.IsEnabled = false;
                    StatusTextBlock.Text = "Microphone switching disabled.";
                }
                else
                {
                    StatusTextBlock.Text = "Initializing mic switching...";
                    
                    // Enable mic switching
                    if (MicOneComboBox.SelectedIndex == -1 || MicTwoComboBox.SelectedIndex == -1)
                    {
                        MessageBox.Show("Please select both microphones.", "Selection Missing", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    
                    StatusTextBlock.Text = "Updating device references...";
                    
                    // Update references to devices
                    UpdateMicNameLabels();
                    
                    if (_micOne == null || _micTwo == null)
                    {
                        MessageBox.Show("Could not find the selected microphones.", "Device Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                    
                    // Log device IDs for debugging
                    System.Diagnostics.Debug.WriteLine($"Mic One ID: {_micOne.ID}");
                    System.Diagnostics.Debug.WriteLine($"Mic Two ID: {_micTwo.ID}");
                    
                    // Validate device IDs
                    if (string.IsNullOrEmpty(_micOne.ID) || string.IsNullOrEmpty(_micTwo.ID))
                    {
                        MessageBox.Show("Invalid device IDs detected.", "Device Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    StatusTextBlock.Text = "Enabling mic switching...";
                    
                    try
                    {
                        _isMicSwitchingEnabled = true;
                        MicSwitchingButton.Content = "Disable Mic Switching";
                        MicSelectionSlider.IsEnabled = true;
                        
                        // Set initial slider position based on default mic
                        UpdateSliderPosition();
                        
                        // Apply initial selection based on slider position
                        ApplyMicSelection();
                        StatusTextBlock.Text = "Microphone switching enabled. Use the slider to select the active microphone.";
                    }
                    catch (Exception ex)
                    {
                        _isMicSwitchingEnabled = false;
                        MicSwitchingButton.Content = "Enable Mic Switching";
                        MicSelectionSlider.IsEnabled = false;
                        throw new Exception($"Failed to initialize microphone switching: {ex.Message}", ex);
                    }
                }
            }
            catch (Exception ex)
            {
                var error = $"Error: {ex.GetType().Name}\nMessage: {ex.Message}\nStack Trace: {ex.StackTrace}";
                if (ex.InnerException != null)
                {
                    error += $"\n\nInner Exception: {ex.InnerException.Message}\nStack Trace: {ex.InnerException.StackTrace}";
                }
                
                MessageBox.Show(error, "Error Details", MessageBoxButton.OK, MessageBoxImage.Error);
                
                _isMicSwitchingEnabled = false;
                MicSwitchingButton.Content = "Enable Mic Switching";
                MicSelectionSlider.IsEnabled = false;
                StatusTextBlock.Text = "Failed to initialize microphone switching.";
            }
        }
        
        private void MicComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateMicNameLabels();
        }
        
        private void MicSelectionSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!_isMicSwitchingEnabled) return;
            
            ApplyMicSelection();
        }        
        private async void ApplyMicSelection()
        {
            if (!_isMicSwitchingEnabled) return;
            
            try
            {
                StatusTextBlock.Text = "Applying microphone selection...";
                
                int selection = (int)Math.Round(MicSelectionSlider.Value);
                MMDevice selectedMic = selection == 0 ? _micOne : _micTwo;
                MMDevice otherMic = selection == 0 ? _micTwo : _micOne;
                
                if (selectedMic == null)
                {
                    throw new InvalidOperationException("Selected microphone device is null");
                }
                
                if (string.IsNullOrEmpty(selectedMic.ID))
                {
                    throw new InvalidOperationException("Selected microphone has invalid device ID");
                }
                
                StatusTextBlock.Text = $"Setting {selectedMic.FriendlyName} as default...";
                
                try
                {
                    bool success = await AudioEndpointManager.SetDefaultAudioEndpointWithTeams(selectedMic.ID, Role.Communications);
                    if (!success)
                    {
                        throw new InvalidOperationException("Failed to set default communications endpoint");
                    }
                    
                    success = await AudioEndpointManager.SetDefaultAudioEndpointWithTeams(selectedMic.ID, Role.Multimedia);
                    if (!success)
                    {
                        throw new InvalidOperationException("Failed to set default multimedia endpoint");
                    }

                    // Update default mic reference and UI
                    _defaultInputDevice = selectedMic;
                    DefaultMicrophoneText.Text = $"{selectedMic.FriendlyName} (System Default)";
                    
                    StatusTextBlock.Text = $"Default microphone set to: {selectedMic.FriendlyName}";
                }
                catch (Exception ex)
                {
                    // Try to restore both microphones to enabled state
                    AudioEndpointManager.SetMicrophoneState(_micOne.ID, true);
                    AudioEndpointManager.SetMicrophoneState(_micTwo.ID, true);
                    throw new InvalidOperationException($"Failed to set default endpoint: {ex.Message}", ex);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error setting default microphone: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusTextBlock.Text = "Failed to set default microphone.";
                
                // Ensure both mics are enabled in case of error
                if (_micOne != null) AudioEndpointManager.SetMicrophoneState(_micOne.ID, true);
                if (_micTwo != null) AudioEndpointManager.SetMicrophoneState(_micTwo.ID, true);
                
                throw;
            }
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            StopOutputFusion();
            _deviceEnumerator?.Dispose();
            _deviceEnumerator = null;
        }
    }
    
    // Improved audio endpoint management
    public static class AudioEndpointManager
    {
        [ComImport, Guid("870af99c-171d-4f9e-af0d-e63df40c2bc9")]
        private class PolicyConfigClient { }

        [ComImport, Guid("568b9108-44bf-40b4-9006-86afe5b5a620")]
        private class PolicyConfigVista { }

        [ComImport, Guid("F8679F50-850A-41CF-9C72-430F290290C8")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPolicyConfig
        {
            [PreserveSig]
            int GetMixFormat([MarshalAs(UnmanagedType.LPWStr)] string deviceId, out IntPtr format);
            
            [PreserveSig]
            int GetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string deviceId, bool useDefaultFormat, out IntPtr format);
            
            [PreserveSig]
            int ResetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string deviceId);
            
            [PreserveSig]
            int SetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string deviceId, IntPtr format, IntPtr matchFormat);
            
            [PreserveSig]
            int GetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string deviceId, bool useDefaultFormat, out long defaultPeriod, out long minimumPeriod);
            
            [PreserveSig]
            int SetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string deviceId, IntPtr processingPeriod);
            
            [PreserveSig]
            int GetShareMode([MarshalAs(UnmanagedType.LPWStr)] string deviceId, out IntPtr shareMode);
            
            [PreserveSig]
            int SetShareMode([MarshalAs(UnmanagedType.LPWStr)] string deviceId, IntPtr shareMode);
            
            [PreserveSig]
            int GetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string deviceId, IntPtr propertyKey, out IntPtr propertyValue);
            
            [PreserveSig]
            int SetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string deviceId, IntPtr propertyKey, IntPtr propertyValue);
            
            [PreserveSig]
            int SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string deviceId, Role role);
        }

        [ComImport, Guid("00000000-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IAgileObject { }

        [ComImport, Guid("568b9108-44bf-40b4-9006-86afe5b5a620")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPolicyConfigVista
        {
            [PreserveSig]
            int GetMixFormat(string deviceId, out IntPtr format);

            [PreserveSig]
            int GetDeviceFormat(string deviceId, int role, out IntPtr format);

            [PreserveSig]
            int SetDeviceFormat(string deviceId, int role, IntPtr format);

            [PreserveSig]
            int GetProcessingPeriod(string deviceId, int role, out long defaultPeriod, out long minimumPeriod);

            [PreserveSig]
            int SetProcessingPeriod(string deviceId, int role, IntPtr period);

            [PreserveSig]
            int GetShareMode(string deviceId, int role, out int shareMode);

            [PreserveSig]
            int SetShareMode(string deviceId, int role, int shareMode);

            [PreserveSig]
            int GetPropertyValue(string deviceId, int role, IntPtr propertyKey, out IntPtr propertyValue);

            [PreserveSig]
            int SetPropertyValue(string deviceId, int role, IntPtr propertyKey, IntPtr propertyValue);

            [PreserveSig]
            int SetDefaultEndpoint(string deviceId, Role role);

            [PreserveSig]
            int SetEndpointVisibility(string deviceId, int isVisible);
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SendMessageTimeout(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam, int fuFlags, int uTimeout, IntPtr lpdwResult);

        private const int HWND_BROADCAST = 0xffff;
        private const int WM_SETTINGCHANGE = 0x001A;
        private const int SMTO_ABORTIFHUNG = 0x0002;

        private static void NotifyDeviceChange()
        {
            SendMessageTimeout((IntPtr)HWND_BROADCAST, WM_SETTINGCHANGE, IntPtr.Zero, IntPtr.Zero, SMTO_ABORTIFHUNG, 1000, IntPtr.Zero);
        }

        public static bool SetDefaultAudioEndpoint(string deviceId, Role role)
        {
            try
            {
                var enumerator = new MMDeviceEnumerator();
                var devices = enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active);
                var device = devices.FirstOrDefault(d => d.ID == deviceId);
                
                if (device == null)
                {
                    MessageBox.Show($"Could not find device with ID: {deviceId}", "Device Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }

                if (device.State != DeviceState.Active)
                {
                    MessageBox.Show("Selected device is not active", "Device Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }

                bool success = false;
                
                // Try with PolicyConfigVista first (better application support)
                try
                {
                    var policyConfigVista = (IPolicyConfigVista)new PolicyConfigVista();
                    int hr = policyConfigVista.SetDefaultEndpoint(deviceId, role);
                    success = hr >= 0;
                    Marshal.ReleaseComObject(policyConfigVista);
                }
                catch
                {
                    // Fall back to regular PolicyConfig
                    try
                    {
                        var policyConfig = (IPolicyConfig)new PolicyConfigClient();
                        int hr = policyConfig.SetDefaultEndpoint(deviceId, role);
                        success = hr >= 0;
                        Marshal.ReleaseComObject(policyConfig);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to set default device: {ex.Message}");
                        success = false;
                    }
                }

                if (success)
                {
                    // Give Windows a moment to process the change
                    System.Threading.Thread.Sleep(100);
                    
                    // Notify all applications about the change
                    NotifyDeviceChange();
                    
                    // Additional sleep to ensure notification is processed
                    System.Threading.Thread.Sleep(100);
                }

                return success;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in SetDefaultAudioEndpoint: {ex.Message}");
                return false;
            }
        }

        public static bool SetMicrophoneState(string deviceId, bool enabled)
        {
            try
            {
                var enumerator = new MMDeviceEnumerator();
                var device = enumerator.GetDevice(deviceId);
                
                if (device == null || device.State != DeviceState.Active)
                {
                    return false;
                }

                // Get the audio endpoint volume
                var audioEndpointVolume = device.AudioEndpointVolume;
                
                if (enabled)
                {
                    // Enable and restore volume
                    audioEndpointVolume.Mute = false;
                }
                else
                {
                    // Disable by muting
                    audioEndpointVolume.Mute = true;
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error setting microphone state: {ex.Message}");
                return false;
            }
        }

        public static async Task<bool> SetDefaultAudioEndpointWithTeams(string deviceId, Role role)
        {
            bool success = SetDefaultAudioEndpoint(deviceId, role);
            
            if (success && role == Role.Communications)
            {
                var enumerator = new MMDeviceEnumerator();
                var devices = enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active);
                
                // Mute all other microphones
                foreach (var device in devices)
                {
                    if (device.ID != deviceId)
                    {
                        SetMicrophoneState(device.ID, false);
                    }
                }
                
                // Unmute the selected microphone
                SetMicrophoneState(deviceId, true);
                
                // Still try to update Teams UI as fallback
                await TeamsAudioManager.UpdateTeamsAudioDevice(deviceId, true);
            }

            return success;
        }
    }

    public static class TeamsAudioManager
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint WM_COMMAND = 0x0111;
        private const int ID_FILE_SETTINGS = 0x0001;

        public static async Task UpdateTeamsAudioDevice(string deviceId, bool isMicrophone)
        {
            try
            {
                // First try using Microsoft Graph API (requires user to be signed in)
                if (await TryUpdateViaGraphApi(deviceId, isMicrophone))
                {
                    return;
                }

                // Fallback: Try direct Teams interaction
                if (await TryUpdateViaTeamsUI(deviceId, isMicrophone))
                {
                    return;
                }

                System.Diagnostics.Debug.WriteLine("Failed to update Teams audio device settings");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating Teams settings: {ex.Message}");
            }
        }

        private static async Task<bool> TryUpdateViaGraphApi(string deviceId, bool isMicrophone)
        {
            try
            {
                // Teams Graph API endpoint
                string graphEndpoint = "https://graph.microsoft.com/v1.0/me/onlineMeetings/settings";
                
                // This requires Microsoft.Graph NuGet package and user authentication
                // For now, return false to use the UI automation fallback
                return false;
            }
            catch
            {
                return false;
            }
        }

        private static async Task<bool> TryUpdateViaTeamsUI(string deviceId, bool isMicrophone)
        {
            try
            {
                // Find Teams main window
                IntPtr teamsWindow = FindWindow("Chrome_WidgetWin_1", null);
                if (teamsWindow == IntPtr.Zero)
                {
                    // Try finding new Teams window
                    teamsWindow = FindWindow("Microsoft.Windows.WebView2.WebView", null);
                }

                if (teamsWindow != IntPtr.Zero)
                {
                    // Try to force Teams to refresh its audio devices
                    SendMessage(teamsWindow, WM_COMMAND, (IntPtr)ID_FILE_SETTINGS, IntPtr.Zero);
                    
                    // Give Teams time to process
                    await Task.Delay(500);
                    
                    // This simulates clicking the settings button
                    // Teams should then detect the system default device change
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }
    }
}
// MainWindow.xaml.cs

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text; // Not strictly needed with this version but often useful
using System.Windows;
using System.Windows.Controls;
using NAudio.CoreAudioApi;
using NAudio.Wave;
using NAudio.Wave.SampleProviders; // For BufferedWaveProvider
using System.ComponentModel;
using System.Runtime.InteropServices; // For COMException and Marshal
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
        private MMDevice _micOne; // Reference to the selected Mic 1 device
        private MMDevice _micTwo; // Reference to the selected Mic 2 device

        public MainWindow()
        {
            InitializeComponent();
            
            // Initialize device enumerator once
            _deviceEnumerator = new MMDeviceEnumerator(); 
            
            // Initialize slider to disabled state
            MicSelectionSlider.IsEnabled = false;
            
            LoadAudioDevices();
        }

        private void DisposeAudioDevice(MMDevice device)
        {
            try
            {
                device?.Dispose();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error disposing MMDevice: {ex.Message}");
            }
        }

        private void DisposeDeviceList(List<MMDevice> devices)
        {
            if (devices != null)
            {
                foreach (var device in devices)
                {
                    DisposeAudioDevice(device);
                }
                devices.Clear(); // Clear the list after disposing its contents
            }
        }

        private void LoadAudioDevices()
        {
            RefreshDevicesButton.IsEnabled = false; // Disable while loading

            try
            {
                StatusTextBlock.Text = "Loading audio devices...";

                // Save selections before clearing
                string selectedSecondaryHeadset = SecondaryHeadsetComboBox.SelectedItem?.ToString();
                string selectedMicOneName = MicOneComboBox.SelectedItem?.ToString();
                string selectedMicTwoName = MicTwoComboBox.SelectedItem?.ToString();

                // Dispose existing MMDevice objects before re-populating
                DisposeDeviceList(_outputDevices);
                DisposeDeviceList(_inputDevices);
                DisposeAudioDevice(_defaultOutputDevice); _defaultOutputDevice = null;
                DisposeAudioDevice(_defaultInputDevice);  _defaultInputDevice = null;
                // _micOne and _micTwo are references from _inputDevices; they will be nulled and re-assigned
                _micOne = null;
                _micTwo = null;
                
                // Get default output device
                try 
                {
                    _defaultOutputDevice = _deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                    DefaultOutputDeviceText.Text = $"{_defaultOutputDevice.FriendlyName} (System Default)";
                }
                catch (COMException cex) when ((uint)cex.ErrorCode == 0x80070490) // ERROR_NOT_FOUND
                {
                    DefaultOutputDeviceText.Text = "No default output device found.";
                    _defaultOutputDevice = null; // Ensure it's null
                    System.Diagnostics.Debug.WriteLine("LoadAudioDevices: No default output device found.");
                }
                catch (Exception ex)
                {
                    DefaultOutputDeviceText.Text = "Error getting default output device.";
                     _defaultOutputDevice = null;
                    System.Diagnostics.Debug.WriteLine($"LoadAudioDevices: Error getting default output device: {ex.Message}");
                }
                
                // Get default input device - try both roles
                try 
                {
                    // Try communications role first as it's often what apps like Teams prefer
                    _defaultInputDevice = _deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications);
                }
                catch (COMException cex) when ((uint)cex.ErrorCode == 0x80070490)
                {
                    System.Diagnostics.Debug.WriteLine("LoadAudioDevices: No default Communications capture device. Trying Multimedia.");
                    try
                    {
                        _defaultInputDevice = _deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Multimedia);
                    }
                    catch (COMException cex2) when ((uint)cex2.ErrorCode == 0x80070490)
                    {
                        _defaultInputDevice = null; // Ensure it's null
                        System.Diagnostics.Debug.WriteLine("LoadAudioDevices: No default Multimedia capture device either.");
                    }
                    catch (Exception ex2)
                    {
                        _defaultInputDevice = null;
                        System.Diagnostics.Debug.WriteLine($"LoadAudioDevices: Error getting Multimedia capture device: {ex2.Message}");
                    }
                }
                catch(Exception ex) // Other exceptions for Communications role
                {
                    System.Diagnostics.Debug.WriteLine($"LoadAudioDevices: Error getting default Communications input device: {ex.Message}");
                    _defaultInputDevice = null;
                }

                if (_defaultInputDevice != null)
                {
                    DefaultMicrophoneText.Text = $"{_defaultInputDevice.FriendlyName} (System Default)";
                }
                else
                {
                    DefaultMicrophoneText.Text = "No default microphone found.";
                }
                
                // Get all active output devices
                _outputDevices.AddRange(_deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active).ToList());
                // Get all active input devices
                _inputDevices.AddRange(_deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active).ToList());
                
                // Clear and repopulate combo boxes
                PopulateComboBoxes();
                
                // Restore selections or select defaults
                RestoreSelection(SecondaryHeadsetComboBox, selectedSecondaryHeadset, _outputDevices);

                // For MicOne, try to select the current system default if available, otherwise restore previous or first
                if (_defaultInputDevice != null)
                {
                    int defaultMicIndexInList = _inputDevices.FindIndex(d => d.ID == _defaultInputDevice.ID);
                    if (defaultMicIndexInList != -1 && MicOneComboBox.Items.Count > defaultMicIndexInList)
                    {
                        MicOneComboBox.SelectedIndex = defaultMicIndexInList;
                    }
                    else
                    {
                        RestoreSelection(MicOneComboBox, selectedMicOneName, _inputDevices);
                    }
                }
                else
                {
                     RestoreSelection(MicOneComboBox, selectedMicOneName, _inputDevices);
                }

                RestoreSelection(MicTwoComboBox, selectedMicTwoName, _inputDevices);
                
                // Ensure different microphones selected if possible
                if (MicOneComboBox.SelectedItem != null && MicOneComboBox.SelectedItem.ToString() == MicTwoComboBox.SelectedItem?.ToString() && MicTwoComboBox.Items.Count > 1)
                {
                    MicTwoComboBox.SelectedIndex = (MicOneComboBox.SelectedIndex + 1) % MicTwoComboBox.Items.Count;
                }
                
                // Update mic name labels and device references (_micOne, _micTwo)
                UpdateMicNameLabels();
                
                // Update slider position based on current default mic (after _micOne/_micTwo are set)
                UpdateSliderPosition();
                
                StatusTextBlock.Text = "Audio devices loaded. Ready.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading audio devices: {ex.Message}\n{ex.StackTrace}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusTextBlock.Text = "Error loading audio devices.";
                 System.Diagnostics.Debug.WriteLine($"CRITICAL ERROR in LoadAudioDevices: {ex.ToString()}");
            }
            finally
            {
                RefreshDevicesButton.IsEnabled = true;
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
            // This method updates the slider based on which of the selected _micOne or _micTwo
            // is currently the system's default. It re-fetches the system default to be sure.

            MMDevice currentSystemDefaultComm = null;
            MMDevice currentSystemDefaultMulti = null;
            MMDevice actualCurrentSystemDefault = null; // This will be disposed if not assigned to _defaultInputDevice

            try
            {
                // Re-fetch current system default input device to ensure accuracy
                try { currentSystemDefaultComm = _deviceEnumerator?.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications); }
                catch (COMException cex) when ((uint)cex.ErrorCode == 0x80070490) { /* Not found, fine */ }
                
                try { currentSystemDefaultMulti = _deviceEnumerator?.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Multimedia); }
                catch (COMException cex) when ((uint)cex.ErrorCode == 0x80070490) { /* Not found, fine */ }

                // Prefer Communications role for "default" in this context if it exists
                actualCurrentSystemDefault = currentSystemDefaultComm ?? currentSystemDefaultMulti;
                // If currentSystemDefaultComm existed, actualCurrentSystemDefault points to it.
                // If it didn't, but currentSystemDefaultMulti did, it points to multi.
                // If neither, it's null.

                if (actualCurrentSystemDefault != null)
                {
                    // If the global _defaultInputDevice is different from the newly fetched one, dispose the old one.
                    if (_defaultInputDevice != null && _defaultInputDevice.ID != actualCurrentSystemDefault.ID)
                    {
                        DisposeAudioDevice(_defaultInputDevice);
                        _defaultInputDevice = null; // Null it out before reassignment
                    }
                    else if (_defaultInputDevice != null && _defaultInputDevice.ID == actualCurrentSystemDefault.ID)
                    {
                        // It's the same device. We need to dispose the 'actualCurrentSystemDefault' we just fetched
                        // to avoid double-managing the same COM object.
                        if (actualCurrentSystemDefault != _defaultInputDevice) // Ensure they are not already the same instance
                        {
                             DisposeAudioDevice(actualCurrentSystemDefault);
                             actualCurrentSystemDefault = _defaultInputDevice; // Point back to the managed instance
                        }
                    }
                    
                    _defaultInputDevice = actualCurrentSystemDefault; // Assign the (potentially new) default
                    DefaultMicrophoneText.Text = $"{_defaultInputDevice.FriendlyName} (System Default)";

                    if (_isMicSwitchingEnabled) // Only adjust slider if mic switching is active
                    {
                        if (_micOne != null && _defaultInputDevice.ID == _micOne.ID)
                        {
                            MicSelectionSlider.Value = 0;
                        }
                        else if (_micTwo != null && _defaultInputDevice.ID == _micTwo.ID)
                        {
                            MicSelectionSlider.Value = 1;
                        }
                        // If default is neither _micOne nor _micTwo, slider position is ambiguous, could leave as is or reset.
                    }
                }
                else // No default input device found by the system
                {
                    DisposeAudioDevice(_defaultInputDevice); // Dispose the old one if it existed
                    _defaultInputDevice = null;
                    DefaultMicrophoneText.Text = "No default microphone found.";
                    // If mic switching is enabled, slider might not match.
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in UpdateSliderPosition: {ex.Message}");
                // Don't change UI text here, as it might be during a sensitive operation.
            }
            finally
            {
                // If currentSystemDefaultComm was fetched but NOT assigned to _defaultInputDevice (e.g., actual was Multi)
                // it needs to be disposed.
                if (currentSystemDefaultComm != null && currentSystemDefaultComm != _defaultInputDevice)
                {
                    DisposeAudioDevice(currentSystemDefaultComm);
                }
                // If currentSystemDefaultMulti was fetched but NOT assigned to _defaultInputDevice (e.g., actual was Comm, or both null)
                // it also needs to be disposed. (This handles the case where actualCurrentSystemDefault was multi but became null)
                if (currentSystemDefaultMulti != null && currentSystemDefaultMulti != _defaultInputDevice && currentSystemDefaultMulti != currentSystemDefaultComm) // Added check to avoid double dispose if Comm == Multi (unlikely but possible)
                {
                    DisposeAudioDevice(currentSystemDefaultMulti);
                }
            }
        }
        
        private void UpdateMicNameLabels()
        {
            _micOne = FindDeviceByName(MicOneComboBox.SelectedItem?.ToString(), _inputDevices);
            _micTwo = FindDeviceByName(MicTwoComboBox.SelectedItem?.ToString(), _inputDevices);

            MicOneName.Text = _micOne?.FriendlyName ?? "Mic 1 (None)";
            MicTwoName.Text = _micTwo?.FriendlyName ?? "Mic 2 (None)";
        }
        
        private MMDevice FindDeviceByName(string friendlyName, List<MMDevice> deviceList)
        {
            if (string.IsNullOrEmpty(friendlyName) || deviceList == null)
                return null;
                
            return deviceList.FirstOrDefault(d => d.FriendlyName == friendlyName);
        }
        
        private void RestoreSelection(ComboBox comboBox, string previousValueName, List<MMDevice> deviceList)
        {
            if (comboBox.Items.Count == 0) return;
            
            if (!string.IsNullOrEmpty(previousValueName))
            {
                // First, try to find by exact FriendlyName in the current list of MMDevices
                var deviceInList = deviceList?.FirstOrDefault(d => d.FriendlyName == previousValueName);
                if (deviceInList != null)
                {
                    // Find the index of this name in the ComboBox items (which are strings)
                    for (int i = 0; i < comboBox.Items.Count; i++)
                    {
                        if (comboBox.Items[i].ToString() == deviceInList.FriendlyName)
                        {
                            comboBox.SelectedIndex = i;
                            return;
                        }
                    }
                }
            }
            
            // If no match or no previous, select first item if available
            if (comboBox.Items.Count > 0)
            {
                comboBox.SelectedIndex = 0;
            }
            else
            {
                comboBox.SelectedIndex = -1; // No items, no selection
            }
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
            if (SecondaryHeadsetComboBox.SelectedIndex == -1 || SecondaryHeadsetComboBox.SelectedItem == null)
            {
                MessageBox.Show("Please select a secondary headset device.", "Selection Missing", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            
            string selectedDeviceName = SecondaryHeadsetComboBox.SelectedItem.ToString();
            MMDevice secondaryDevice = FindDeviceByName(selectedDeviceName, _outputDevices); // Use our helper
            
            if (secondaryDevice == null)
            {
                MessageBox.Show("Cannot find selected output device. It may have been disconnected. Please refresh.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            
            if (_defaultOutputDevice == null) {
                 MessageBox.Show("System default output device is not available. Cannot start fusion.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (_defaultOutputDevice.ID == secondaryDevice.ID)
            {
                MessageBox.Show("The secondary headset cannot be the same as the system default output for this function.", "Device Conflict", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                _audioSourceCapture = new WasapiLoopbackCapture(_defaultOutputDevice); // Capture from current system default
                _outputFusionBuffer = new BufferedWaveProvider(_audioSourceCapture.WaveFormat)
                { BufferDuration = TimeSpan.FromMilliseconds(200), DiscardOnBufferOverflow = true };

                _audioSourceCapture.DataAvailable += (s, args) => _outputFusionBuffer?.AddSamples(args.Buffer, 0, args.BytesRecorded);
                
                _audioSourceCapture.RecordingStopped += (s, args) =>
                {
                    // This can be called due to device changes or errors.
                    Dispatcher.Invoke(() => {
                        if (_isOutputFusionRunning) // Only act if we thought we were running
                        {
                            StatusTextBlock.Text = "Output fusion capture stopped unexpectedly.";
                            StopOutputFusion(); // Gracefully attempt to clean up
                        }
                         if (args.Exception != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"Loopback capture stopped with exception: {args.Exception.Message}");
                        }
                    });
                };

                _secondaryHeadsetPlayer = new WasapiOut(secondaryDevice, AudioClientShareMode.Shared, true, 100); // useLatency: true, latency: 100ms
                _secondaryHeadsetPlayer.Init(_outputFusionBuffer);

                _secondaryHeadsetPlayer.PlaybackStopped += (s, args) =>
                {
                    Dispatcher.Invoke(() => {
                        if (args.Exception != null && _isOutputFusionRunning)
                        {
                            StatusTextBlock.Text = $"Secondary headset playback error: {args.Exception.Message.Split('\n')[0]}";
                            System.Diagnostics.Debug.WriteLine($"Secondary headset playback error: {args.Exception.ToString()}");
                            StopOutputFusion();
                        }
                        else if (_isOutputFusionRunning) // Stopped without error, but we didn't initiate stop
                        {
                            StatusTextBlock.Text = "Secondary headset playback stopped unexpectedly.";
                            StopOutputFusion();
                        }
                    });
                };

                _audioSourceCapture.StartRecording();
                _secondaryHeadsetPlayer.Play();

                _isOutputFusionRunning = true;
                OutputFusionButton.Content = "Stop Output Fusion";
                SecondaryHeadsetComboBox.IsEnabled = false;
                RefreshDevicesButton.IsEnabled = false; // Disable refresh during fusion
                StatusTextBlock.Text = $"Output Fusion: Mirroring '{_defaultOutputDevice.FriendlyName}' to '{secondaryDevice.FriendlyName}'";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error starting output fusion: {ex.Message.Split('\n')[0]}", "Output Fusion Error", MessageBoxButton.OK, MessageBoxImage.Error);
                System.Diagnostics.Debug.WriteLine($"Error starting output fusion: {ex.ToString()}");
                StopOutputFusion(); // Clean up anything that might have started
            }
        }

        private void StopOutputFusion()
        {
            bool wasRunning = _isOutputFusionRunning;
            _isOutputFusionRunning = false; // Set this early to prevent re-entry from event handlers

            try
            {
                if (_audioSourceCapture != null)
                {
                    _audioSourceCapture.StopRecording(); // This will trigger RecordingStopped event
                    _audioSourceCapture.Dispose();
                    _audioSourceCapture = null;
                }
                
                if (_secondaryHeadsetPlayer != null)
                {
                    _secondaryHeadsetPlayer.Stop(); // This will trigger PlaybackStopped event
                    _secondaryHeadsetPlayer.Dispose();
                    _secondaryHeadsetPlayer = null;
                }
                
                _outputFusionBuffer = null; // Let GC handle this after consumers are done
                
                OutputFusionButton.Content = "Start Output Fusion";
                SecondaryHeadsetComboBox.IsEnabled = true;
                RefreshDevicesButton.IsEnabled = true; // Re-enable refresh
                
                if (wasRunning && !StatusTextBlock.Text.ToLower().Contains("error") && !StatusTextBlock.Text.ToLower().Contains("unexpectedly"))
                {
                    StatusTextBlock.Text = "Output fusion stopped.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error stopping output fusion: {ex.Message.Split('\n')[0]}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                System.Diagnostics.Debug.WriteLine($"Error stopping output fusion: {ex.ToString()}");
                // Ensure UI state is reset even if an error occurs during stop
                OutputFusionButton.Content = "Start Output Fusion";
                SecondaryHeadsetComboBox.IsEnabled = true;
                RefreshDevicesButton.IsEnabled = true;
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
                    // Optionally, re-enable combo boxes if they were disabled
                    // MicOneComboBox.IsEnabled = true;
                    // MicTwoComboBox.IsEnabled = true;
                    StatusTextBlock.Text = "Microphone switching disabled.";
                }
                else
                {
                    // Enable mic switching
                    if (_micOne == null || _micTwo == null)
                    {
                        MessageBox.Show("Please select valid primary and secondary microphones.", "Selection Missing", MessageBoxButton.OK, MessageBoxImage.Warning);
                        StatusTextBlock.Text = "Mic selection incomplete for switching.";
                        return;
                    }
                    
                    if (_micOne.ID == _micTwo.ID)
                    {
                         MessageBox.Show("Primary and secondary microphones cannot be the same device for switching.", "Device Conflict", MessageBoxButton.OK, MessageBoxImage.Warning);
                         StatusTextBlock.Text = "Cannot switch: Mics are the same.";
                        return;
                    }

                    _isMicSwitchingEnabled = true;
                    MicSwitchingButton.Content = "Disable Mic Switching";
                    MicSelectionSlider.IsEnabled = true;
                    // Optionally, disable combo boxes while switching is active
                    // MicOneComboBox.IsEnabled = false;
                    // MicTwoComboBox.IsEnabled = false;
                    
                    UpdateSliderPosition(); // Set initial slider position based on current default
                    ApplyMicSelection(); // Apply initial selection based on slider (or current default)
                    StatusTextBlock.Text = "Microphone switching enabled. Use slider.";
                }
            }
            catch (Exception ex)
            {
                var error = $"Error toggling mic switching: {ex.Message}\n{ex.StackTrace}";
                System.Diagnostics.Debug.WriteLine(error);
                MessageBox.Show(error, "Mic Switching Error", MessageBoxButton.OK, MessageBoxImage.Error);
                
                _isMicSwitchingEnabled = false;
                MicSwitchingButton.Content = "Enable Mic Switching";
                MicSelectionSlider.IsEnabled = false;
                StatusTextBlock.Text = "Error initializing mic switching.";
            }
        }
        
        private void MicComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_deviceEnumeratorReady()) return; // Avoid operations if LoadAudioDevices hasn't completed or enumerator is null

            UpdateMicNameLabels(); // Update _micOne, _micTwo and their name labels

            // If mic switching is active, and a combo box changed,
            // we might need to re-evaluate the slider or the current default.
            if (_isMicSwitchingEnabled)
            {
                // Check if the new selection for _micOne or _micTwo is now the default
                UpdateSliderPosition();
                // It might be good to re-apply selection if a combobox changed while active.
                // However, ApplyMicSelection can be intensive. For now, let slider change trigger it.
            }
        }
        
        private bool _deviceEnumeratorReady() => _deviceEnumerator != null && _inputDevices != null && _outputDevices != null;

        private async void MicSelectionSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!_isMicSwitchingEnabled || !MicSelectionSlider.IsFocused && !MicSelectionSlider.IsMouseOver && e.OldValue == e.NewValue)
            {
                // Don't apply if not enabled, or if not user-initiated (e.g. programmatic update of value)
                // IsFocused/IsMouseOver is a heuristic for user interaction. Better would be an IsPressed check or a flag.
                return;
            }
            if (IsLoaded) // Ensure window is fully loaded to avoid issues during init.
            {
               await ApplyMicSelectionAsync();
            }
        }        

        private async Task ApplyMicSelectionAsync() // Renamed to Async and made async Task
        {
            if (!_isMicSwitchingEnabled) return;

            if (_micOne == null || _micTwo == null)
            {
                StatusTextBlock.Text = "Microphone devices not properly selected.";
                // MessageBox.Show("Please ensure both microphones are selected from the dropdowns.", "Device Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (_micOne.ID == _micTwo.ID)
            {
                StatusTextBlock.Text = "Mic 1 and Mic 2 cannot be the same device for switching.";
                // MessageBox.Show("Please select two different microphones for switching.", "Device Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Disable UI elements during the operation
            MicSelectionSlider.IsEnabled = false;
            RefreshDevicesButton.IsEnabled = false;
            MicSwitchingButton.IsEnabled = false;
            MicOneComboBox.IsEnabled = false;
            MicTwoComboBox.IsEnabled = false;

            try
            {
                StatusTextBlock.Text = "Applying microphone selection...";
                await Task.Delay(50); // Brief delay to allow UI update
                
                int selection = (int)Math.Round(MicSelectionSlider.Value);
                MMDevice targetMic = selection == 0 ? _micOne : _micTwo;
                MMDevice otherMic = selection == 0 ? _micTwo : _micOne;

                if (targetMic == null || string.IsNullOrEmpty(targetMic.ID))
                {
                    throw new InvalidOperationException($"Target microphone ({ (selection == 0 ? "Mic1" : "Mic2") }) is null or has an invalid ID.");
                }
                System.Diagnostics.Debug.WriteLine($"Target Mic: {targetMic.FriendlyName} (ID: {targetMic.ID})");
                if (otherMic != null) System.Diagnostics.Debug.WriteLine($"Other Mic: {otherMic.FriendlyName} (ID: {otherMic.ID})");


                // 1. Set target mic as default for Communications and Multimedia
                StatusTextBlock.Text = $"Setting {targetMic.FriendlyName} as system default...";
                await Task.Delay(50);
                bool commsSuccess = AudioEndpointManager.SetDefaultAudioEndpoint(targetMic.ID, Role.Communications);
                bool multimediaSuccess = AudioEndpointManager.SetDefaultAudioEndpoint(targetMic.ID, Role.Multimedia);

                if (!commsSuccess && !multimediaSuccess) // If neither role could be set
                {
                     // Try multimedia if comms failed, or vice versa, just in case
                    if (!commsSuccess) commsSuccess = AudioEndpointManager.SetDefaultAudioEndpoint(targetMic.ID, Role.Communications);
                    if (!multimediaSuccess) multimediaSuccess = AudioEndpointManager.SetDefaultAudioEndpoint(targetMic.ID, Role.Multimedia);

                    if (!commsSuccess && !multimediaSuccess)
                    {
                        throw new InvalidOperationException($"Failed to set {targetMic.FriendlyName} as system default for any role. Check admin rights or device state. Comms: {commsSuccess}, Multi: {multimediaSuccess}");
                    }
                }
                
                // Update internal reference and UI text for the default mic
                DisposeAudioDevice(_defaultInputDevice); // Dispose the old one
                _defaultInputDevice = _deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications); // Re-fetch
                DefaultMicrophoneText.Text = $"{_defaultInputDevice.FriendlyName} (System Default)";
                StatusTextBlock.Text = $"System default: {_defaultInputDevice.FriendlyName}. Influencing Teams...";
                System.Diagnostics.Debug.WriteLine($"System default set to: {_defaultInputDevice.FriendlyName}");

                await Task.Delay(300); // Brief pause for system to settle (was 250)

                // 2. Temporarily disable the "other" microphone to encourage Teams to switch
                if (otherMic != null && !string.IsNullOrEmpty(otherMic.ID))
                {
                    string otherMicInstanceId = AudioEndpointManager.GetDeviceInstanceId(otherMic.ID);
                    if (!string.IsNullOrEmpty(otherMicInstanceId))
                    {
                        System.Diagnostics.Debug.WriteLine($"Attempting to disable other mic: {otherMic.FriendlyName} (PnP Instance ID: {otherMicInstanceId})");
                        StatusTextBlock.Text = $"Disabling {otherMic.FriendlyName}...";
                        await Task.Delay(50);
                        if (AudioEndpointManager.SetDeviceState(otherMicInstanceId, false)) // Disable
                        {
                            System.Diagnostics.Debug.WriteLine($"Successfully disabled {otherMic.FriendlyName}. Waiting for Teams/OS to react...");
                            StatusTextBlock.Text = $"{otherMic.FriendlyName} disabled. Waiting for Teams...";
                            await Task.Delay(2000); // Give Teams/OS time to react ( crucial delay - was 1500)

                            System.Diagnostics.Debug.WriteLine($"Attempting to re-enable other mic: {otherMic.FriendlyName}");
                            StatusTextBlock.Text = $"Re-enabling {otherMic.FriendlyName}...";
                            await Task.Delay(50);
                            if(AudioEndpointManager.SetDeviceState(otherMicInstanceId, true)) // Re-enable
                            {
                                System.Diagnostics.Debug.WriteLine($"Successfully re-enabled {otherMic.FriendlyName}.");
                                StatusTextBlock.Text = $"{targetMic.FriendlyName} is default. {otherMic.FriendlyName} re-enabled. Check Teams.";
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"Failed to re-enable {otherMic.FriendlyName}. Manual check needed.");
                                StatusTextBlock.Text = $"Error re-enabling {otherMic.FriendlyName}.";
                                MessageBox.Show($"Failed to re-enable microphone '{otherMic.FriendlyName}'. You may need to enable it manually in Windows Sound settings.", "Device Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"Failed to disable other mic: {otherMic.FriendlyName}");
                            StatusTextBlock.Text = $"Could not disable {otherMic.FriendlyName}.";
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"Could not get PnP Instance ID for other mic: {otherMic.FriendlyName}. Cannot disable/enable it.");
                        StatusTextBlock.Text = $"No PnP ID for {otherMic.FriendlyName}.";
                    }
                }
                await Task.Delay(500); // Final small delay
                StatusTextBlock.Text = $"Default: {targetMic.FriendlyName}. Output Fusion: {(_isOutputFusionRunning ? "ON" : "OFF")}.";
            }
            catch (Exception ex)
            {
                string errorDetails = $"Error applying mic selection: {ex.Message}\n{ex.StackTrace}";
                if (ex.InnerException != null) errorDetails += $"\nInner: {ex.InnerException.Message}\n{ex.InnerException.StackTrace}";
                System.Diagnostics.Debug.WriteLine(errorDetails);
                MessageBox.Show(errorDetails, "Microphone Switch Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusTextBlock.Text = "Error switching microphone.";

                // Fallback: try to ensure both selected mics are re-enabled if an error occurred
                try
                {
                    if (_micOne != null && !string.IsNullOrEmpty(_micOne.ID)) {
                        string id1 = AudioEndpointManager.GetDeviceInstanceId(_micOne.ID);
                        if(!string.IsNullOrEmpty(id1)) AudioEndpointManager.SetDeviceState(id1, true);
                    }
                    if (_micTwo != null && !string.IsNullOrEmpty(_micTwo.ID)) {
                        string id2 = AudioEndpointManager.GetDeviceInstanceId(_micTwo.ID);
                        if(!string.IsNullOrEmpty(id2)) AudioEndpointManager.SetDeviceState(id2, true);
                    }
                } catch (Exception restEx) {
                     System.Diagnostics.Debug.WriteLine($"Error during fallback re-enable of mics: {restEx.Message}");
                }
            }
            finally
            {
                // Re-enable UI elements
                MicSelectionSlider.IsEnabled = true;
                RefreshDevicesButton.IsEnabled = true;
                MicSwitchingButton.IsEnabled = true;
                MicOneComboBox.IsEnabled = true;
                MicTwoComboBox.IsEnabled = true;

                // Refresh slider position based on the actual current default, which might have changed
                // or reverted if something went wrong.
                UpdateSliderPosition();
                // LoadAudioDevices(); // This is too disruptive. UpdateSliderPosition should suffice.
            }
        }

        // Original ApplyMicSelection (non-async) kept for reference or if needed, but async version is preferred.
        private void ApplyMicSelection() 
        {
            // This is a synchronous wrapper for the async method.
            // However, for UI responsiveness, directly calling the async method from event handlers is better.
            // If you need to call it from a non-async context and wait, you might do:
            // Task.Run(async () => await ApplyMicSelectionAsync()).GetAwaiter().GetResult();
            // But this is generally discouraged on UI threads.
            // For now, let's assume MicSelectionSlider_ValueChanged calls ApplyMicSelectionAsync directly.
            System.Diagnostics.Debug.WriteLine("Synchronous ApplyMicSelection called - consider using async version.");
            _ = ApplyMicSelectionAsync(); // Fire and forget (not ideal, but better than blocking UI)
        }


        private void Window_Closing(object sender, CancelEventArgs e)
        {
            // Stop any ongoing operations
            StopOutputFusion(); // Disposes its NAudio components

            // If mic switching was enabled, consider reverting to a known state or ensuring mics are enabled.
            if (_isMicSwitchingEnabled)
            {
                _isMicSwitchingEnabled = false; // Prevent further actions
                // Optionally, try to ensure both _micOne and _micTwo are enabled if they were part of the switching.
                // This is a "best effort" cleanup.
                try
                {
                    if (_micOne != null && !string.IsNullOrEmpty(_micOne.ID)) {
                        string id1 = AudioEndpointManager.GetDeviceInstanceId(_micOne.ID);
                        if(!string.IsNullOrEmpty(id1)) AudioEndpointManager.SetDeviceState(id1, true);
                    }
                    if (_micTwo != null && !string.IsNullOrEmpty(_micTwo.ID)) {
                        string id2 = AudioEndpointManager.GetDeviceInstanceId(_micTwo.ID);
                        if(!string.IsNullOrEmpty(id2)) AudioEndpointManager.SetDeviceState(id2, true);
                    }
                } catch (Exception ex) {
                     System.Diagnostics.Debug.WriteLine($"Error ensuring mics enabled on close: {ex.Message}");
                }
            }

            // Dispose individual top-level MMDevice references that might not be in the lists
            // or are separately managed.
            DisposeAudioDevice(_defaultOutputDevice); _defaultOutputDevice = null;
            DisposeAudioDevice(_defaultInputDevice); _defaultInputDevice = null;
            // _micOne and _micTwo are references to devices in _inputDevices,
            // so they will be disposed when _inputDevices is disposed.

            // Dispose all devices collected in the lists
            DisposeDeviceList(_outputDevices);
            DisposeDeviceList(_inputDevices);

            // Finally, dispose the MMDeviceEnumerator itself
            _deviceEnumerator?.Dispose();
            _deviceEnumerator = null;

            System.Diagnostics.Debug.WriteLine("AudioFusion closed and resources released.");
        }
    }
    
    // AudioEndpointManager and TeamsAudioManager classes would follow here
    // (but are in their own section in the next message for clarity)
}
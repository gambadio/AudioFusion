using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.ComponentModel;
using System.Runtime.InteropServices;
using NAudio.CoreAudioApi;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;

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

        // Input Fusion Components
        private WasapiCapture _primaryMicrophoneCapture;
        private WasapiCapture _secondaryMicrophoneCapture;
        private MixingSampleProvider _inputMixer;
        private WaveOut _virtualMicOutput;
        private bool _isInputFusionRunning = false;
        private IWaveProvider _primaryMicrophoneBuffer;
        private IWaveProvider _secondaryMicrophoneBuffer;
        private VirtualAudioCaptureClient _virtualAudioCaptureClient;

        public MainWindow()
        {
            InitializeComponent();
            
            // Initialize device enumerator once
            _deviceEnumerator = new MMDeviceEnumerator();
            
            LoadAudioDevices();

            SecondaryVolumeSlider.ValueChanged += SecondaryVolumeSlider_ValueChanged;
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

                // Save selection before clearing
                string selectedSecondaryHeadset = SecondaryHeadsetComboBox.SelectedItem?.ToString();
                string selectedSecondaryMic = SecondaryMicrophoneComboBox.SelectedItem?.ToString();

                // Dispose existing MMDevice objects before re-populating
                DisposeDeviceList(_outputDevices);
                DisposeDeviceList(_inputDevices);
                DisposeAudioDevice(_defaultOutputDevice); _defaultOutputDevice = null;
                DisposeAudioDevice(_defaultInputDevice); _defaultInputDevice = null;
                
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
                
                // Get default input device
                try 
                {
                    // Try Multimedia first (matches Windows Sound UI), fallback to Communications
                    try {
                        _defaultInputDevice = _deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Multimedia);
                        DefaultInputDeviceText.Text = $"{_defaultInputDevice.FriendlyName} (System Default)";
                    } catch (Exception)
                    {
                        _defaultInputDevice = _deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications);
                        DefaultInputDeviceText.Text = $"{_defaultInputDevice.FriendlyName} (System Default)";
                    }
                }
                catch (COMException cex) when ((uint)cex.ErrorCode == 0x80070490) // ERROR_NOT_FOUND
                {
                    DefaultInputDeviceText.Text = "No default microphone found.";
                    _defaultInputDevice = null; // Ensure it's null
                    System.Diagnostics.Debug.WriteLine("LoadAudioDevices: No default microphone found.");
                }
                catch (Exception ex)
                {
                    DefaultInputDeviceText.Text = "Error getting default microphone.";
                     _defaultInputDevice = null;
                    System.Diagnostics.Debug.WriteLine($"LoadAudioDevices: Error getting default microphone: {ex.Message}");
                }
                
                // Get all active output devices
                _outputDevices.AddRange(_deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active).ToList());
                
                // Get all active input devices
                _inputDevices.AddRange(_deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active).ToList());
                
                // Clear and repopulate combo boxes
                PopulateComboBoxes();
                
                // Restore selections or select defaults
                RestoreSelection(SecondaryHeadsetComboBox, selectedSecondaryHeadset, _outputDevices);
                RestoreSelection(SecondaryMicrophoneComboBox, selectedSecondaryMic, _inputDevices);
                
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
            
            // SECONDARY MICROPHONE
            SecondaryMicrophoneComboBox.Items.Clear();
            foreach (var device in _inputDevices)
            {
                SecondaryMicrophoneComboBox.Items.Add(device.FriendlyName);
            }
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
        }        private void Window_Closing(object sender, CancelEventArgs e)
        {
            // Stop any ongoing operations
            StopOutputFusion(); // Disposes its NAudio components
            StopInputFusion(); // Dispose input fusion components

            // Dispose individual top-level MMDevice references that might not be in the lists
            DisposeAudioDevice(_defaultOutputDevice); _defaultOutputDevice = null;
            DisposeAudioDevice(_defaultInputDevice); _defaultInputDevice = null;

            // Dispose all devices collected in the lists
            DisposeDeviceList(_outputDevices);
            DisposeDeviceList(_inputDevices);

            // Finally, dispose the MMDeviceEnumerator itself
            _deviceEnumerator?.Dispose();
            _deviceEnumerator = null;

            System.Diagnostics.Debug.WriteLine("AudioFusion closed and resources released.");
        }

        private void SecondaryVolumeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            SetSecondaryDeviceVolume((int)e.NewValue);
        }

        private void SetSecondaryDeviceVolume(int volume)
        {
            try
            {
                if (SecondaryHeadsetComboBox.SelectedItem == null) return;

                string selectedDeviceName = SecondaryHeadsetComboBox.SelectedItem.ToString();
                MMDevice secondaryDevice = FindDeviceByName(selectedDeviceName, _outputDevices);

                if (secondaryDevice != null)
                {
                    // Convert the volume from 0-100 scale to 0-1 scale
                    float volumeLevel = volume / 100f;
                    secondaryDevice.AudioEndpointVolume.MasterVolumeLevelScalar = volumeLevel;
                    StatusTextBlock.Text = $"Secondary headset volume: {volume}%";
                }
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = "Error setting volume";
                System.Diagnostics.Debug.WriteLine($"Error setting volume: {ex}");
            }
        }

        private void InputFusionButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isInputFusionRunning) StopInputFusion();
            else StartInputFusion();
        }

        private void StartInputFusion()
        {
            if (SecondaryMicrophoneComboBox.SelectedIndex == -1 || SecondaryMicrophoneComboBox.SelectedItem == null)
            {
                MessageBox.Show("Please select a secondary microphone device.", "Selection Missing", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            
            string selectedDeviceName = SecondaryMicrophoneComboBox.SelectedItem.ToString();
            MMDevice secondaryDevice = FindDeviceByName(selectedDeviceName, _inputDevices);
            
            if (secondaryDevice == null)
            {
                MessageBox.Show("Cannot find selected microphone device. It may have been disconnected. Please refresh.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            
            if (_defaultInputDevice == null) {
                MessageBox.Show("System default microphone is not available. Cannot start fusion.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (_defaultInputDevice.ID == secondaryDevice.ID)
            {
                MessageBox.Show("The secondary microphone cannot be the same as the system default microphone for this function.", "Device Conflict", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                // Capture from primary microphone (system default)
                _primaryMicrophoneCapture = new WasapiCapture(_defaultInputDevice);
                var primaryWaveFormat = _primaryMicrophoneCapture.WaveFormat;
                
                // Create buffered providers for both microphone inputs
                var primaryBufferedProvider = new BufferedWaveProvider(primaryWaveFormat)
                {
                    BufferDuration = TimeSpan.FromMilliseconds(100),
                    DiscardOnBufferOverflow = true
                };
                _primaryMicrophoneBuffer = primaryBufferedProvider;
                
                // Capture from secondary microphone
                _secondaryMicrophoneCapture = new WasapiCapture(secondaryDevice);
                
                // We need to ensure both sources use the same format for mixing
                // Create a resampler if necessary for the secondary mic
                var secondaryBufferedProvider = new BufferedWaveProvider(_secondaryMicrophoneCapture.WaveFormat)
                {
                    BufferDuration = TimeSpan.FromMilliseconds(100),
                    DiscardOnBufferOverflow = true
                };
                  // Set up a wave format converter if necessary
                IWaveProvider secondaryProvider = secondaryBufferedProvider;
                if (_secondaryMicrophoneCapture.WaveFormat.SampleRate != primaryWaveFormat.SampleRate ||
                    _secondaryMicrophoneCapture.WaveFormat.Channels != primaryWaveFormat.Channels)
                {
                    // If formats differ, convert to ISampleProvider and resample
                    var secondaryAsSampleProvider = ConvertToSampleProvider(secondaryBufferedProvider);
                    secondaryProvider = new WdlResamplingSampleProvider(secondaryAsSampleProvider, primaryWaveFormat.SampleRate).ToWaveProvider();
                }
                _secondaryMicrophoneBuffer = secondaryProvider;
                
                // Register data handlers for both microphone captures
                _primaryMicrophoneCapture.DataAvailable += (s, args) => 
                {
                    primaryBufferedProvider.AddSamples(args.Buffer, 0, args.BytesRecorded);
                };
                
                _secondaryMicrophoneCapture.DataAvailable += (s, args) => 
                {
                    secondaryBufferedProvider.AddSamples(args.Buffer, 0, args.BytesRecorded);
                };
                
                // Set up recording stopped handlers
                _primaryMicrophoneCapture.RecordingStopped += (s, args) =>
                {
                    Dispatcher.Invoke(() => 
                    {
                        if (_isInputFusionRunning)
                        {
                            StatusTextBlock.Text = "Primary microphone capture stopped unexpectedly.";
                            StopInputFusion();
                        }
                    });
                };
                
                _secondaryMicrophoneCapture.RecordingStopped += (s, args) =>
                {
                    Dispatcher.Invoke(() => 
                    {
                        if (_isInputFusionRunning)
                        {
                            StatusTextBlock.Text = "Secondary microphone capture stopped unexpectedly.";
                            StopInputFusion();
                        }
                    });
                };
                  // Create a mixer to combine both microphone inputs
                var primarySampleProvider = ConvertToSampleProvider(primaryBufferedProvider);
                var secondarySampleProviderForMixer = ConvertToSampleProvider(secondaryProvider);
                _inputMixer = new MixingSampleProvider(primarySampleProvider.WaveFormat);
                _inputMixer.AddMixerInput(primarySampleProvider);
                _inputMixer.AddMixerInput(secondarySampleProviderForMixer);
                // Set up the virtual output for the combined microphone streams
                _virtualAudioCaptureClient = new VirtualAudioCaptureClient(_inputMixer.ToWaveProvider());
                
                // Start recording from both microphones
                _primaryMicrophoneCapture.StartRecording();
                _secondaryMicrophoneCapture.StartRecording();
                
                // Start the virtual output
                _virtualAudioCaptureClient.Start();
                
                // Update UI
                _isInputFusionRunning = true;
                InputFusionButton.Content = "Stop Input Fusion";
                SecondaryMicrophoneComboBox.IsEnabled = false;
                RefreshDevicesButton.IsEnabled = false;
                StatusTextBlock.Text = $"Input Fusion: Combining '{_defaultInputDevice.FriendlyName}' with '{secondaryDevice.FriendlyName}'";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error starting input fusion: {ex.Message.Split('\n')[0]}", "Input Fusion Error", MessageBoxButton.OK, MessageBoxImage.Error);
                System.Diagnostics.Debug.WriteLine($"Error starting input fusion: {ex.ToString()}");
                StopInputFusion(); // Clean up anything that might have started
            }
        }

        private void StopInputFusion()
        {
            bool wasRunning = _isInputFusionRunning;
            _isInputFusionRunning = false; // Set this early to prevent re-entry from event handlers

            try
            {
                if (_primaryMicrophoneCapture != null)
                {
                    _primaryMicrophoneCapture.StopRecording();
                    _primaryMicrophoneCapture.Dispose();
                    _primaryMicrophoneCapture = null;
                }
                
                if (_secondaryMicrophoneCapture != null)
                {
                    _secondaryMicrophoneCapture.StopRecording();
                    _secondaryMicrophoneCapture.Dispose();
                    _secondaryMicrophoneCapture = null;
                }
                
                if (_virtualAudioCaptureClient != null)
                {
                    _virtualAudioCaptureClient.Stop();
                    _virtualAudioCaptureClient = null;
                }
                
                _primaryMicrophoneBuffer = null;
                _secondaryMicrophoneBuffer = null;
                _inputMixer = null;
                
                InputFusionButton.Content = "Start Input Fusion";
                SecondaryMicrophoneComboBox.IsEnabled = true;
                RefreshDevicesButton.IsEnabled = true;
                
                if (wasRunning && !StatusTextBlock.Text.ToLower().Contains("error") && !StatusTextBlock.Text.ToLower().Contains("unexpectedly"))
                {
                    StatusTextBlock.Text = "Input fusion stopped.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error stopping input fusion: {ex.Message.Split('\n')[0]}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                System.Diagnostics.Debug.WriteLine($"Error stopping input fusion: {ex.ToString()}");
                // Ensure UI state is reset even if an error occurs during stop
                InputFusionButton.Content = "Start Input Fusion";
                SecondaryMicrophoneComboBox.IsEnabled = true;
                RefreshDevicesButton.IsEnabled = true;
            }
        }

        // Custom method to convert any IWaveProvider to ISampleProvider safely
        private ISampleProvider ConvertToSampleProvider(IWaveProvider provider)
        {
            if (provider == null) throw new ArgumentNullException(nameof(provider));
            
            // Handle WaveStream directly
            if (provider is WaveStream waveStream)
            {
                return new SampleProvider(waveStream);
            }
            
            // For other IWaveProvider types including BufferedWaveProvider, use NAudio's built-in converter
            return new NAudio.Wave.SampleProviders.WaveToSampleProvider(provider);
        }
    }
}
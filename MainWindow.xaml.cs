using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;  // Add this line
using NAudio.CoreAudioApi;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;
using System.ComponentModel;

namespace AudioFusion
{
    public partial class MainWindow : Window
    {
        // Core components
        private MMDeviceEnumerator _deviceEnumerator;
        private List<MMDevice> _outputDevices = new List<MMDevice>();
        private List<MMDevice> _inputDevices = new List<MMDevice>();
        private MMDevice _defaultOutputDevice;
        
        // Output Fusion Components
        private WasapiLoopbackCapture _audioSourceCapture;
        private WasapiOut _secondaryHeadsetPlayer;
        private BufferedWaveProvider _outputFusionBuffer;
        private bool _isOutputFusionRunning = false;

        // Input Fusion Components
        private WasapiCapture _primaryMicCapture;
        private WasapiCapture _secondaryMicCapture;
        private WasapiOut _mixedAudioPlayer;
        private MixingSampleProvider _inputMixer;
        private BufferedWaveProvider _primaryMicBuffer;
        private BufferedWaveProvider _secondaryMicBuffer;
        private bool _isInputFusionRunning = false;

        public MainWindow()
        {
            InitializeComponent();
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
                string selectedPrimaryMic = PrimaryMicComboBox.SelectedItem?.ToString();
                string selectedSecondaryMic = SecondaryMicComboBox.SelectedItem?.ToString();
                string selectedMixedOutput = MixedAudioOutputDeviceComboBox.SelectedItem?.ToString();
                
                // Get default output device
                _defaultOutputDevice = _deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                DefaultOutputDeviceText.Text = $"{_defaultOutputDevice.FriendlyName} (System Default)";
                
                // Get output devices
                var outputs = _deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active).ToList();
                _outputDevices.AddRange(outputs);
                
                // Get input devices
                var inputs = _deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active).ToList();
                _inputDevices.AddRange(inputs);
                
                // Clear and repopulate combo boxes (with ComboBox.Items.Add for better performance)
                // SECONDARY HEADSET
                SecondaryHeadsetComboBox.Items.Clear();
                foreach (var device in _outputDevices)
                {
                    SecondaryHeadsetComboBox.Items.Add(device.FriendlyName);
                }
                
                // MIXED OUTPUT
                MixedAudioOutputDeviceComboBox.Items.Clear();
                foreach (var device in _outputDevices)
                {
                    MixedAudioOutputDeviceComboBox.Items.Add(device.FriendlyName);
                }
                
                // PRIMARY MIC
                PrimaryMicComboBox.Items.Clear();
                foreach (var device in _inputDevices)
                {
                    PrimaryMicComboBox.Items.Add(device.FriendlyName);
                }
                
                // SECONDARY MIC
                SecondaryMicComboBox.Items.Clear();
                foreach (var device in _inputDevices)
                {
                    SecondaryMicComboBox.Items.Add(device.FriendlyName);
                }
                
                // Restore selections or select defaults
                RestoreSelection(SecondaryHeadsetComboBox, selectedSecondaryHeadset);
                RestoreSelection(PrimaryMicComboBox, selectedPrimaryMic);
                RestoreSelection(SecondaryMicComboBox, selectedSecondaryMic);
                RestoreSelection(MixedAudioOutputDeviceComboBox, selectedMixedOutput);
                
                // Ensure different secondary mic if possible
                if (PrimaryMicComboBox.SelectedIndex == SecondaryMicComboBox.SelectedIndex && SecondaryMicComboBox.Items.Count > 1)
                {
                    SecondaryMicComboBox.SelectedIndex = (PrimaryMicComboBox.SelectedIndex + 1) % SecondaryMicComboBox.Items.Count;
                }
                
                StatusTextBlock.Text = "Audio devices loaded. Ready.";
                RefreshDevicesButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading audio devices: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusTextBlock.Text = "Error loading audio devices.";
            }
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

        private void InputFusionButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isInputFusionRunning) StopInputFusion();
            else StartInputFusion();
        }

        private void StartInputFusion()
        {
            if (PrimaryMicComboBox.SelectedIndex == -1 || SecondaryMicComboBox.SelectedIndex == -1 || MixedAudioOutputDeviceComboBox.SelectedIndex == -1)
            {
                MessageBox.Show("Please select all devices for input fusion.", "Selection Missing", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            
            // Get selected devices by name
            string primaryMicName = PrimaryMicComboBox.SelectedItem.ToString();
            string secondaryMicName = SecondaryMicComboBox.SelectedItem.ToString();
            string mixedOutputName = MixedAudioOutputDeviceComboBox.SelectedItem.ToString();
            
            // Find the actual device objects
            MMDevice primaryMic = _inputDevices.FirstOrDefault(d => d.FriendlyName == primaryMicName);
            MMDevice secondaryMic = _inputDevices.FirstOrDefault(d => d.FriendlyName == secondaryMicName);
            MMDevice mixedOutput = _outputDevices.FirstOrDefault(d => d.FriendlyName == mixedOutputName);
            
            if (primaryMic == null || secondaryMic == null || mixedOutput == null)
            {
                MessageBox.Show("One or more selected devices not found.", "Device Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            
            if (primaryMic.ID == secondaryMic.ID)
            {
                 MessageBoxResult result = MessageBox.Show("Primary and Secondary microphones are the same. Proceed?", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                 if (result == MessageBoxResult.No) return;
            }

            try
            {
                var stereoMixFormat = WaveFormat.CreateIeeeFloatWaveFormat(48000, 2);
                _inputMixer = new MixingSampleProvider(stereoMixFormat) { ReadFully = true };

                // Primary mic
                _primaryMicCapture = new WasapiCapture(primaryMic);
                _primaryMicBuffer = new BufferedWaveProvider(_primaryMicCapture.WaveFormat)
                { BufferDuration = TimeSpan.FromMilliseconds(100), DiscardOnBufferOverflow = true };
                
                // Convert to proper format
                ISampleProvider primarySampleProvider = new WaveToSampleProvider(_primaryMicBuffer);
                if (_primaryMicCapture.WaveFormat.SampleRate != stereoMixFormat.SampleRate)
                    primarySampleProvider = new WdlResamplingSampleProvider(primarySampleProvider, stereoMixFormat.SampleRate);
                if (primarySampleProvider.WaveFormat.Channels != stereoMixFormat.Channels)
                    primarySampleProvider = primarySampleProvider.WaveFormat.Channels == 1 ? 
                        (ISampleProvider)new MonoToStereoSampleProvider(primarySampleProvider) : 
                        new StereoToMonoSampleProvider(primarySampleProvider);
                
                _inputMixer.AddMixerInput(primarySampleProvider);
                _primaryMicCapture.DataAvailable += (s, args) => _primaryMicBuffer?.AddSamples(args.Buffer, 0, args.BytesRecorded);
                
                // Only add second mic if it's different (or same but user confirmed)
                if (primaryMic.ID != secondaryMic.ID || primaryMic.ID == secondaryMic.ID)
                {
                    _secondaryMicCapture = new WasapiCapture(secondaryMic);
                    _secondaryMicBuffer = new BufferedWaveProvider(_secondaryMicCapture.WaveFormat)
                    { BufferDuration = TimeSpan.FromMilliseconds(100), DiscardOnBufferOverflow = true };

                    // Convert to proper format
                    ISampleProvider secondarySampleProvider = new WaveToSampleProvider(_secondaryMicBuffer);
                    if (_secondaryMicCapture.WaveFormat.SampleRate != stereoMixFormat.SampleRate)
                        secondarySampleProvider = new WdlResamplingSampleProvider(secondarySampleProvider, stereoMixFormat.SampleRate);
                    if (secondarySampleProvider.WaveFormat.Channels != stereoMixFormat.Channels)
                        secondarySampleProvider = secondarySampleProvider.WaveFormat.Channels == 1 ? 
                            (ISampleProvider)new MonoToStereoSampleProvider(secondarySampleProvider) : 
                            new StereoToMonoSampleProvider(secondarySampleProvider);
                    
                    _inputMixer.AddMixerInput(secondarySampleProvider);
                    _secondaryMicCapture.DataAvailable += (s, args) => _secondaryMicBuffer?.AddSamples(args.Buffer, 0, args.BytesRecorded);
                }

                // Output the mixed audio
                _mixedAudioPlayer = new WasapiOut(mixedOutput, AudioClientShareMode.Shared, true, 100);
                _mixedAudioPlayer.Init(_inputMixer);

                // Start everything
                _primaryMicCapture.StartRecording();
                if (_secondaryMicCapture != null)
                    _secondaryMicCapture.StartRecording();
                _mixedAudioPlayer.Play();

                // Update UI
                _isInputFusionRunning = true;
                InputFusionButton.Content = "Stop Input Fusion";
                PrimaryMicComboBox.IsEnabled = false;
                SecondaryMicComboBox.IsEnabled = false;
                MixedAudioOutputDeviceComboBox.IsEnabled = false;
                
                string mic2Name = (_secondaryMicCapture != null) ? $" & '{secondaryMic.FriendlyName}'" : "";
                StatusTextBlock.Text = $"Input Fusion: Mixing '{primaryMic.FriendlyName}'{mic2Name} -> '{mixedOutput.FriendlyName}'";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error starting input fusion: {ex.Message}", "Input Fusion Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StopInputFusion();
            }
        }

        private void StopInputFusion()
        {
            try
            {
                if (_primaryMicCapture != null)
                {
                    _primaryMicCapture.StopRecording();
                    _primaryMicCapture.Dispose();
                    _primaryMicCapture = null;
                }
                
                if (_secondaryMicCapture != null)
                {
                    _secondaryMicCapture.StopRecording();
                    _secondaryMicCapture.Dispose();
                    _secondaryMicCapture = null;
                }
                
                if (_mixedAudioPlayer != null)
                {
                    _mixedAudioPlayer.Stop();
                    _mixedAudioPlayer.Dispose();
                    _mixedAudioPlayer = null;
                }
                
                _primaryMicBuffer = null;
                _secondaryMicBuffer = null;
                _inputMixer = null;
                
                _isInputFusionRunning = false;
                InputFusionButton.Content = "Start Input Fusion";
                PrimaryMicComboBox.IsEnabled = true;
                SecondaryMicComboBox.IsEnabled = true;
                MixedAudioOutputDeviceComboBox.IsEnabled = true;
                
                if (!StatusTextBlock.Text.Contains("Error") && !StatusTextBlock.Text.Contains("unexpectedly"))
                    StatusTextBlock.Text = "Input fusion stopped.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error stopping input fusion: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            StopOutputFusion();
            StopInputFusion();
            
            // No AudioDeviceItem wrappers to dispose in this version
            _deviceEnumerator?.Dispose();
            _deviceEnumerator = null;
        }
    }
}
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

namespace AudioFusion
{
    /// <summary>
    /// Simple resampling provider if WdlResamplingSampleProvider is not available
    /// </summary>
    public class SimpleSampleRateConverter : ISampleProvider
    {
        private readonly ISampleProvider source;
        private readonly WaveFormat targetFormat;
        private float[] sourceBuffer;
        
        public SimpleSampleRateConverter(ISampleProvider source, int newSampleRate)
        {
            this.source = source;
            this.targetFormat = new WaveFormat(newSampleRate, 
                source.WaveFormat.Channels,
                source.WaveFormat.BitsPerSample);
            
            sourceBuffer = new float[source.WaveFormat.SampleRate * source.WaveFormat.Channels];
        }
        
        public int Read(float[] buffer, int offset, int count)
        {
            // Simple nearest-neighbor resampling
            double ratio = (double)source.WaveFormat.SampleRate / targetFormat.SampleRate;
            int outputSamples = count / WaveFormat.Channels;
            
            int sourceSamplesNeeded = (int)(outputSamples * ratio);
            if (sourceBuffer.Length < sourceSamplesNeeded * WaveFormat.Channels)
                sourceBuffer = new float[sourceSamplesNeeded * WaveFormat.Channels];
            
            int sourceSamplesRead = source.Read(sourceBuffer, 0, sourceSamplesNeeded * WaveFormat.Channels);
            int outputSamplesWritten = (int)(sourceSamplesRead / (ratio * WaveFormat.Channels));
            
            for (int outputSample = 0; outputSample < outputSamplesWritten; outputSample++)
            {
                int inputSample = (int)(outputSample * ratio);
                for (int channel = 0; channel < WaveFormat.Channels; channel++)
                {
                    buffer[offset + outputSample * WaveFormat.Channels + channel] = 
                        sourceBuffer[inputSample * WaveFormat.Channels + channel];
                }
            }
            
            return outputSamplesWritten * WaveFormat.Channels;
        }
        
        public WaveFormat WaveFormat => targetFormat;
    }
    
    public static class AudioExtensions
    {
        public static VolumeSampleProvider VolumeSampleProvider(this ISampleProvider provider, float volume)
        {
            return new VolumeSampleProvider(provider) { Volume = volume };
        }
        
        // Helper methods for audio format conversion
        public static ISampleProvider ResampleTo(this ISampleProvider provider, int sampleRate)
        {
            if (provider.WaveFormat.SampleRate == sampleRate)
                return provider;
                
            try
            {
                // Try to use WdlResamplingSampleProvider if available
                var resamplerType = Type.GetType("NAudio.Wave.SampleProviders.WdlResamplingSampleProvider, NAudio");
                if (resamplerType != null)
                {
                    var constructor = resamplerType.GetConstructor(new[] { 
                        typeof(ISampleProvider), typeof(int) 
                    });
                    
                    return (ISampleProvider)constructor.Invoke(new object[] { provider, sampleRate });
                }
                else
                {
                    // Fallback to our simple resampler
                    return new SimpleSampleRateConverter(provider, sampleRate);
                }
            }
            catch
            {
                // Fallback to our simple resampler
                return new SimpleSampleRateConverter(provider, sampleRate);
            }
        }
    }

    public partial class MainWindow : Window
    {
        // Core components
        private MMDeviceEnumerator _deviceEnumerator;
        private List<MMDevice> _outputDevices = new List<MMDevice>();
        private MMDevice _defaultOutputDevice;
        
        // Output Fusion Components
        private WasapiLoopbackCapture _audioSourceCapture;
        private WasapiOut _secondaryHeadsetPlayer;
        private BufferedWaveProvider _outputFusionBuffer;
        private bool _isOutputFusionRunning = false;

    // Input Fusion Components
        private List<MMDevice> _inputDevices = new List<MMDevice>();
        private WasapiCapture _primaryMicCapture;
        private WasapiCapture _secondaryMicCapture;
        private BufferedWaveProvider _primaryMicBuffer;
        private BufferedWaveProvider _secondaryMicBuffer;
        private WaveOut _inputFusionEchoPlayer;
        private IWavePlayer _secondaryToLoopbackPlayer;
        private MixingSampleProvider _inputFusionMixer;
        private bool _isInputFusionRunning = false;
        private string _originalDefaultMicId;

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

                // Dispose existing MMDevice objects before re-populating
                DisposeDeviceList(_outputDevices);
                DisposeAudioDevice(_defaultOutputDevice); _defaultOutputDevice = null;
                
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
                
                // Get all active output devices
                _outputDevices.AddRange(_deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active).ToList());
                
                // Clear and repopulate combo boxes
                PopulateComboBoxes();
                
                // Restore selections or select defaults
                RestoreSelection(SecondaryHeadsetComboBox, selectedSecondaryHeadset, _outputDevices);
                
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
            if (_inputDevices == null) _inputDevices = new List<MMDevice>();
            _inputDevices.Clear();
            _inputDevices.AddRange(_deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active).ToList());
            PrimaryMicrophoneComboBox.Items.Clear();
            SecondaryMicrophoneComboBox.Items.Clear();
            foreach (var device in _inputDevices)
            {
                PrimaryMicrophoneComboBox.Items.Add(device.FriendlyName);
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
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            // Stop any ongoing operations
            StopOutputFusion(); // Disposes its NAudio components

            // Dispose individual top-level MMDevice references that might not be in the lists
            DisposeAudioDevice(_defaultOutputDevice); _defaultOutputDevice = null;

            // Dispose all devices collected in the lists
            DisposeDeviceList(_outputDevices);

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
        }        private void StartInputFusion()
        {
            if (PrimaryMicrophoneComboBox.SelectedIndex == -1 || SecondaryMicrophoneComboBox.SelectedIndex == -1)
            {
                MessageBox.Show("Please select both primary and secondary microphones.", "Selection Missing", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (PrimaryMicrophoneComboBox.SelectedIndex == SecondaryMicrophoneComboBox.SelectedIndex)
            {
                MessageBox.Show("Primary and secondary microphones must be different.", "Device Conflict", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            
            var primaryDevice = _inputDevices[PrimaryMicrophoneComboBox.SelectedIndex];
            var secondaryDevice = _inputDevices[SecondaryMicrophoneComboBox.SelectedIndex];
            
            // Store the original default microphone ID to restore later
            try
            {
                // Get the current default communication device to restore later
                var defaultCommunicationDevice = _deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications);
                _originalDefaultMicId = defaultCommunicationDevice?.ID;
                defaultCommunicationDevice?.Dispose();
                
                // Set the primary microphone as the default communication device
                if (!AudioEndpointManager.SetDefaultAudioEndpoint(primaryDevice.ID, Role.Communications))
                {
                    MessageBox.Show("Failed to set primary microphone as the default communication device. Input Fusion may not work correctly with Teams.", 
                                   "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                
                _primaryMicCapture = new WasapiCapture(primaryDevice);
                _secondaryMicCapture = new WasapiCapture(secondaryDevice);
                _primaryMicBuffer = new BufferedWaveProvider(_primaryMicCapture.WaveFormat);
                _secondaryMicBuffer = new BufferedWaveProvider(_secondaryMicCapture.WaveFormat);
                
                _primaryMicCapture.DataAvailable += (s, a) => _primaryMicBuffer.AddSamples(a.Buffer, 0, a.BytesRecorded);
                _secondaryMicCapture.DataAvailable += (s, a) => _secondaryMicBuffer.AddSamples(a.Buffer, 0, a.BytesRecorded);                // Convert to ISampleProvider for mixing
                var primarySample = _primaryMicBuffer.ToSampleProvider();
                var secondarySample = _secondaryMicBuffer.ToSampleProvider();
                
                // Get and print the formats for debugging
                var format1 = primarySample.WaveFormat;
                var format2 = secondarySample.WaveFormat;
                
                System.Diagnostics.Debug.WriteLine($"Primary format: {format1.SampleRate}Hz, {format1.Channels} channels, {format1.Encoding}");
                System.Diagnostics.Debug.WriteLine($"Secondary format: {format2.SampleRate}Hz, {format2.Channels} channels, {format2.Encoding}");
                  // Convert both to the same format before mixing (standardize on the primary format)
                var primaryWithVolume = primarySample.VolumeSampleProvider((float)(PrimaryMicVolumeSlider.Value / 100.0));
                
                // Convert the second sample to match the first one's format if needed
                ISampleProvider secondaryWithFormat;
                if (!format1.Equals(format2))
                {
                    // Resample to match sample rate
                    secondaryWithFormat = secondarySample.ResampleTo(format1.SampleRate);
                    
                    // Match channel count if needed
                    if (format1.Channels != format2.Channels)
                    {
                        // If channel count differs, convert to mono or stereo as needed
                        try {
                            if (format1.Channels == 1 && secondaryWithFormat.WaveFormat.Channels == 2)
                                secondaryWithFormat = new StereoToMonoSampleProvider(secondaryWithFormat);
                            else if (format1.Channels == 2 && secondaryWithFormat.WaveFormat.Channels == 1)
                                secondaryWithFormat = new MonoToStereoSampleProvider(secondaryWithFormat);
                        }
                        catch (Exception ex) {
                            System.Diagnostics.Debug.WriteLine($"Error converting channels: {ex.Message}");
                            // If channel conversion fails, try to create a format with matching channels
                            var newFormat = new WaveFormat(format1.SampleRate, format1.Channels);
                            MessageBox.Show($"Warning: Input devices have different formats. Audio quality may be affected.", "Format Mismatch", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }
                }
                else
                {
                    secondaryWithFormat = secondarySample;
                }
                
                // Apply volume after format conversion
                var secondaryWithVolume = secondaryWithFormat.VolumeSampleProvider((float)(SecondaryMicVolumeSlider.Value / 100.0));
                
                // Create mixer with the primary format
                _inputFusionMixer = new MixingSampleProvider(format1);
                _inputFusionMixer.AddMixerInput(primaryWithVolume);
                _inputFusionMixer.AddMixerInput(secondaryWithVolume);
                _inputFusionMixer.ReadFully = true;                // Create a secondary output to the primary mic's loopback if possible
                try
                {
                    var secondarySampleProvider = _secondaryMicBuffer.ToSampleProvider();
                    
                    // Apply volume
                    var volumeProvider = secondarySampleProvider.VolumeSampleProvider((float)(SecondaryMicVolumeSlider.Value / 100.0));
                    
                    // Format conversion will happen inside the CreateLoopbackPlayer method
                    _secondaryToLoopbackPlayer = CreateLoopbackPlayer(primaryDevice, volumeProvider);
                    
                    if (_secondaryToLoopbackPlayer != null)
                    {
                        _secondaryToLoopbackPlayer.Play();
                        System.Diagnostics.Debug.WriteLine("Secondary microphone successfully routed to primary device loopback");
                    }
                }
                catch (Exception loopbackEx)
                {
                    System.Diagnostics.Debug.WriteLine($"Loopback not supported: {loopbackEx.Message}");
                    // Just continue, this is an enhancement if available
                }
                
                // Set up echo monitoring (hearing the mix in the headphones) if enabled
                if (EchoMonitoringCheckBox.IsChecked == true)
                {
                    _inputFusionEchoPlayer = new WaveOut();
                    _inputFusionEchoPlayer.Init(_inputFusionMixer);
                    _inputFusionEchoPlayer.Play();
                }
                
                // Start recording from both microphones
                _primaryMicCapture.StartRecording();
                _secondaryMicCapture.StartRecording();
                
                PrimaryMicrophoneComboBox.IsEnabled = false;
                SecondaryMicrophoneComboBox.IsEnabled = false;
                PrimaryMicVolumeSlider.IsEnabled = false; // Real-time volume changes are complicated
                SecondaryMicVolumeSlider.IsEnabled = false;
                EchoMonitoringCheckBox.IsEnabled = false;
                
                _isInputFusionRunning = true;
                InputFusionButton.Content = "Stop Input Fusion";
                
                StatusTextBlock.Text = $"Input Fusion: Mixing '{primaryDevice.FriendlyName}' with '{secondaryDevice.FriendlyName}'";
                InputFusionStatusText.Text = $"Input Fusion active. Using '{primaryDevice.FriendlyName}' as Teams microphone with mixed audio. Echo monitoring: {(EchoMonitoringCheckBox.IsChecked == true ? "ON" : "OFF")}.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error starting input fusion: {ex.Message}", "Input Fusion Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StopInputFusion();
            }
        }
          // Helper method to create a loopback player if possible
        private IWavePlayer CreateLoopbackPlayer(MMDevice targetDevice, ISampleProvider sourceProvider)
        {
            try
            {
                // When feeding to WasapiOut, we need to ensure format compatibility
                // First try using the device's preferred format if we can determine it
                WasapiOut player;
                
                try
                {
                    // Try to create a WasapiOut with the device's format
                    player = new WasapiOut(targetDevice, AudioClientShareMode.Shared, false, 100);
                    
                    // Get device's preferred format if possible
                    var deviceFormat = player.OutputWaveFormat;
                    System.Diagnostics.Debug.WriteLine($"Output device format: {deviceFormat.SampleRate}Hz, {deviceFormat.Channels} channels");
                    
                    // Check if we need to resample
                    if (deviceFormat.SampleRate != sourceProvider.WaveFormat.SampleRate ||
                        deviceFormat.Channels != sourceProvider.WaveFormat.Channels)
                    {                        // Need to adapt the format
                        var resampled = sourceProvider.ResampleTo(deviceFormat.SampleRate);
                        
                        // Convert channels if needed
                        ISampleProvider channelConverter = resampled;
                        if (deviceFormat.Channels != resampled.WaveFormat.Channels)
                        {
                            try {
                                if (deviceFormat.Channels == 1 && resampled.WaveFormat.Channels == 2)
                                    channelConverter = new StereoToMonoSampleProvider(resampled);
                                else if (deviceFormat.Channels == 2 && resampled.WaveFormat.Channels == 1)
                                    channelConverter = new MonoToStereoSampleProvider(resampled);
                            } 
                            catch (Exception ex) {
                                System.Diagnostics.Debug.WriteLine($"Error converting channels for loopback: {ex.Message}");
                                // If conversion fails, just use the resampled version and hope for the best
                            }
                        }
                        
                        player.Init(channelConverter);
                    }
                    else
                    {
                        // Format is already compatible
                        player.Init(sourceProvider);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error creating output with device format: {ex.Message}");
                    // Fallback to a standard format
                    player = new WasapiOut(targetDevice, AudioClientShareMode.Shared, false, 100);
                    player.Init(sourceProvider);
                }
                
                return player;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Loopback not supported: {ex.Message}");
                return null; // Loopback not supported on this device
            }
        }private void StopInputFusion()
        {
            _isInputFusionRunning = false;
            try
            {
                _primaryMicCapture?.StopRecording();
                _primaryMicCapture?.Dispose();
                _primaryMicCapture = null;
                _secondaryMicCapture?.StopRecording();
                _secondaryMicCapture?.Dispose();
                _secondaryMicCapture = null;
                _primaryMicBuffer = null;
                _secondaryMicBuffer = null;
                _inputFusionEchoPlayer?.Stop();
                _inputFusionEchoPlayer?.Dispose();
                _inputFusionEchoPlayer = null;
                _secondaryToLoopbackPlayer?.Stop();
                _secondaryToLoopbackPlayer?.Dispose();
                _secondaryToLoopbackPlayer = null;
                _inputFusionMixer = null;
                InputFusionButton.Content = "Start Input Fusion";
                InputFusionStatusText.Text = "Input fusion stopped.";
            }
            catch (Exception ex)
            {
                InputFusionStatusText.Text = "Error stopping input fusion.";
                System.Diagnostics.Debug.WriteLine($"Error stopping input fusion: {ex}");
            }
        }

        // Helper for volume
        private void PrimaryMicVolumeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            // No real-time volume change for now; restart fusion to apply
        }
        private void SecondaryMicVolumeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            // No real-time volume change for now; restart fusion to apply
        }
    }
}
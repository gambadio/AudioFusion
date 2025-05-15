#nullable enable // Enable nullable reference types for this file

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using NAudio.CoreAudioApi;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;
using System.ComponentModel;

namespace AudioFusion
{
    public partial class MainWindow : Window
    {
        private MMDeviceEnumerator? _deviceEnumerator;

        // Output Fusion Components
        private MMDevice? _defaultOutputDevice;
        private MMDevice? _selectedSecondaryHeadsetDevice;
        private WasapiLoopbackCapture? _audioSourceCapture;
        private WasapiOut? _secondaryHeadsetPlayer;
        private BufferedWaveProvider? _outputFusionBuffer;
        private bool _isOutputFusionRunning = false;

        // Input Fusion Components
        private MMDevice? _selectedPrimaryMicDevice;
        private MMDevice? _selectedSecondaryMicDevice;
        private MMDevice? _selectedMixedAudioOutputDevice;
        private WasapiCapture? _primaryMicCapture;
        private WasapiCapture? _secondaryMicCapture;
        private WasapiOut? _mixedAudioPlayer;
        private MixingSampleProvider? _inputMixer;
        private BufferedWaveProvider? _primaryMicBuffer;
        private BufferedWaveProvider? _secondaryMicBuffer;
        private bool _isInputFusionRunning = false;

        public MainWindow()
        {
            InitializeComponent(); // This call is essential and should be the first thing
            _deviceEnumerator = new MMDeviceEnumerator();
            LoadAudioDevices();
        }

        private void LoadAudioDevices()
        {
            if (_deviceEnumerator == null) return; 

            try
            {
                _defaultOutputDevice = _deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                DefaultOutputDeviceText.Text = $"{_defaultOutputDevice.FriendlyName} (System Default)";

                var activeOutputDevices = _deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active).ToList();

                SecondaryHeadsetComboBox.Items.Clear();
                foreach (var device in activeOutputDevices)
                {
                    if (device.ID != _defaultOutputDevice.ID || activeOutputDevices.Count == 1)
                    {
                        SecondaryHeadsetComboBox.Items.Add(new AudioDeviceItem(device));
                    }
                }
                if (SecondaryHeadsetComboBox.Items.Count > 0)
                    SecondaryHeadsetComboBox.SelectedIndex = 0;

                MixedAudioOutputDeviceComboBox.Items.Clear();
                foreach (var device in activeOutputDevices)
                {
                     MixedAudioOutputDeviceComboBox.Items.Add(new AudioDeviceItem(device));
                }
                if (MixedAudioOutputDeviceComboBox.Items.Count > 1)
                {
                    int preferredIndex = MixedAudioOutputDeviceComboBox.Items.Cast<AudioDeviceItem>().ToList().FindIndex(d => d.Device.ID != _defaultOutputDevice.ID);
                    MixedAudioOutputDeviceComboBox.SelectedIndex = (preferredIndex != -1) ? preferredIndex : 0;
                }
                else if (MixedAudioOutputDeviceComboBox.Items.Count > 0)
                {
                    MixedAudioOutputDeviceComboBox.SelectedIndex = 0;
                }

                var activeInputDevices = _deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active).ToList();

                PrimaryMicComboBox.Items.Clear();
                SecondaryMicComboBox.Items.Clear();

                foreach (var device in activeInputDevices)
                {
                    PrimaryMicComboBox.Items.Add(new AudioDeviceItem(device));
                    SecondaryMicComboBox.Items.Add(new AudioDeviceItem(device));
                }

                if (PrimaryMicComboBox.Items.Count > 0)
                    PrimaryMicComboBox.SelectedIndex = 0;

                if (SecondaryMicComboBox.Items.Count > 1)
                    SecondaryMicComboBox.SelectedIndex = 1;
                else if (SecondaryMicComboBox.Items.Count > 0)
                    SecondaryMicComboBox.SelectedIndex = 0;

                StatusTextBlock.Text = "Audio devices loaded. Ready.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error initializing audio devices: {ex.Message}\n\nMake sure you have playback and recording devices enabled and connected.", "Audio Initialization Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusTextBlock.Text = "Error initializing audio devices.";
                OutputFusionButton.IsEnabled = false;
                InputFusionButton.IsEnabled = false;
            }
        }

        private void OutputFusionButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isOutputFusionRunning) StopOutputFusion();
            else StartOutputFusion();
        }

        private void StartOutputFusion()
        {
            if (SecondaryHeadsetComboBox.SelectedItem == null || _selectedSecondaryHeadsetDevice == null) // Added check for _selectedSecondaryHeadsetDevice
            {
                MessageBox.Show("Please select a secondary headset device.", "Selection Missing", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (_defaultOutputDevice == null)
            {
                MessageBox.Show("Default output device not found. Cannot start output fusion.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                LoadAudioDevices(); 
                return;
            }

            // This was already correct from selection, but ensuring it's not null if we proceed
             _selectedSecondaryHeadsetDevice = ((AudioDeviceItem)SecondaryHeadsetComboBox.SelectedItem).Device;

            if (_defaultOutputDevice.ID == _selectedSecondaryHeadsetDevice.ID)
            {
                 MessageBox.Show("The secondary headset cannot be the same as the system default output for this function.", "Device Conflict", MessageBoxButton.OK, MessageBoxImage.Warning);
                 return;
            }

            try
            {
                _audioSourceCapture = new WasapiLoopbackCapture(_defaultOutputDevice);
                _outputFusionBuffer = new BufferedWaveProvider(_audioSourceCapture.WaveFormat)
                {
                    BufferDuration = TimeSpan.FromMilliseconds(200),
                    DiscardOnBufferOverflow = true
                };

                _audioSourceCapture.DataAvailable += (s, args) => _outputFusionBuffer?.AddSamples(args.Buffer, 0, args.BytesRecorded);
                _audioSourceCapture.RecordingStopped += (s, args) =>
                {
                    if (_isOutputFusionRunning) Dispatcher.Invoke(() => {
                        StopOutputFusion();
                        StatusTextBlock.Text = "Output fusion stopped unexpectedly.";
                    });
                };

                // ***** CORRECTED WasapiOut CONSTRUCTOR *****
                _secondaryHeadsetPlayer = new WasapiOut(_selectedSecondaryHeadsetDevice, AudioClientShareMode.Shared, true, 100); // Added 'true' for useEventSync
                // ******************************************

                if (_outputFusionBuffer != null) // Null check before Init
                {
                    _secondaryHeadsetPlayer.Init(_outputFusionBuffer);
                }
                else
                {
                    throw new InvalidOperationException("Output fusion buffer was not initialized.");
                }

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
                StatusTextBlock.Text = $"Output Fusion: Mirroring '{_defaultOutputDevice.FriendlyName}' to '{_selectedSecondaryHeadsetDevice.FriendlyName}'";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error starting output fusion: {ex.Message}", "Output Fusion Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StopOutputFusion();
            }
        }

        private void StopOutputFusion()
        {
            _audioSourceCapture?.StopRecording();
            _audioSourceCapture?.Dispose();
            _audioSourceCapture = null;

            _secondaryHeadsetPlayer?.Stop();
            _secondaryHeadsetPlayer?.Dispose();
            _secondaryHeadsetPlayer = null;

            _outputFusionBuffer = null;

            _isOutputFusionRunning = false;
            OutputFusionButton.Content = "Start Output Fusion";
            SecondaryHeadsetComboBox.IsEnabled = true;
            if (!StatusTextBlock.Text.Contains("Error") && !StatusTextBlock.Text.Contains("unexpectedly")) StatusTextBlock.Text = "Output fusion stopped.";
        }

        private void InputFusionButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isInputFusionRunning) StopInputFusion();
            else StartInputFusion();
        }

        private void StartInputFusion()
        {
            if (PrimaryMicComboBox.SelectedItem == null || SecondaryMicComboBox.SelectedItem == null || MixedAudioOutputDeviceComboBox.SelectedItem == null)
            {
                MessageBox.Show("Please select both microphones and the mixed audio output device.", "Selection Missing", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _selectedPrimaryMicDevice = ((AudioDeviceItem)PrimaryMicComboBox.SelectedItem).Device;
            _selectedSecondaryMicDevice = ((AudioDeviceItem)SecondaryMicComboBox.SelectedItem).Device;
            _selectedMixedAudioOutputDevice = ((AudioDeviceItem)MixedAudioOutputDeviceComboBox.SelectedItem).Device;

            if (_selectedPrimaryMicDevice == null || _selectedSecondaryMicDevice == null || _selectedMixedAudioOutputDevice == null)
            {
                 MessageBox.Show("One or more selected audio devices are invalid. Please reselect.", "Device Error", MessageBoxButton.OK, MessageBoxImage.Error);
                 LoadAudioDevices(); // Refresh device lists
                 return;
            }

            if (_selectedPrimaryMicDevice.ID == _selectedSecondaryMicDevice.ID)
            {
                 MessageBoxResult result = MessageBox.Show("Primary and Secondary microphones are the same. Do you want to proceed?", "Device Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                 if (result == MessageBoxResult.No) return;
            }

            try
            {
                var stereoMixFormat = WaveFormat.CreateIeeeFloatWaveFormat(48000, 2);
                _inputMixer = new MixingSampleProvider(stereoMixFormat) { ReadFully = true };

                _primaryMicCapture = new WasapiCapture(_selectedPrimaryMicDevice);
                _primaryMicBuffer = new BufferedWaveProvider(_primaryMicCapture.WaveFormat)
                { BufferDuration = TimeSpan.FromMilliseconds(100), DiscardOnBufferOverflow = true };
                
                ISampleProvider primarySampleProvider = new WaveToSampleProvider(_primaryMicBuffer); 
                if (_primaryMicCapture.WaveFormat.SampleRate != stereoMixFormat.SampleRate || _primaryMicCapture.WaveFormat.Channels != stereoMixFormat.Channels)
                {
                    primarySampleProvider = new WdlResamplingSampleProvider(primarySampleProvider, stereoMixFormat.SampleRate);
                    if (primarySampleProvider.WaveFormat.Channels != stereoMixFormat.Channels)
                    {
                        primarySampleProvider = stereoMixFormat.Channels == 1 ?
                            (ISampleProvider)new StereoToMonoSampleProvider(primarySampleProvider) :
                            new MonoToStereoSampleProvider(primarySampleProvider);
                    }
                }
                _inputMixer.AddMixerInput(primarySampleProvider);
                _primaryMicCapture.DataAvailable += (s, args) => _primaryMicBuffer?.AddSamples(args.Buffer, 0, args.BytesRecorded);
                _primaryMicCapture.RecordingStopped += (s, args) => { if (_isInputFusionRunning) Dispatcher.Invoke(StopInputFusion); };

                if (_selectedPrimaryMicDevice.ID != _selectedSecondaryMicDevice.ID)
                {
                    _secondaryMicCapture = new WasapiCapture(_selectedSecondaryMicDevice);
                    _secondaryMicBuffer = new BufferedWaveProvider(_secondaryMicCapture.WaveFormat)
                    { BufferDuration = TimeSpan.FromMilliseconds(100), DiscardOnBufferOverflow = true };

                    ISampleProvider secondarySampleProvider = new WaveToSampleProvider(_secondaryMicBuffer); 
                    if (_secondaryMicCapture.WaveFormat.SampleRate != stereoMixFormat.SampleRate || _secondaryMicCapture.WaveFormat.Channels != stereoMixFormat.Channels)
                    {
                        secondarySampleProvider = new WdlResamplingSampleProvider(secondarySampleProvider, stereoMixFormat.SampleRate);
                        if (secondarySampleProvider.WaveFormat.Channels != stereoMixFormat.Channels)
                        {
                             secondarySampleProvider = stereoMixFormat.Channels == 1 ?
                                (ISampleProvider)new StereoToMonoSampleProvider(secondarySampleProvider) :
                                new MonoToStereoSampleProvider(secondarySampleProvider);
                        }
                    }
                    _inputMixer.AddMixerInput(secondarySampleProvider);
                    _secondaryMicCapture.DataAvailable += (s, args) => _secondaryMicBuffer?.AddSamples(args.Buffer, 0, args.BytesRecorded);
                    _secondaryMicCapture.RecordingStopped += (s, args) => { if (_isInputFusionRunning) Dispatcher.Invoke(StopInputFusion); };
                }

                // ***** CORRECTED WasapiOut CONSTRUCTOR *****
                _mixedAudioPlayer = new WasapiOut(_selectedMixedAudioOutputDevice, AudioClientShareMode.Shared, true, 100); // Added 'true' for useEventSync
                // ******************************************
                
                if (_inputMixer != null) // Null check before Init
                {
                     _mixedAudioPlayer.Init(_inputMixer);
                }
                else
                {
                    throw new InvalidOperationException("Input mixer was not initialized.");
                }

                _mixedAudioPlayer.PlaybackStopped += (s, args) =>
                {
                    if (args.Exception != null && _isInputFusionRunning)
                    {
                        Dispatcher.Invoke(() => {
                            StatusTextBlock.Text = $"Mixed audio playback error: {args.Exception.Message}";
                            StopInputFusion();
                        });
                    }
                };

                _primaryMicCapture.StartRecording();
                _secondaryMicCapture?.StartRecording(); 
                _mixedAudioPlayer.Play();

                _isInputFusionRunning = true;
                InputFusionButton.Content = "Stop Input Fusion";
                PrimaryMicComboBox.IsEnabled = false;
                SecondaryMicComboBox.IsEnabled = false;
                MixedAudioOutputDeviceComboBox.IsEnabled = false;
                string mic2Name = (_secondaryMicCapture != null && _selectedSecondaryMicDevice != null) ? $" & '{_selectedSecondaryMicDevice.FriendlyName}'" : "";
                StatusTextBlock.Text = $"Input Fusion: Mixing '{_selectedPrimaryMicDevice.FriendlyName}'{mic2Name} -> '{_selectedMixedAudioOutputDevice.FriendlyName}'";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error starting input fusion: {ex.Message}", "Input Fusion Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StopInputFusion();
            }
        }

        private void StopInputFusion()
        {
            _primaryMicCapture?.StopRecording();
            _primaryMicCapture?.Dispose();
            _primaryMicCapture = null;
            _primaryMicBuffer = null;

            _secondaryMicCapture?.StopRecording();
            _secondaryMicCapture?.Dispose();
            _secondaryMicCapture = null;
            _secondaryMicBuffer = null;

            _mixedAudioPlayer?.Stop();
            _mixedAudioPlayer?.Dispose();
            _mixedAudioPlayer = null;

            _inputMixer = null;

            _isInputFusionRunning = false;
            InputFusionButton.Content = "Start Input Fusion";
            PrimaryMicComboBox.IsEnabled = true;
            SecondaryMicComboBox.IsEnabled = true;
            MixedAudioOutputDeviceComboBox.IsEnabled = true;
            if (!StatusTextBlock.Text.Contains("Error") && !StatusTextBlock.Text.Contains("unexpectedly")) StatusTextBlock.Text = "Input fusion stopped.";
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            StopOutputFusion();
            StopInputFusion();
            _deviceEnumerator?.Dispose();
            _deviceEnumerator = null;
        }
    }

    public class AudioDeviceItem
    {
        public MMDevice Device { get; }
        public string Name => Device.FriendlyName;

        public AudioDeviceItem(MMDevice device)
        {
            Device = device ?? throw new ArgumentNullException(nameof(device));
        }

        public override string ToString() => Name;
    }
}
#nullable disable
using System;
using System.Threading;
using NAudio.Wave;

namespace AudioFusion
{
    public class VirtualAudioCaptureClient
    {
        private IWaveProvider _sourceProvider;
        private WaveFormat _waveFormat;
        private bool _isRunning;
        private Thread _processingThread;
        private WaveOut _outputDevice; // Virtual microphone output
        
        public event EventHandler<DataAvailableEventArgs> DataAvailable;

        public VirtualAudioCaptureClient(IWaveProvider sourceProvider)
        {
            _sourceProvider = sourceProvider ?? throw new ArgumentNullException(nameof(sourceProvider));
            _waveFormat = sourceProvider.WaveFormat;
            _outputDevice = new WaveOut();
        }

        public void Start()
        {
            if (_isRunning)
                return;

            _isRunning = true;
            _processingThread = new Thread(ProcessingThread)
            {
                IsBackground = true,
                Name = "Virtual Audio Capture Thread",
                Priority = ThreadPriority.AboveNormal
            };
            _processingThread.Start();

            // Initialize the virtual output device
            _outputDevice.Init(_sourceProvider);
            _outputDevice.Play();
        }

        public void Stop()
        {
            if (!_isRunning)
                return;

            _isRunning = false;
            
            // Stop the output device
            if (_outputDevice != null)
            {
                _outputDevice.Stop();
                _outputDevice.Dispose();
                _outputDevice = null;
            }
            
            // Wait for the processing thread to finish
            if (_processingThread != null)
            {
                if (_processingThread.IsAlive && !_processingThread.Join(1000))
                {
                    try
                    {
                        _processingThread.Abort();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error aborting processing thread: {ex.Message}");
                    }
                }
                _processingThread = null;
            }
        }

        private void ProcessingThread()
        {
            int bufferSize = _waveFormat.AverageBytesPerSecond / 10; // 100ms buffer
            byte[] buffer = new byte[bufferSize];
            
            while (_isRunning)
            {
                if (_sourceProvider != null)
                {
                    int bytesRead = _sourceProvider.Read(buffer, 0, buffer.Length);
                    if (bytesRead > 0)
                    {
                        // Buffer was filled, raise the event
                        DataAvailable?.Invoke(this, new DataAvailableEventArgs(buffer, bytesRead));
                    }
                    else
                    {
                        // No data available, sleep a bit to avoid CPU spinning
                        Thread.Sleep(10);
                    }
                }
                else
                {
                    Thread.Sleep(10);
                }
            }
        }

        public class DataAvailableEventArgs : EventArgs
        {
            public DataAvailableEventArgs(byte[] buffer, int bytesRecorded)
            {
                Buffer = buffer;
                BytesRecorded = bytesRecorded;
            }

            public byte[] Buffer { get; }
            public int BytesRecorded { get; }
        }
    }
}

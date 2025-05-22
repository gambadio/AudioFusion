using System;
using NAudio.Wave;

namespace AudioFusion
{
    /// <summary>
    /// Converts from an ISampleProvider to an IWaveProvider
    /// </summary>
    public class SampleToWaveProvider : IWaveProvider
    {
        private readonly ISampleProvider _source;
        private readonly byte[] _sourceBuffer;
        private readonly int _bytesPerSample;
        
        /// <summary>
        /// Initializes a new instance of the SampleToWaveProvider class
        /// </summary>
        /// <param name="source">Source sample provider</param>
        public SampleToWaveProvider(ISampleProvider source)
        {
            _source = source ?? throw new ArgumentNullException(nameof(source));
            WaveFormat = source.WaveFormat;
            _bytesPerSample = 4; // IEEE float
            _sourceBuffer = new byte[source.WaveFormat.SampleRate * source.WaveFormat.Channels * _bytesPerSample];
        }

        /// <summary>
        /// Gets the WaveFormat of this IWaveProvider
        /// </summary>
        public WaveFormat WaveFormat { get; }

        /// <summary>
        /// Reads from this provider
        /// </summary>
        public int Read(byte[] buffer, int offset, int count)
        {
            if (buffer == null) throw new ArgumentNullException(nameof(buffer));

            int sampleCount = count / _bytesPerSample;
            float[] sampleBuffer = new float[sampleCount];
            int samplesRead = _source.Read(sampleBuffer, 0, sampleCount);
            
            // convert to bytes
            for (int sample = 0; sample < samplesRead; sample++)
            {
                byte[] bytes = BitConverter.GetBytes(sampleBuffer[sample]);
                for (int b = 0; b < _bytesPerSample; b++)
                {
                    buffer[offset + (sample * _bytesPerSample) + b] = bytes[b];
                }
            }
            
            return samplesRead * _bytesPerSample;
        }
    }
}

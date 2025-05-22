using System;
using NAudio.Wave;

namespace AudioFusion
{
    public class SampleProvider : ISampleProvider
    {
        private readonly WaveStream _sourceStream;
        private readonly float[] _sampleBuffer;
        private readonly int _channels;

        public SampleProvider(WaveStream sourceStream)
        {
            _sourceStream = sourceStream ?? throw new ArgumentNullException(nameof(sourceStream));
            WaveFormat = WaveFormat.CreateIeeeFloatWaveFormat(sourceStream.WaveFormat.SampleRate, sourceStream.WaveFormat.Channels);
            _channels = sourceStream.WaveFormat.Channels;
            _sampleBuffer = new float[sourceStream.WaveFormat.SampleRate * sourceStream.WaveFormat.Channels];
        }

        public WaveFormat WaveFormat { get; }

        public int Read(float[] buffer, int offset, int count)
        {
            if (buffer == null) throw new ArgumentNullException(nameof(buffer));
            if (offset < 0) throw new ArgumentOutOfRangeException(nameof(offset));
            if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
            if (offset + count > buffer.Length) throw new ArgumentException("Buffer too small");

            int bytesNeeded = count * 4; // 4 bytes per float sample
            byte[] byteBuffer = new byte[bytesNeeded];
            int bytesRead = _sourceStream.Read(byteBuffer, 0, bytesNeeded);
            int samplesRead = bytesRead / 4;

            // Convert bytes to float samples
            for (int i = 0; i < samplesRead; i++)
            {
                float sampleValue = BitConverter.ToSingle(byteBuffer, i * 4);
                buffer[offset + i] = sampleValue;
            }

            return samplesRead;
        }
    }
}

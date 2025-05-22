using System;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;

namespace AudioFusion
{
    public static class WaveExtensionMethods
    {
        // Convert IWaveProvider to ISampleProvider
        public static ISampleProvider ToSampleProvider(this IWaveProvider waveProvider)
        {
            if (waveProvider is WaveStream waveStream)
            {
                return new SampleProvider(waveStream);
            }
            throw new InvalidOperationException("ToSampleProvider only supports WaveStream-based providers in this implementation.");
        }

        // Convert ISampleProvider to IWaveProvider
        public static IWaveProvider ToWaveProvider(this ISampleProvider sampleProvider)
        {
            return new SampleToWaveProvider(sampleProvider);
        }
    }
}

using CSCore;
using CSCore.MediaFoundation;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace AnnotatePPTX
{
    public class AACEncoder
    {
        private readonly ILogger _logger;

        public AACEncoder(ILogger logger)
        {
            this._logger = logger;
        }

        public byte[] FromWav(byte[] audioFile)
        {
            var supportedFormats = MediaFoundationEncoder.GetEncoderMediaTypes(AudioSubTypes.MPEG_HEAAC);

            _logger.Verbose("Checking for support of AAC encoding.");

            if (!supportedFormats.Any())
            {
                _logger.Verbose("The current platform does not support AAC encoding.");
                throw new ApplicationException("Current platform does not support AAC encoding.");
            }

            MemoryStream inStream = null;
            MemoryStream outStream = null;
            IWaveSource source = null;
            bool sourceDisposed = false;
            try
            {
                _logger.Verbose("Creating input stream and decoder.");

                inStream = new MemoryStream(audioFile);
                source = new CSCore.MediaFoundation.MediaFoundationDecoder(inStream);

                //in case the encoder does not support the input sample rate -> convert it to any supported samplerate
                //choose the best sample rate
                _logger.Verbose("Searching for the optimal sample rate.");
                _logger.Verbose($"Input wave format: {source.WaveFormat.ToString()}");
                int sampleRate =
                    supportedFormats.OrderBy(x => Math.Abs(source.WaveFormat.SampleRate - x.SampleRate))
                        .First(x => x.Channels == source.WaveFormat.Channels)
                        .SampleRate;
                if (source.WaveFormat.SampleRate != sampleRate)
                {
                    _logger.Verbose($"Changing sample rate of the source: {source.WaveFormat.SampleRate} -> {sampleRate}.");
                    source = source.ChangeSampleRate(sampleRate);
                }

                _logger.Verbose("Encoding WAV to AAC");
                outStream = new MemoryStream();
                using (source)
                {
                    using (var encoder = MediaFoundationEncoder.CreateAACEncoder(source.WaveFormat, outStream))
                    {
                        byte[] buffer = new byte[source.WaveFormat.BytesPerSecond];
                        int read;
                        while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            encoder.Write(buffer, 0, read);
                        }
                    }
                    sourceDisposed = true;
                }

                _logger.Verbose("Encoding is complete");

                return outStream.ToArray();
            }
            finally
            {
                _logger.Verbose("Cleaning up resources");

                if (inStream != null)
                    inStream.Dispose();

                if (source != null && !sourceDisposed)
                    source.Dispose();

                if (outStream != null)
                    inStream.Dispose();
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using Serilog;
using Microsoft.Extensions.DependencyInjection;
using System.IO;
using Microsoft.Extensions.Configuration;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace AnnotatePPTX
{
    class Program
    {
        private static ILogger _logger;
        private static IConfiguration _config;
        private static AppConfig _appConfiguration;
        // V8_(Recording)LMF Global Sales Training_deDE_enus_VO.pptx
        private static readonly Regex _pptxFileNameLocaleRegex = new Regex(".+_([a-zA-Z]+)_([a-zA-Z]+)_.*");

        static void Main(string[] args)
        {
            var services = Startup();
            try
            {
                Run(args);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "There was an exception while running the app.");
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        private static void Run(string[] args)
        {
            if (string.IsNullOrWhiteSpace(_appConfiguration.TranslatedAudioSourcePath)
                || string.IsNullOrWhiteSpace(_appConfiguration.PPTXSourcePath)
                || string.IsNullOrWhiteSpace(_appConfiguration.OutputPath))
            {
                _logger.Error("One or more required parameters are missing from the appsettings.json configuration file.");
                return;
            }

            if (!Directory.Exists(_appConfiguration.TranslatedAudioSourcePath))
                throw new ApplicationException($"Folder '{_appConfiguration.TranslatedAudioSourcePath}' does not exist. Please, check the 'translatedAudioSourcePath' configuration parameter.");

            if (!Directory.Exists(_appConfiguration.PPTXSourcePath))
                throw new ApplicationException($"Folder '{_appConfiguration.PPTXSourcePath}' does not exist. Please, check the 'pptxSourcePath' configuration parameter.");

            // check if output directory exists and create if needed
            if (!Directory.Exists(_appConfiguration.OutputPath))
            {
                _logger.Debug($"Creating output directory '{_appConfiguration.OutputPath}'.");

                try
                {
                    Directory.CreateDirectory(_appConfiguration.OutputPath);
                }
                catch(UnauthorizedAccessException authEx)
                {
                    _logger.Error(authEx, $"Not enough permissions to create output path '{_appConfiguration.OutputPath}'");
                    return;
                }
            }


            var stopwatch = new Stopwatch();
            var allPresentations = ReadListOfPresentations(_appConfiguration.PPTXSourcePath);
            _logger.Verbose($"Found {allPresentations.Count()} presentations in the '{_appConfiguration.PPTXSourcePath}' directory.");
            foreach (var pptxPath in allPresentations)
            {
                stopwatch.Reset();
                _logger.Information($"*****************************************");
                _logger.Information($"Processing file: '{pptxPath}'.");
                _logger.Information($"*****************************************");

                stopwatch.Start();
                var _pptxFileName = Path.GetFileName(pptxPath);

                // get locale from file name
                var localeCode = _pptxFileNameLocaleRegex.Match(_pptxFileName)?.Groups?.Skip(1)?.FirstOrDefault()?.Value;
                _logger.Debug($"Locale code from file name: '{localeCode}'");
                if (string.IsNullOrWhiteSpace(localeCode))
                {
                    _logger.Error("Could not parse file name for locale code. Please, check that it follows naming conventions: <file_name>_<target_locale>_<source_locale>_<file_name_suffix>.pptx");
                    continue;
                }

                // copy pptx file to the output directory. We will be working on this copy, not the original file.
                var fileCopyPath = System.IO.Path.Combine(_appConfiguration.OutputPath, _pptxFileName);

                _logger.Debug("Copying pptx to the output directory.");
                File.Copy(pptxPath, fileCopyPath, true);

                // build path to localized wav files based on locale
                var localizedAudioSourcePath = Path.Combine(_appConfiguration.TranslatedAudioSourcePath, localeCode);
                _logger.Debug($"Source path with localized audio files: '{localizedAudioSourcePath}'");
                if (!Directory.Exists(localizedAudioSourcePath))
                {
                    _logger.Warning($"Could not find directory '{localizedAudioSourcePath}' with localized audio files based on locale from pptx file name. The file will have original annotations.");
                    continue;
                }

                var doc = new Presentation(fileCopyPath, Log.Logger, new AACEncoder(Log.Logger));

                var comments = doc.GetAllSlideNotes();
                var translatedAnnotations = Directory.EnumerateFiles(localizedAudioSourcePath, "*.wav", SearchOption.AllDirectories);
                _logger.Verbose($"Found '{translatedAnnotations.Count()}' translated annotations for the presentation.");

                doc.ReplaceAllSlideAudioAnnotations(
                    comments.ToDictionary(
                        c => c.SlideRelationshipId, 
                        c => translatedAnnotations.FirstOrDefault(path => path.Contains($"_Slide {c.SlideShowIndex}_"))
                    )
                );

                _logger.Information($"Finished processing file: '{pptxPath}'");
                stopwatch.Stop();
                _logger.Verbose($"Elapsed time: {stopwatch.Elapsed.ToString(@"hh\:mm\:ss")}");
            }
        }

        private static IServiceProvider Startup()
        {
            _config = new ConfigurationBuilder()
              .AddJsonFile("appsettings.json", true, true)
              .Build();

            _appConfiguration = _config.GetSection("application").Get<AppConfig>();

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Verbose()
                .WriteTo.File("runtime.log", restrictedToMinimumLevel: Serilog.Events.LogEventLevel.Verbose)
                .WriteTo.Console(restrictedToMinimumLevel: Serilog.Events.LogEventLevel.Information)
                .CreateLogger();

            _logger = Log.Logger;

            var serviceCollection = new ServiceCollection();
            ConfigureServices(serviceCollection);

            var serviceProvider = serviceCollection.BuildServiceProvider();
            return serviceProvider;
        }

        private static void ConfigureServices(IServiceCollection services)
        {
            services.AddLogging(configure => configure.AddSerilog());
        }

        private static IEnumerable<string> ReadListOfPresentations(string sourcePath)
        {
            return System.IO.Directory.GetFiles(sourcePath, "*.pptx");
        }
    }
}

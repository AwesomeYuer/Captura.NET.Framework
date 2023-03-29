using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Captura.Audio;
using Captura.FFmpeg;
using Captura.Models;
using Captura.SharpAvi;
using Captura.Video;
using Captura.ViewModels;
using Captura.Webcam;
using Captura.Windows;
using static System.Console;

namespace Captura
{
    // ReSharper disable once ClassNeverInstantiated.Global
    class ConsoleManager : IDisposable
    {
        readonly Settings _settings;
        readonly RecordingModel _recordingModel;
        readonly ScreenShotModel _screenShotModel;
        readonly IEnumerable<IVideoSourceProvider> _videoSourceProviders;
        readonly IEnumerable<IVideoWriterProvider> _videoWriterProviders;
        readonly WebcamModel _webcamModel;
        readonly IPlatformServices _platformServices;
        readonly IMessageProvider _messageProvider;
        readonly IAudioSource _audioSource;

        public ConsoleManager(Settings Settings,
            RecordingModel RecordingModel,
            ScreenShotModel ScreenShotModel,
            IEnumerable<IVideoSourceProvider> VideoSourceProviders,
            IEnumerable<IVideoWriterProvider> VideoWriterProviders,
            IPlatformServices PlatformServices,
            WebcamModel WebcamModel,
            IMessageProvider MessageProvider,
            IAudioSource AudioSource)
        {
            _settings = Settings;
            _recordingModel = RecordingModel;
            _screenShotModel = ScreenShotModel;
            _videoSourceProviders = VideoSourceProviders;
            _videoWriterProviders = VideoWriterProviders;
            _platformServices = PlatformServices;
            _webcamModel = WebcamModel;
            _messageProvider = MessageProvider;
            _audioSource = AudioSource;

            // Hide on Full Screen Screenshot doesn't work on Console
            Settings.UI.HideOnFullScreenShot = false;
        }

        public void Dispose()
        {
            ServiceProvider.Dispose();
        }

        public void CopySettings()
        {
            // Load settings dummy
            var dummySettings = new Settings(new FFmpegSettings(), new WindowsSettings());
            dummySettings.Load();

            _settings.WebcamOverlay = dummySettings.WebcamOverlay;
            _settings.MousePointerOverlay = dummySettings.MousePointerOverlay;
            _settings.Clicks = dummySettings.Clicks;
            _settings.Keystrokes = dummySettings.Keystrokes;
            _settings.Elapsed = dummySettings.Elapsed;

            // add by Awesome Yuer
            _settings.Audio.RecordMicrophone = dummySettings.Audio.RecordMicrophone;
            _settings.Audio.RecordSpeaker = dummySettings.Audio.RecordSpeaker;


            

            _settings.OutPath = dummySettings.OutPath;

            // FFmpeg Path
            _settings.FFmpeg.FolderPath = dummySettings.FFmpeg.FolderPath;

            // add by Awesome Yuer
            // _settings = dummySettings;

            foreach (var overlay in dummySettings.Censored)
            {
                _settings.Censored.Add(overlay);
            }

            foreach (var overlay in dummySettings.TextOverlays)
            {
                _settings.TextOverlays.Add(overlay);
            }

            foreach (var overlay in dummySettings.ImageOverlays)
            {
                _settings.ImageOverlays.Add(overlay);
            }
        }

        public void Start(StartCmdOptions startOptions)
        {
            _settings.IncludeCursor = startOptions.Cursor;
            _settings.Clicks.Display = startOptions.Clicks;
            _settings.Keystrokes.Display = startOptions.Keys;

            // Output Folder
            // add by Awsesome Yuer
            if
                (
                    !string.IsNullOrEmpty(startOptions.OutPath)
                    &&
                    !string.IsNullOrWhiteSpace(startOptions.OutPath)
                )
            {
                _settings.OutPath = startOptions.OutPath;
            }
            else
            {
                startOptions.OutPath = _settings.OutPath;
            }
            
            if (File.Exists(startOptions.FileName))
            {
                if (!startOptions.Overwrite)
                {
                    if (!_messageProvider
                        .ShowYesNo("Output File Already Exists, Do you want to overwrite?", ""))
                        return;
                }

                File.Delete(startOptions.FileName);
            }

            var videoSourceKind = HandleVideoSource(startOptions);

            if (videoSourceKind == null)
            {
                WriteLine("Video source not set or invalid");

                return;
            }

            HandleAudioSource(startOptions, out var mic, out var speaker);

            HandleWebcam(startOptions);

            if (startOptions.FrameRate is int frameRate)
                _settings.Video.FrameRate = frameRate;

            if (startOptions.AudioQuality is int aq)
                _settings.Audio.Quality = aq;

            if (startOptions.VideoQuality is int vq)
                _settings.Video.Quality = vq;

            if (startOptions.Replay is int replayDuration && replayDuration > 0)
            {
                _settings.Video.RecorderMode = RecorderMode.Replay;
                _settings.Video.ReplayDuration = replayDuration;
            }

            var videoWriter = HandleVideoEncoder(startOptions, out _);

            if (startOptions.Delay > 0)
                Thread.Sleep(startOptions.Delay);

            if (!_recordingModel.StartRecording(new RecordingModelParams
            {
                VideoSourceKind = videoSourceKind,
                VideoWriter = videoWriter,
                Microphone = mic,
                Speaker = speaker
            }, startOptions.FileName))
                return;

            Task.Factory.StartNew(() =>
            {
                Loop(startOptions);

                _recordingModel.StopRecording().Wait();

                Application.Exit();
            });

            // MouseKeyHook requires a Window Handle to register
            Application.Run(new ApplicationContext());
        }

        public void Shot(ShotCmdOptions ShotOptions)
        {
            _settings.IncludeCursor = ShotOptions.Cursor;

            // Screenshot Window with Transparency
            if (Regex.IsMatch(ShotOptions.Source, @"win:\d+"))
            {
                var ptr = int.Parse(ShotOptions.Source.Substring(4));

                try
                {
                    var win = _platformServices.GetWindow(new IntPtr(ptr));
                    var bmp = _screenShotModel.ScreenShotWindow(win);

                    _screenShotModel.SaveScreenShot(bmp, ShotOptions.FileName).Wait();
                }
                catch
                {
                    // Suppress Errors
                }
            }
            else
            {
                var videoSourceKind = HandleVideoSource(ShotOptions);

                var bmp = _screenShotModel.GetScreenShot(videoSourceKind).Result;

                _screenShotModel.SaveScreenShot(bmp, ShotOptions.FileName).Wait();
            }
        }

        IVideoSourceProvider HandleVideoSource(CommonCmdOptions CommonOptions)
        {
            var provider = _videoSourceProviders.FirstOrDefault(M => M.ParseCli(CommonOptions.Source));

            return provider;
        }

        void HandleAudioSource
                            (
                                  StartCmdOptions startOptions
                                , out IAudioItem microphone
                                , out IAudioItem speaker
                            )
        {
            microphone = speaker = null;

            var microphones = _audioSource
                                    .Microphones
                                    .ToArray();

            var speakers = _audioSource
                                    .Speakers
                                    .ToArray();

            

            if (microphones.Length > 0)
            {
                if
                    (
                        startOptions.Microphone > -1
                        &&
                        startOptions.Microphone < microphones.Length
                    )
                {
                    _settings.Audio.RecordMicrophone = true;
                    microphone = microphones[startOptions.Microphone];
                }
                else if
                    (
                        _settings.Audio.RecordMicrophone
                    )
                {
                    var microphoneId = _settings.Audio.Microphone;
                    if
                        (
                            string.IsNullOrEmpty(_settings.Audio.Microphone)
                            ||
                            string.IsNullOrWhiteSpace(_settings.Audio.Microphone)
                        )
                    {
                        microphoneId = "0";
                    }
                    if
                        (
                            int
                                .TryParse
                                        (
                                            microphoneId
                                            , out var settingsMicrophone
                                        )
                            &&
                            settingsMicrophone < microphones.Length
                            &&
                            settingsMicrophone > -1
                        )
                    {
                        microphone = microphones[settingsMicrophone];
                        startOptions.Microphone = settingsMicrophone;
                    }
                }
            }


            if (speakers.Length > 0)
            {
                if
                    (
                        startOptions.Speaker > -1
                        &&
                        startOptions.Speaker < speakers.Length
                    )
                {
                    _settings.Audio.RecordSpeaker = true;
                    speaker = speakers[startOptions.Speaker];
                }
                else if
                    (
                        _settings.Audio.RecordSpeaker
                    )
                {
                    var speakerId = _settings.Audio.Speaker;
                    if
                        (
                            string.IsNullOrEmpty(_settings.Audio.Speaker)
                            ||
                            string.IsNullOrWhiteSpace(_settings.Audio.Speaker)
                        )
                    {
                        speakerId = "0";
                    }
                    if
                        (
                            int
                                .TryParse
                                        (
                                            speakerId
                                            , out var settingsSpeaker
                                        )
                            &&
                            settingsSpeaker < speakers.Length
                            &&
                            settingsSpeaker > -1
                        )
                    {
                        speaker = speakers[settingsSpeaker];
                        startOptions.Speaker = settingsSpeaker;
                    }
                }
            }
        }

        IVideoWriterItem HandleVideoEncoder(StartCmdOptions StartOptions, out IVideoWriterProvider VideoWriterKind)
        {
            var selected = _videoWriterProviders
                .Select(M => new
                {
                    kind = M,
                    writer = M.ParseCli(StartOptions.Encoder)
                })
                .FirstOrDefault(M => M.writer != null);

            if (selected != null)
            {
                VideoWriterKind = selected.kind;

                return selected.writer;
            }

            var sharpAviWriterProvider = ServiceProvider.Get<SharpAviWriterProvider>();

            // Steps in video
            if (StartOptions.Encoder == "steps:video")
            {
                _settings.Video.RecorderMode = RecorderMode.Steps;

                VideoWriterKind = null;
                return new StepsVideoWriterItem(sharpAviWriterProvider.First());
            }

            // Steps in set of images
            if (StartOptions.Encoder == "steps:images")
            {
                _settings.Video.RecorderMode = RecorderMode.Steps;

                VideoWriterKind = null;
                return new ImageFolderWriterItem();
            }

            VideoWriterKind = sharpAviWriterProvider;
            return sharpAviWriterProvider.First();
        }

        void HandleWebcam(StartCmdOptions StartOptions)
        {
            if (StartOptions.Webcam != -1 && StartOptions.Webcam < _webcamModel.AvailableCams.Count - 1)
            {
                _webcamModel.SelectedCam = _webcamModel.AvailableCams[StartOptions.Webcam + 1];

                // HACK: Sleep to prevent AccessViolationException
                Thread.Sleep(500);
            }
        }

        void Loop(StartCmdOptions StartOptions)
        {
            if (StartOptions.Length > 0)
            {
                var elapsed = 0;

                Write(TimeSpan.Zero);

                while (elapsed++ < StartOptions.Length)
                {
                    Thread.Sleep(1000);
                    Write(new string('\b', 8) + TimeSpan.FromSeconds(elapsed));
                }

                Write(new string('\b', 8));
            }
            else
            {
                const string recordingText = "Press p to pause or resume, q to quit";

                WriteLine(recordingText);

                char ReadChar()
                {
                    if (IsInputRedirected)
                    {
                        var line = ReadLine();

                        if (line != null && line.Length == 1)
                            return line[0];

                        return char.MinValue;
                    }

                    return char.ToLower(ReadKey(true).KeyChar);
                }

                char c;

                do
                {
                    c = ReadChar();

                    if (c != 'p')
                        continue;

                    _recordingModel.OnPauseExecute();

                    if (_recordingModel.RecorderState != RecorderState.Paused)
                    {
                        WriteLine("Resumed");
                    }
                } while (c != 'q');
            }
        }
    }
}
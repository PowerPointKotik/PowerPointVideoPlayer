
using System;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Globalization;
using System.Threading;
using System.Collections.Generic;
using System.Windows.Forms;

using OpenTK;
using OpenTK.Graphics;
using OpenTK.Graphics.OpenGL;
using OpenTK.Input;

using NAudio.Wave;

namespace PowerPointVideoPlayer
{
    public class OpenGLWindow : GameWindow
    {
        private readonly IntPtr parentHandle;

        private const int VIDEO_WIDTH = 1280;
        private const int VIDEO_HEIGHT = 720;
        private const int VIDEO_FPS = 30;
        private const int FRAME_SIZE = VIDEO_WIDTH * VIDEO_HEIGHT * 3;

        private int videoTexture;
        private int fpsTexture;
        private readonly Dictionary<string, int> textTextures =
            new Dictionary<string, int>();

        private readonly List<string> playlist = new List<string>();
        private int currentTrackIndex = -1;
        private bool isPaused = false;
        private bool trackEndHandled = false;
        private bool controlsVisible = false;
        private bool isDraggingSeekBar = false;
        private double seekPreviewSeconds = 0;
        private double mediaDurationSeconds = 0;
        private double playbackStartOffset = 0;

        private Process videoProcess;
        private Process audioProcess;

        private Thread videoThread;
        private Thread audioThread;

        private volatile bool isReading = false;

        private readonly Queue<VideoFrame> frames =
            new Queue<VideoFrame>();

        private readonly object frameLock = new object();

        private VideoFrame currentFrame;

        private BufferedWaveProvider audioBuffer;
        private TrackingWaveProvider audioClock;
        private WaveOutEvent audioOutput;

        private readonly ManualResetEventSlim audioReady =
            new ManualResetEventSlim(false);

        private int frameCounter = 0;
        private int currentFPS = 0;
        private double fpsTimer = 0;

        private const int GWL_STYLE = -16;
        private const int WS_CHILD = 0x40000000;
        private const int WS_VISIBLE = 0x10000000;

        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_FRAMECHANGED = 0x0020;
        private const uint SWP_SHOWWINDOW = 0x0040;

        [DllImport("user32.dll")]
        private static extern IntPtr SetParent(
            IntPtr hWndChild,
            IntPtr hWndNewParent);

        [DllImport("user32.dll")]
        private static extern int SetWindowLong(
            IntPtr hWnd,
            int nIndex,
            int dwNewLong);

        [DllImport("user32.dll")]
        private static extern bool GetClientRect(
            IntPtr hWnd,
            out RECT rect);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(
            IntPtr hWnd,
            IntPtr hWndInsertAfter,
            int X,
            int Y,
            int cx,
            int cy,
            uint uFlags);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private class VideoFrame
        {
            public byte[] Data;
            public double Timestamp;
        }

        private class TrackingWaveProvider : IWaveProvider
        {
            private readonly IWaveProvider source;
            private long bytesRead;

            public TrackingWaveProvider(IWaveProvider source)
            {
                this.source = source;
            }

            public WaveFormat WaveFormat
            {
                get { return source.WaveFormat; }
            }

            public double SecondsPlayed
            {
                get
                {
                    return Interlocked.Read(ref bytesRead) /
                        (double)WaveFormat.AverageBytesPerSecond;
                }
            }

            public int Read(byte[] buffer, int offset, int count)
            {
                int read = source.Read(buffer, offset, count);

                Interlocked.Add(ref bytesRead, read);

                return read;
            }
        }

        public OpenGLWindow(IntPtr parent)
            : base(800, 600, GraphicsMode.Default,
                  "PowerPoint Video Player")
        {
            parentHandle = parent;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            IntPtr hwnd = WindowInfo.Handle;

            if (parentHandle != IntPtr.Zero &&
                GetClientRect(parentHandle, out RECT rect))
            {
                int width = rect.Right - rect.Left;
                int height = rect.Bottom - rect.Top;

                SetWindowLong(
                    hwnd,
                    GWL_STYLE,
                    WS_CHILD | WS_VISIBLE);

                SetParent(hwnd, parentHandle);

                SetWindowPos(
                    hwnd,
                    IntPtr.Zero,
                    0,
                    0,
                    width,
                    height,
                    SWP_NOZORDER |
                    SWP_FRAMECHANGED |
                    SWP_SHOWWINDOW);
            }

            GL.ClearColor(0f, 0f, 0f, 1f);
            GL.Disable(EnableCap.DepthTest);

            // Ограничиваем перерисовку частотой обновления монитора.
            // Это уменьшает лишнюю нагрузку на CPU/GPU.
            VSync = VSyncMode.On;

            // Текстура видео
            videoTexture = GL.GenTexture();

            GL.BindTexture(
                TextureTarget.Texture2D,
                videoTexture);

            GL.TexParameter(
                TextureTarget.Texture2D,
                TextureParameterName.TextureMinFilter,
                (int)TextureMinFilter.Linear);

            GL.TexParameter(
                TextureTarget.Texture2D,
                TextureParameterName.TextureMagFilter,
                (int)TextureMagFilter.Linear);

            GL.TexImage2D(
                TextureTarget.Texture2D,
                0,
                PixelInternalFormat.Rgb,
                VIDEO_WIDTH,
                VIDEO_HEIGHT,
                0,
                OpenTK.Graphics.OpenGL.PixelFormat.Rgb,
                PixelType.UnsignedByte,
                IntPtr.Zero);

            // Текстура FPS
            fpsTexture = GL.GenTexture();

            GL.BindTexture(
                TextureTarget.Texture2D,
                fpsTexture);

            GL.TexParameter(
                TextureTarget.Texture2D,
                TextureParameterName.TextureMinFilter,
                (int)TextureMinFilter.Linear);

            GL.TexParameter(
                TextureTarget.Texture2D,
                TextureParameterName.TextureMagFilter,
                (int)TextureMagFilter.Linear);

            UpdateFPSTexture();

            // Выбор видео
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "Выбери видео";

                dialog.Filter =
                    "Video files|*.mp4;*.mkv;*.avi;*.mov;*.webm|All files|*.*";

                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    Exit();
                    return;
                }

                playlist.Add(dialog.FileName);
                currentTrackIndex = 0;
                StartPlayback(dialog.FileName);
            }
        }

        private void StartPlayback(string file, double startAtSeconds = 0)
        {
            try
            {
                isReading = true;
                isPaused = false;
                trackEndHandled = false;
                playbackStartOffset = Math.Max(0, startAtSeconds);
                mediaDurationSeconds = GetMediaDuration(file);
                seekPreviewSeconds = playbackStartOffset;
                audioReady.Reset();
                currentFrame = null;
                lock (frameLock)
                {
                    frames.Clear();
                    Monitor.PulseAll(frameLock);
                }

                // Настройка аудио
                audioBuffer = new BufferedWaveProvider(
                    new WaveFormat(44100, 16, 2));

                audioBuffer.BufferDuration =
                    TimeSpan.FromSeconds(5);

                audioBuffer.DiscardOnBufferOverflow = true;

                audioClock = new TrackingWaveProvider(audioBuffer);

                audioOutput = new WaveOutEvent();

                audioOutput.DesiredLatency = 100;
                audioOutput.NumberOfBuffers = 3;

                audioOutput.Init(audioClock);

                // FFmpeg: видео
                ProcessStartInfo videoInfo =
                    new ProcessStartInfo();

                videoInfo.FileName = "ffmpeg";

                videoInfo.Arguments =
                    "-nostdin -loglevel error " +
                    "-ss " + playbackStartOffset.ToString(CultureInfo.InvariantCulture) + " " +
                    "-i \"" + file + "\" " +
                    "-vf \"scale=1280:720:force_original_aspect_ratio=decrease," +
                    "pad=1280:720:(ow-iw)/2:(oh-ih)/2,fps=30\" " +
                    "-an -sn " +
                    "-f rawvideo -pix_fmt rgb24 pipe:1";

                videoInfo.UseShellExecute = false;
                videoInfo.CreateNoWindow = true;
                videoInfo.RedirectStandardOutput = true;
                videoInfo.RedirectStandardError = true;

                videoProcess = new Process();
                videoProcess.StartInfo = videoInfo;
                videoProcess.Start();

                videoProcess.BeginErrorReadLine();

                // FFmpeg: аудио
                ProcessStartInfo audioInfo =
                    new ProcessStartInfo();

                audioInfo.FileName = "ffmpeg";

                audioInfo.Arguments =
                    "-nostdin -loglevel error " +
                    "-ss " + playbackStartOffset.ToString(CultureInfo.InvariantCulture) + " " +
                    "-i \"" + file + "\" " +
                    "-vn -sn " +
                    "-ac 2 -ar 44100 " +
                    "-acodec pcm_s16le " +
                    "-f s16le pipe:1";

                audioInfo.UseShellExecute = false;
                audioInfo.CreateNoWindow = true;
                audioInfo.RedirectStandardOutput = true;
                audioInfo.RedirectStandardError = true;

                audioProcess = new Process();
                audioProcess.StartInfo = audioInfo;
                audioProcess.Start();

                audioProcess.BeginErrorReadLine();

                // Запуск потоков декодирования
                videoThread = new Thread(ReadVideoFrames);
                videoThread.IsBackground = true;
                videoThread.Start();

                audioThread = new Thread(ReadAudio);
                audioThread.IsBackground = true;
                audioThread.Start();

                // Ждём, пока появится аудиобуфер
                audioReady.Wait(5000);

                // Запускаем аудио
                audioOutput.Play();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка запуска воспроизведения:\n" +
                    ex.Message);
            }
        }

        private void ReadVideoFrames()
        {
            try
            {
                Stream stream =
                    videoProcess.StandardOutput.BaseStream;

                int frameNumber = 0;

                while (isReading)
                {
                    byte[] frame = new byte[FRAME_SIZE];

                    int totalRead = 0;

                    while (totalRead < FRAME_SIZE && isReading)
                    {
                        int read = stream.Read(
                            frame,
                            totalRead,
                            FRAME_SIZE - totalRead);

                        if (read <= 0)
                            break;

                        totalRead += read;
                    }

                    if (totalRead != FRAME_SIZE)
                        break;

                    VideoFrame videoFrame = new VideoFrame();

                    videoFrame.Data = frame;
                    videoFrame.Timestamp =
                        playbackStartOffset + frameNumber / (double)VIDEO_FPS;

                    frameNumber++;

                    lock (frameLock)
                    {
                        while (frames.Count >= 12 && isReading)
                        {
                            Monitor.Wait(frameLock, 50);
                        }

                        if (!isReading)
                            break;

                        frames.Enqueue(videoFrame);

                        Monitor.PulseAll(frameLock);
                    }
                }
            }
            catch
            {
            }

            lock (frameLock)
            {
                Monitor.PulseAll(frameLock);
            }
        }

        private void ReadAudio()
        {
            try
            {
                Stream stream =
                    audioProcess.StandardOutput.BaseStream;

                byte[] buffer = new byte[8192];

                while (isReading)
                {
                    // Не даём аудиобуферу переполняться
                    if (audioBuffer.BufferedBytes >
                        audioBuffer.WaveFormat.AverageBytesPerSecond * 2)
                    {
                        Thread.Sleep(10);
                        continue;
                    }

                    int read = stream.Read(
                        buffer,
                        0,
                        buffer.Length);

                    if (read <= 0)
                        break;

                    audioBuffer.AddSamples(
                        buffer,
                        0,
                        read);

                    if (audioBuffer.BufferedBytes >=
                        audioBuffer.WaveFormat.AverageBytesPerSecond / 2)
                    {
                        audioReady.Set();
                    }
                }
            }
            catch
            {
            }

            audioReady.Set();
        }

        private void UpdateVideoTexture(byte[] frame)
        {
            GL.BindTexture(
                TextureTarget.Texture2D,
                videoTexture);

            GL.PixelStore(
                PixelStoreParameter.UnpackAlignment,
                1);

            GCHandle handle = GCHandle.Alloc(
                frame,
                GCHandleType.Pinned);

            try
            {
                GL.TexSubImage2D(
                    TextureTarget.Texture2D,
                    0,
                    0,
                    0,
                    VIDEO_WIDTH,
                    VIDEO_HEIGHT,
                    OpenTK.Graphics.OpenGL.PixelFormat.Rgb,
                    PixelType.UnsignedByte,
                    handle.AddrOfPinnedObject());
            }
            finally
            {
                handle.Free();
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);

            if (Width <= 0 || Height <= 0)
                return;

            GL.Viewport(0, 0, Width, Height);
        }

        protected override void OnRenderFrame(FrameEventArgs e)
        {
            base.OnRenderFrame(e);

            // FPS
            frameCounter++;
            fpsTimer += e.Time;

            if (fpsTimer >= 1.0)
            {
                currentFPS = frameCounter;

                frameCounter = 0;
                fpsTimer = 0;

                UpdateFPSTexture();
            }

            // Синхронизация видео по аудиочасам.
            // Обновляем OpenGL-текстуру только при получении нового кадра.
            // Раньше один и тот же кадр загружался в GPU на каждом render-цикле,
            // даже если видео-кадр не изменился.
            bool newFrameAvailable = false;

            if (audioClock != null)
            {
                double audioTime = playbackStartOffset + audioClock.SecondsPlayed;

                lock (frameLock)
                {
                    while (frames.Count > 0)
                    {
                        VideoFrame next = frames.Peek();

                        if (next.Timestamp > audioTime)
                            break;

                        currentFrame = frames.Dequeue();
                        newFrameAvailable = true;

                        Monitor.PulseAll(frameLock);
                    }
                }
            }

            if (newFrameAvailable && currentFrame != null)
            {
                UpdateVideoTexture(currentFrame.Data);
            }

            GL.Clear(ClearBufferMask.ColorBufferBit);

            DrawVideo();
            DrawFPS();
            if (controlsVisible)
                DrawControls();

            // Автоматически переходим к следующему файлу после завершения обоих декодеров.
            bool noQueuedFrames;
            lock (frameLock) { noQueuedFrames = frames.Count == 0; }
            bool audioDrained = audioBuffer == null || audioBuffer.BufferedBytes == 0;
            if (!trackEndHandled && isReading &&
                videoProcess != null && audioProcess != null &&
                SafeHasExited(videoProcess) && SafeHasExited(audioProcess) &&
                noQueuedFrames && audioDrained)
            {
                trackEndHandled = true;
                if (currentTrackIndex + 1 < playlist.Count)
                    PlayTrack(currentTrackIndex + 1);
                else
                    StopPlayback();
            }

            SwapBuffers();
        }

        private bool SafeHasExited(Process process)
        {
            try { return process.HasExited; }
            catch { return true; }
        }

        private void DrawControls()
        {
            // Нижняя панель управления.
            GL.MatrixMode(MatrixMode.Projection);
            GL.PushMatrix();
            GL.LoadIdentity();
            GL.Ortho(0, Width, Height, 0, -1, 1);
            GL.MatrixMode(MatrixMode.Modelview);
            GL.PushMatrix();
            GL.LoadIdentity();

            GL.Disable(EnableCap.Texture2D);
            GL.Color4(0.07f, 0.08f, 0.11f, 0.94f);
            GL.Begin(PrimitiveType.Quads);
            GL.Vertex2(0, Height - 100);
            GL.Vertex2(Width, Height - 100);
            GL.Vertex2(Width, Height);
            GL.Vertex2(0, Height);
            GL.End();

            DrawSeekBar();

            DrawButton(12, Height - 55, 76, 42, "OPEN");
            DrawButton(98, Height - 55, 76, 42, "PREV");
            DrawButton(184, Height - 55, 104, 42, isPaused ? "PLAY" : "PAUSE");
            DrawButton(298, Height - 55, 76, 42, "NEXT");

            GL.Enable(EnableCap.Texture2D);
            GL.Color4(1f, 1f, 1f, 1f);
            DrawText("PLAYLIST", 395, Height - 47, 82, 20);

            int visible = Math.Min(playlist.Count, 4);
            int start = Math.Max(0, Math.Min(currentTrackIndex - 1, playlist.Count - visible));
            for (int i = 0; i < visible; i++)
            {
                int index = start + i;
                string name = Path.GetFileName(playlist[index]);
                if (name.Length > 24) name = name.Substring(0, 21) + "...";
                if (index == currentTrackIndex)
                {
                    GL.Disable(EnableCap.Texture2D);
                    GL.Color4(0.18f, 0.35f, 0.25f, 1f);
                    GL.Begin(PrimitiveType.Quads);
                    GL.Vertex2(490, Height - 59 + i * 14);
                    GL.Vertex2(Math.Min(Width - 8, 760), Height - 59 + i * 14);
                    GL.Vertex2(Math.Min(Width - 8, 760), Height - 46 + i * 14);
                    GL.Vertex2(490, Height - 46 + i * 14);
                    GL.End();
                    GL.Enable(EnableCap.Texture2D);
                }
                DrawText(name, 495, Height - 57 + i * 14, Math.Min(250, Math.Max(55, name.Length * 8)), 12);
            }

            GL.Disable(EnableCap.Texture2D);
            GL.PopMatrix();
            GL.MatrixMode(MatrixMode.Projection);
            GL.PopMatrix();
            GL.MatrixMode(MatrixMode.Modelview);
        }

        private double GetPlaybackPosition()
        {
            if (isDraggingSeekBar)
                return seekPreviewSeconds;

            if (audioClock != null)
                return Math.Min(mediaDurationSeconds > 0 ? mediaDurationSeconds : double.MaxValue,
                    playbackStartOffset + audioClock.SecondsPlayed);

            return playbackStartOffset;
        }

        private string FormatTime(double seconds)
        {
            if (double.IsNaN(seconds) || double.IsInfinity(seconds) || seconds < 0)
                seconds = 0;

            TimeSpan time = TimeSpan.FromSeconds(seconds);
            int totalMinutes = (int)Math.Floor(time.TotalMinutes);
            return totalMinutes.ToString("00") + ":" + time.Seconds.ToString("00");
        }

        private void DrawSeekBar()
        {
            int left = 16;
            int right = Math.Max(left + 1, Width - 16);
            int barY = Height - 88;
            int barHeight = 6;

            double position = GetPlaybackPosition();
            double progress = mediaDurationSeconds > 0
                ? Math.Max(0, Math.Min(1, position / mediaDurationSeconds))
                : 0;

            GL.Disable(EnableCap.Texture2D);

            // Фоновая дорожка.
            GL.Color4(0.30f, 0.32f, 0.37f, 1f);
            GL.Begin(PrimitiveType.Quads);
            GL.Vertex2(left, barY);
            GL.Vertex2(right, barY);
            GL.Vertex2(right, barY + barHeight);
            GL.Vertex2(left, barY + barHeight);
            GL.End();

            // Заполненная часть.
            int filled = left + (int)((right - left) * progress);
            GL.Color4(0.20f, 0.65f, 0.95f, 1f);
            GL.Begin(PrimitiveType.Quads);
            GL.Vertex2(left, barY);
            GL.Vertex2(filled, barY);
            GL.Vertex2(filled, barY + barHeight);
            GL.Vertex2(left, barY + barHeight);
            GL.End();

            // Ползунок.
            GL.Color4(1f, 1f, 1f, 1f);
            GL.Begin(PrimitiveType.Quads);
            GL.Vertex2(filled - 4, barY - 3);
            GL.Vertex2(filled + 4, barY - 3);
            GL.Vertex2(filled + 4, barY + barHeight + 3);
            GL.Vertex2(filled - 4, barY + barHeight + 3);
            GL.End();

            GL.Enable(EnableCap.Texture2D);
            DrawText(FormatTime(position), left, Height - 78, 52, 14);
            DrawText(FormatTime(mediaDurationSeconds), Math.Max(left + 60, right - 52), Height - 78, 52, 14);
        }

        private void UpdateSeekPreview(int mouseX)
        {
            if (mediaDurationSeconds <= 0 || Width <= 32)
                return;

            double fraction = (mouseX - 16.0) / (Width - 32.0);
            fraction = Math.Max(0, Math.Min(1, fraction));
            seekPreviewSeconds = fraction * mediaDurationSeconds;
        }

        private void SeekTo(double seconds)
        {
            if (currentTrackIndex < 0 || currentTrackIndex >= playlist.Count || mediaDurationSeconds <= 0)
                return;

            seconds = Math.Max(0, Math.Min(mediaDurationSeconds, seconds));
            StopPlayback();
            StartPlayback(playlist[currentTrackIndex], seconds);
        }

        private double GetMediaDuration(string file)
        {
            try
            {
                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "ffprobe";
                info.Arguments = "-v error -show_entries format=duration " +
                    "-of default=noprint_wrappers=1:nokey=1 \"" + file + "\"";
                info.UseShellExecute = false;
                info.CreateNoWindow = true;
                info.RedirectStandardOutput = true;

                using (Process probe = Process.Start(info))
                {
                    string output = probe.StandardOutput.ReadToEnd();
                    if (!probe.WaitForExit(5000))
                    {
                        try { probe.Kill(); } catch { }
                        return 0;
                    }

                    double duration;
                    if (double.TryParse(output.Trim(), NumberStyles.Float,
                        CultureInfo.InvariantCulture, out duration) && duration > 0)
                        return duration;
                }
            }
            catch
            {
                // Если ffprobe недоступен, воспроизведение продолжится,
                // но длительность и перемотка будут недоступны.
            }

            return 0;
        }

        private void DrawButton(int x, int y, int width, int height, string label)
        {
            GL.Disable(EnableCap.Texture2D);
            GL.Color4(0.20f, 0.23f, 0.29f, 1f);
            GL.Begin(PrimitiveType.Quads);
            GL.Vertex2(x, y);
            GL.Vertex2(x + width, y);
            GL.Vertex2(x + width, y + height);
            GL.Vertex2(x, y + height);
            GL.End();
            GL.Enable(EnableCap.Texture2D);
            DrawText(label, x + 10, y + 12, label.Length * 9, 18);
        }

        private void DrawText(string value, int x, int y, int width, int height)
        {
            int texture;
            if (!textTextures.TryGetValue(value, out texture))
            {
                // Ограничиваем кеш, поскольку время на полоске меняется каждую секунду.
                if (textTextures.Count >= 128)
                {
                    foreach (int oldTexture in textTextures.Values)
                        GL.DeleteTexture(oldTexture);
                    textTextures.Clear();
                }

                using (Bitmap bitmap = new Bitmap(Math.Max(64, value.Length * 12), 28))
                using (Graphics graphics = Graphics.FromImage(bitmap))
                using (Font font = new Font("Arial", 12, FontStyle.Bold))
                {
                    graphics.Clear(Color.Transparent);
                    graphics.DrawString(value, font, Brushes.White, 1, 2);
                    BitmapData data = bitmap.LockBits(
                        new Rectangle(0, 0, bitmap.Width, bitmap.Height),
                        ImageLockMode.ReadOnly,
                        System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    try
                    {
                        texture = GL.GenTexture();
                        GL.BindTexture(TextureTarget.Texture2D, texture);
                        GL.TexParameter(TextureTarget.Texture2D, TextureParameterName.TextureMinFilter, (int)TextureMinFilter.Linear);
                        GL.TexParameter(TextureTarget.Texture2D, TextureParameterName.TextureMagFilter, (int)TextureMagFilter.Linear);
                        GL.TexImage2D(TextureTarget.Texture2D, 0, PixelInternalFormat.Rgba,
                            bitmap.Width, bitmap.Height, 0,
                            OpenTK.Graphics.OpenGL.PixelFormat.Bgra,
                            PixelType.UnsignedByte, data.Scan0);
                        textTextures[value] = texture;
                    }
                    finally { bitmap.UnlockBits(data); }
                }
            }

            GL.Enable(EnableCap.Blend);
            GL.BlendFunc(BlendingFactor.SrcAlpha, BlendingFactor.OneMinusSrcAlpha);
            GL.BindTexture(TextureTarget.Texture2D, texture);
            GL.Color4(1f, 1f, 1f, 1f);
            GL.Begin(PrimitiveType.Quads);
            GL.TexCoord2(0, 0); GL.Vertex2(x, y);
            GL.TexCoord2(1, 0); GL.Vertex2(x + width, y);
            GL.TexCoord2(1, 1); GL.Vertex2(x + width, y + height);
            GL.TexCoord2(0, 1); GL.Vertex2(x, y + height);
            GL.End();
            GL.Disable(EnableCap.Blend);
        }

        protected override void OnMouseMove(MouseMoveEventArgs e)
        {
            base.OnMouseMove(e);

            // Панель видна только когда курсор находится в нижней
            // области окна, где расположены кнопки и плейлист.
            controlsVisible = isDraggingSeekBar || e.Y >= Height - 100;

            if (isDraggingSeekBar)
                UpdateSeekPreview(e.X);
        }

        protected override void OnMouseDown(MouseButtonEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button != MouseButton.Left) return;
            if (!controlsVisible) return;

            int y = e.Y;
            if (y >= Height - 96 && y <= Height - 80)
            {
                if (mediaDurationSeconds > 0)
                {
                    isDraggingSeekBar = true;
                    UpdateSeekPreview(e.X);
                }
            }
            else if (y >= Height - 66)
            {
                int x = e.X;
                if (x >= 12 && x <= 88) OpenVideo();
                else if (x >= 98 && x <= 174) PlayTrack(Math.Max(0, currentTrackIndex - 1));
                else if (x >= 184 && x <= 288) TogglePause();
                else if (x >= 298 && x <= 374) PlayTrack(Math.Min(playlist.Count - 1, currentTrackIndex + 1));
            }
            else if (e.X >= 490 && playlist.Count > 0)
            {
                int visible = Math.Min(playlist.Count, 4);
                int start = Math.Max(0, Math.Min(currentTrackIndex - 1, playlist.Count - visible));
                int row = (y - (Height - 59)) / 14;
                if (row >= 0 && row < visible)
                    PlayTrack(start + row);
            }
        }

        protected override void OnMouseUp(MouseButtonEventArgs e)
        {
            base.OnMouseUp(e);

            if (e.Button == MouseButton.Left && isDraggingSeekBar)
            {
                isDraggingSeekBar = false;
                SeekTo(seekPreviewSeconds);
            }
        }

        private void OpenVideo()
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "Добавить видео в плейлист";
                dialog.Filter = "Video files|*.mp4;*.mkv;*.avi;*.mov;*.webm|All files|*.*";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    playlist.Add(dialog.FileName);
                    PlayTrack(playlist.Count - 1);
                }
            }
        }

        private void PlayTrack(int index)
        {
            if (index < 0 || index >= playlist.Count || index == currentTrackIndex && isReading)
                return;
            StopPlayback();
            currentTrackIndex = index;
            StartPlayback(playlist[index]);
        }

        private void TogglePause()
        {
            if (audioOutput == null) return;
            if (isPaused)
            {
                audioOutput.Play();
                isPaused = false;
            }
            else
            {
                audioOutput.Pause();
                isPaused = true;
            }
        }

        private void StopPlayback()
        {
            isReading = false;
            lock (frameLock) { frames.Clear(); Monitor.PulseAll(frameLock); }
            try { if (audioOutput != null) { audioOutput.Stop(); audioOutput.Dispose(); } } catch { }
            audioOutput = null;
            try { if (videoProcess != null && !videoProcess.HasExited) videoProcess.Kill(); } catch { }
            try { if (audioProcess != null && !audioProcess.HasExited) audioProcess.Kill(); } catch { }
            try { if (videoThread != null && videoThread.IsAlive) videoThread.Join(1000); } catch { }
            try { if (audioThread != null && audioThread.IsAlive) audioThread.Join(1000); } catch { }
            try { if (videoProcess != null) videoProcess.Dispose(); } catch { }
            try { if (audioProcess != null) audioProcess.Dispose(); } catch { }
            videoProcess = null;
            audioProcess = null;
            videoThread = null;
            audioThread = null;
            audioBuffer = null;
            audioClock = null;
        }

        private void DrawVideo()
        {
            GL.MatrixMode(MatrixMode.Projection);
            GL.PushMatrix();
            GL.LoadIdentity();

            GL.Ortho(0, Width, Height, 0, -1, 1);

            GL.MatrixMode(MatrixMode.Modelview);
            GL.PushMatrix();
            GL.LoadIdentity();

            GL.Enable(EnableCap.Texture2D);

            GL.BindTexture(
                TextureTarget.Texture2D,
                videoTexture);

            GL.Color4(1f, 1f, 1f, 1f);

            GL.Begin(PrimitiveType.Quads);

            GL.TexCoord2(0f, 0f);
            GL.Vertex2(0, 0);

            GL.TexCoord2(1f, 0f);
            GL.Vertex2(Width, 0);

            GL.TexCoord2(1f, 1f);
            GL.Vertex2(Width, Height);

            GL.TexCoord2(0f, 1f);
            GL.Vertex2(0, Height);

            GL.End();

            GL.Disable(EnableCap.Texture2D);

            GL.PopMatrix();

            GL.MatrixMode(MatrixMode.Projection);
            GL.PopMatrix();

            GL.MatrixMode(MatrixMode.Modelview);
        }

        private void UpdateFPSTexture()
        {
            using (Bitmap bitmap = new Bitmap(160, 45))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(
                        Color.FromArgb(190, 20, 20, 30));

                    using (Font font = new Font(
                        "Arial",
                        19,
                        FontStyle.Bold))
                    {
                        graphics.DrawString(
                            "FPS: " + currentFPS,
                            font,
                            Brushes.White,
                            10,
                            7);
                    }
                }

                BitmapData data = bitmap.LockBits(
                    new Rectangle(0, 0, 160, 45),
                    ImageLockMode.ReadOnly,
                    System.Drawing.Imaging.PixelFormat.Format32bppArgb);

                try
                {
                    GL.BindTexture(
                        TextureTarget.Texture2D,
                        fpsTexture);

                    GL.TexImage2D(
                        TextureTarget.Texture2D,
                        0,
                        PixelInternalFormat.Rgba,
                        160,
                        45,
                        0,
                        OpenTK.Graphics.OpenGL.PixelFormat.Bgra,
                        PixelType.UnsignedByte,
                        data.Scan0);
                }
                finally
                {
                    bitmap.UnlockBits(data);
                }
            }
        }

        private void DrawFPS()
        {
            GL.MatrixMode(MatrixMode.Projection);
            GL.PushMatrix();
            GL.LoadIdentity();

            GL.Ortho(0, Width, Height, 0, -1, 1);

            GL.MatrixMode(MatrixMode.Modelview);
            GL.PushMatrix();
            GL.LoadIdentity();

            GL.Enable(EnableCap.Texture2D);
            GL.Enable(EnableCap.Blend);

            GL.BlendFunc(
                BlendingFactor.SrcAlpha,
                BlendingFactor.OneMinusSrcAlpha);

            GL.BindTexture(
                TextureTarget.Texture2D,
                fpsTexture);

            GL.Color4(1f, 1f, 1f, 1f);

            GL.Begin(PrimitiveType.Quads);

            GL.TexCoord2(0f, 0f);
            GL.Vertex2(10, 10);

            GL.TexCoord2(1f, 0f);
            GL.Vertex2(170, 10);

            GL.TexCoord2(1f, 1f);
            GL.Vertex2(170, 55);

            GL.TexCoord2(0f, 1f);
            GL.Vertex2(10, 55);

            GL.End();

            GL.Disable(EnableCap.Blend);
            GL.Disable(EnableCap.Texture2D);

            GL.PopMatrix();

            GL.MatrixMode(MatrixMode.Projection);
            GL.PopMatrix();

            GL.MatrixMode(MatrixMode.Modelview);
        }

        protected override void OnUnload(EventArgs e)
        {
            StopPlayback();

            if (videoTexture != 0)
                GL.DeleteTexture(videoTexture);

            if (fpsTexture != 0)
                GL.DeleteTexture(fpsTexture);

            foreach (int texture in textTextures.Values)
                GL.DeleteTexture(texture);
            textTextures.Clear();

            audioReady.Dispose();
            base.OnUnload(e);
        }
    }
}

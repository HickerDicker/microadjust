using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MicroAdjust
{
    public partial class Form1 : Form
    {
        private const string TimerResDirectory = @"C:\TimerRes";
        private const string SetTimerResolutionExeUrl = "https://github.com/valleyofdoom/TimerResolution/releases/download/SetTimerResolution-v1.0.0/SetTimerResolution.exe";
        private const string MeasureSleepExeUrl = "https://github.com/HickerDicker/SapphireOS/raw/main/src/PostInstall/Tweaks/MeasureSleep.exe";
        private const string ResultsFilePath = @"C:\TimerRes\results.txt";
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect, int nWidthEllipse, int nHeightEllipse);

        public Form1()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.None;
            panel3.MouseDown += Panel3_MouseDown;
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));
        }
        private void Panel3_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            if (!int.TryParse(textBox1.Text, out int startResolution) ||
                !int.TryParse(textBox2.Text, out int endResolution) ||
                !int.TryParse(textBox3.Text, out int increment) ||
                !int.TryParse(textBox4.Text, out int samples))
            {
                MessageBox.Show("Please enter valid numeric values.");
                return;
            }

            try
            {
                if (!Directory.Exists(TimerResDirectory))
                {
                    Directory.CreateDirectory(TimerResDirectory);
                    using (WebClient webClient = new WebClient())
                    {
                        await webClient.DownloadFileTaskAsync(new Uri(SetTimerResolutionExeUrl), Path.Combine(TimerResDirectory, "SetTimerResolution.exe"));
                        await webClient.DownloadFileTaskAsync(new Uri(MeasureSleepExeUrl), Path.Combine(TimerResDirectory, "MeasureSleep.exe"));
                    }
                }

                File.WriteAllText(ResultsFilePath, "RequestedResolutionMs,DeltaMs,STDEV,Max,Min\n");

                for (int currentResolution = startResolution; currentResolution <= endResolution; currentResolution += increment)
                {
                    await BenchmarkResolution(currentResolution, samples);
                }

                MessageBox.Show("Info: results saved in results.txt");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
            finally
            {
                Environment.Exit(0);
            }
        }

        private async Task BenchmarkResolution(int currentResolution, int samples)
        {

            KillProcess("SetTimerResolution.exe");
            KillProcess("MeasureSleep.exe");
            await Task.Delay(100);

            RunCommand("cmd", $"/c {Path.Combine(TimerResDirectory, "SetTimerResolution.exe")} --no-console --resolution {currentResolution}");

            using (Process measureSleepProcess = new Process())
            {
                measureSleepProcess.StartInfo = new ProcessStartInfo
                {
                    FileName = Path.Combine(TimerResDirectory, "MeasureSleep.exe"),
                    Arguments = $"--samples {samples}",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                measureSleepProcess.Start();
                string output = await measureSleepProcess.StandardOutput.ReadToEndAsync();
                measureSleepProcess.WaitForExit();

                ParseAndSaveResults(output, currentResolution);
            }
            
        }

        private void ParseAndSaveResults(string output, int currentResolution)
        {
            double avg = 0, stdev = 0, max = 0, min = 0;
            int reportedResolution = 0;

            foreach (var line in output.Split('\n'))
            {
                var avgMatch = Regex.Match(line, @"Avg: (.*)");
                var stdevMatch = Regex.Match(line, @"STDEV: (.*)");
                var maxMatch = Regex.Match(line, @"Max: (.*)");
                var minMatch = Regex.Match(line, @"Min: (.*)");
                var resolutionMatch = Regex.Match(line, @"Resolution: (\d+)");

                if (avgMatch.Success) avg = double.Parse(avgMatch.Groups[1].Value);
                if (stdevMatch.Success) stdev = double.Parse(stdevMatch.Groups[1].Value);
                if (maxMatch.Success) max = double.Parse(maxMatch.Groups[1].Value);
                if (minMatch.Success) min = double.Parse(minMatch.Groups[1].Value);
                if (resolutionMatch.Success) reportedResolution = int.Parse(resolutionMatch.Groups[1].Value);
            }

            File.AppendAllText(ResultsFilePath, $"{currentResolution}, {avg}, {stdev}, {max}, {min}\n");
        }

        private void RunCommand(string command, string arguments)
        {

            using (Process process = new Process())
            {
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = command,
                    Arguments = arguments,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                process.Start();
            }
        }

        private void KillProcess(string processName)
        {
            foreach (var process in Process.GetProcessesByName(Path.GetFileNameWithoutExtension(processName)))
            {
                process.Kill();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Application.ExitThread();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }
    }
}
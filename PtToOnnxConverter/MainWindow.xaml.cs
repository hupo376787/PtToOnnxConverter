using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;

namespace PtToOnnxConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string ptFile;

        public MainWindow()
        {
            InitializeComponent();

            AppendLog($"I can convert PyTorch file to onnx file. Please ensure python is installed.");
        }

        private async void Select_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Filter = "PyTorch Model (.pt)|*.pt|All Files (.*)|*.*";
            bool? result = dialog.ShowDialog();
            if (result == true)
            {
                string filename = ptFile = dialog.FileName;

                AppendLog($"You selected {filename}");

                var pythonContent = "from ultralytics import YOLO\r\n" +
                    $"model = YOLO(\"{filename.Replace("\\", "/")}\")\r\n" +
                    "model.export(format=\"onnx\")";
                File.WriteAllText("conv.py", pythonContent);

                if (string.IsNullOrEmpty(AutoFindPython()))
                {
                    AppendLog($"Can not find python, please ensure python is installed and add to environment variables");
                    return;
                }

                //安装ultralytics
                await CallTerminal("powershell.exe", "pip install ultralytics");

                //运行conv.py
                await CallTerminal(AutoFindPython()!, "conv.py");
                
                string onnx = Path.GetDirectoryName(ptFile) + "\\" + Path.GetFileNameWithoutExtension(ptFile) + ".onnx";
                if (File.Exists(onnx))
                    AppendLog($"转换完成，输出路径{ptFile}");
            }
        }

        private string? AutoFindPython()
        {
            string userPath = Environment.GetEnvironmentVariable("Path", EnvironmentVariableTarget.User)!;
            string systemPath = Environment.GetEnvironmentVariable("Path", EnvironmentVariableTarget.Machine)!;

            foreach (var item in $"{userPath};{systemPath}".Split(";"))
            {
                var python = item + "/python.exe";
                if (!string.IsNullOrEmpty(item) && File.Exists(python))
                {
                    return python;
                }
            }

            return null;
        }

        #region Python调用
        private Process process;
        public async Task<string> CallTerminal(string exe, string cmd)
        {
            process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "\"" + exe + "\"",
                    Arguments = $"\"{cmd}\"", //路径外面增加双引号，防止有空格出现
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                },
                EnableRaisingEvents = true
            };
            Debug.WriteLine(process.StartInfo.Arguments);

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            process.StartInfo.StandardOutputEncoding = Encoding.UTF8;
            process.StartInfo.StandardErrorEncoding = Encoding.UTF8;
            process.ErrorDataReceived += Process_ErrorDataReceived;
            process.OutputDataReceived += Process_OutputDataReceived;

            process.Exited += (sender, args) =>
            {
                process.Dispose();
            };
            // Start the process asynchronously and wait for it to finish
            process.Start();
            process.BeginErrorReadLine();
            process.BeginOutputReadLine();

            // Wait for the process to exit before continuing with the next one
            await Task.Run(() => process.WaitForExit());

            return "1";
        }

        public void TerminatePython()
        {
            try
            {
                if (process != null)
                {
                    process.Kill();
                }
            }
            catch { }
        }

        private void Process_ErrorDataReceived(object sender, DataReceivedEventArgs e)
        {
            try
            {
                AppendLog(e.Data!);
            }
            catch (Exception err)
            {
                throw err;
            }
        }

        private void Process_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            try
            {
                AppendLog(e.Data!);
            }
            catch (Exception err)
            {
                throw err;
            }
        }

        private void AppendLog(string log)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                tbLog.Text = $"{DateTime.Now}, {log}\r\n" + tbLog.Text;
            });
        }

        #endregion
    }
}
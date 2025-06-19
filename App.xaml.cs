using System.Configuration;
using System.Data;
using System.Windows;
using ModernWpf;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace MarkdownToPdf
{
    /// <summary>
    /// Markdown转PDF应用程序入口
    /// 配置现代化主题和应用程序生命周期
    /// </summary>
    public partial class App : Application
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool FreeConsole();

        protected override void OnStartup(StartupEventArgs e)
        {
            try
            {
                // 分配控制台窗口用于调试输出
                AllocConsole();
                
                // 重定向控制台输出
                Console.WriteLine("=== Markdown转PDF工具 调试控制台 ===");
                Console.WriteLine($"启动时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine($"程序目录: {AppDomain.CurrentDomain.BaseDirectory}");
                Console.WriteLine("=========================================");

                // 设置全局异常处理
                SetupExceptionHandling();

                // 应用ModernWpf主题
                Console.WriteLine("应用ModernWpf主题...");
                ThemeManager.Current.ApplicationTheme = ApplicationTheme.Light;
                Console.WriteLine("主题应用完成");

                base.OnStartup(e);
                Console.WriteLine("应用程序启动完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用程序启动失败: {ex}");
                WriteErrorToFile($"应用程序启动失败: {ex}");
                MessageBox.Show($"应用程序启动失败:\n{ex.Message}", "启动错误", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
                Environment.Exit(1);
            }
        }

        private void SetupExceptionHandling()
        {
            Console.WriteLine("设置全局异常处理...");
            
            // UI线程异常
            DispatcherUnhandledException += (sender, e) =>
            {
                var errorMsg = $"UI线程未处理异常: {e.Exception}";
                Console.WriteLine(errorMsg);
                WriteErrorToFile(errorMsg);
                
                MessageBox.Show($"程序遇到未处理的错误:\n{e.Exception.Message}\n\n详细信息已记录到错误日志", 
                    "程序错误", MessageBoxButton.OK, MessageBoxImage.Error);
                
                e.Handled = true; // 防止程序崩溃
            };

            // 非UI线程异常
            AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
            {
                var errorMsg = $"非UI线程未处理异常: {e.ExceptionObject}";
                Console.WriteLine(errorMsg);
                WriteErrorToFile(errorMsg);
                
                if (!e.IsTerminating)
                {
                    MessageBox.Show($"程序遇到严重错误:\n{e.ExceptionObject}\n\n程序将尝试继续运行", 
                        "严重错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            };

            Console.WriteLine("全局异常处理设置完成");
        }

        private void WriteErrorToFile(string errorMessage)
        {
            try
            {
                var errorFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "error.log");
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                var logEntry = $"[{timestamp}] {errorMessage}{Environment.NewLine}";
                File.AppendAllText(errorFile, logEntry);
            }
            catch
            {
                // 忽略日志写入错误
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            try
            {
                Console.WriteLine("应用程序正在退出...");
                Console.WriteLine($"退出时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                
                base.OnExit(e);
                
                // 释放控制台
                FreeConsole();
            }
            catch (Exception ex)
            {
                WriteErrorToFile($"应用程序退出时出错: {ex}");
            }
        }
    }
} 
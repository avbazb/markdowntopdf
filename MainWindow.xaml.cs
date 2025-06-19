using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using System.Diagnostics;
using MarkdownToPdf.Services;
using MarkdownToPdf.Models;
using ModernWpf;

namespace MarkdownToPdf
{
    /// <summary>
    /// 主窗口交互逻辑
    /// 处理用户界面事件和文件转换流程
    /// </summary>
    public partial class MainWindow : Window
    {
        #region 私有字段

        private readonly MarkdownService _markdownService;
        private readonly PdfService _pdfService;
        private readonly DirectMarkdownToPdfService _directService;
        private ConversionSettings _settings;
        private CancellationTokenSource? _cancellationTokenSource;
        private string? _inputFilePath;
        private string? _outputFilePath;

        #endregion

        #region 构造函数

        public MainWindow()
        {
            try
            {
                WriteDebugLog("开始初始化主窗口...");
                InitializeComponent();
                WriteDebugLog("InitializeComponent完成");
                
                // 初始化服务
                WriteDebugLog("初始化服务...");
                _markdownService = new MarkdownService();
                WriteDebugLog("MarkdownService创建完成");
                _pdfService = new PdfService();
                WriteDebugLog("PdfService创建完成");
                _directService = new DirectMarkdownToPdfService();
                WriteDebugLog("DirectMarkdownToPdfService创建完成");
                _settings = new ConversionSettings();
                WriteDebugLog("ConversionSettings创建完成");
                
                // 订阅进度事件
                WriteDebugLog("订阅进度事件...");
                _markdownService.ProgressChanged += OnProgressChanged;
                _pdfService.ProgressChanged += OnProgressChanged;
                _directService.ProgressChanged += OnDirectProgressChanged;
                
                // 初始化界面
                WriteDebugLog("初始化界面...");
                InitializeUI();
                
                // 启动内存监控
                WriteDebugLog("启动内存监控...");
                StartMemoryMonitoring();
                
                WriteDebugLog("主窗口初始化完成");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"主窗口初始化失败: {ex}");
                MessageBox.Show($"程序启动失败:\n{ex.Message}\n\n详细信息请查看调试输出", "启动错误", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region 界面初始化

        /// <summary>
        /// 初始化用户界面
        /// </summary>
        private void InitializeUI()
        {
            try
            {
                WriteDebugLog("开始初始化UI...");
                
                // 检查控件状态
                CheckControlsState();
                
                // 设置窗口图标（如果有的话）
                try
                {
                    // Icon = new BitmapImage(new Uri("pack://application:,,,/Resources/Icons/app-icon.png"));
                    WriteDebugLog("图标设置跳过");
                }
                catch (Exception ex)
                {
                    WriteDebugLog($"图标加载错误: {ex.Message}");
                }

                // 绑定设置到界面控件
                WriteDebugLog("开始绑定设置到UI...");
                BindSettingsToUI();
                WriteDebugLog("设置绑定完成");
                
                // 更新状态
                WriteDebugLog("更新初始状态...");
                UpdateStatus("应用程序已启动，请选择Markdown文件开始转换");
                
                WriteDebugLog("UI初始化完成");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"UI初始化失败: {ex}");
                throw; // 重新抛出异常以便上层处理
            }
        }

        /// <summary>
        /// 将设置绑定到界面控件
        /// </summary>
        private void BindSettingsToUI()
        {
            try
            {
                // 页面设置
                if (PageSizeComboBox != null)
                {
                    PageSizeComboBox.SelectedIndex = (int)_settings.PageSize;
                }
                
                // 字体设置
                if (FontFamilyComboBox != null)
                {
                    FontFamilyComboBox.SelectedIndex = GetFontFamilyIndex(_settings.FontFamily);
                }
                
                if (FontSizeTextBox != null)
                {
                    FontSizeTextBox.Text = _settings.FontSize.ToString();
                }
                
                // 图片设置
                if (ImageQualityComboBox != null)
                {
                    ImageQualityComboBox.SelectedIndex = GetImageQualityIndex(_settings.ImageQuality);
                }
                
                if (CompressImagesCheckBox != null)
                {
                    CompressImagesCheckBox.IsChecked = _settings.CompressImages;
                }
                
                // 性能设置
                if (LargeFileOptimizationCheckBox != null)
                {
                    LargeFileOptimizationCheckBox.IsChecked = _settings.EnableLargeFileOptimization;
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"界面绑定错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取字体族索引
        /// </summary>
        private int GetFontFamilyIndex(string fontFamily)
        {
            return fontFamily switch
            {
                "Microsoft YaHei" => 0,
                "SimSun" => 1,
                "Arial" => 2,
                _ => 0
            };
        }

        /// <summary>
        /// 获取图片质量索引
        /// </summary>
        private int GetImageQualityIndex(ImageQuality quality)
        {
            return quality switch
            {
                ImageQuality.Low => 0,
                ImageQuality.Medium => 1,
                ImageQuality.High => 2,
                _ => 2
            };
        }

        #endregion

        #region 事件处理器

        /// <summary>
        /// 浏览按钮点击事件
        /// </summary>
        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "选择Markdown文件",
                Filter = "Markdown文件 (*.md;*.markdown)|*.md;*.markdown|所有文件 (*.*)|*.*",
                DefaultExt = ".md"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _inputFilePath = openFileDialog.FileName;
                InputFileTextBox.Text = _inputFilePath;
                
                // 自动设置输出路径
                if (string.IsNullOrEmpty(_outputFilePath))
                {
                    var directory = Path.GetDirectoryName(_inputFilePath);
                    var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(_inputFilePath);
                    _outputFilePath = Path.Combine(directory!, $"{fileNameWithoutExtension}.pdf");
                    OutputFileTextBox.Text = _outputFilePath;
                }
                
                // 加载文件预览
                LoadFilePreview(_inputFilePath);
                
                UpdateStatus($"已选择文件: {Path.GetFileName(_inputFilePath)}");
            }
        }

        /// <summary>
        /// 另存为按钮点击事件
        /// </summary>
        private void SaveAsButton_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog
            {
                Title = "保存PDF文件",
                Filter = "PDF文件 (*.pdf)|*.pdf",
                DefaultExt = ".pdf"
            };

            if (!string.IsNullOrEmpty(_inputFilePath))
            {
                var directory = Path.GetDirectoryName(_inputFilePath);
                var fileName = Path.GetFileNameWithoutExtension(_inputFilePath);
                saveFileDialog.InitialDirectory = directory;
                saveFileDialog.FileName = $"{fileName}.pdf";
            }

            if (saveFileDialog.ShowDialog() == true)
            {
                _outputFilePath = saveFileDialog.FileName;
                OutputFileTextBox.Text = _outputFilePath;
                UpdateStatus($"输出路径: {Path.GetFileName(_outputFilePath)}");
            }
        }

        /// <summary>
        /// 转换按钮点击事件
        /// </summary>
        private async void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                WriteDebugLog("=== 开始转换流程 ===");
                WriteDebugLog($"输入文件: {_inputFilePath ?? "NULL"}");
                WriteDebugLog($"输出文件: {_outputFilePath ?? "NULL"}");
                
                if (string.IsNullOrEmpty(_inputFilePath) || string.IsNullOrEmpty(_outputFilePath))
                {
                    WriteDebugLog("文件路径检查失败");
                    MessageBox.Show("请先选择输入文件和输出路径", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!File.Exists(_inputFilePath))
                {
                    WriteDebugLog("输入文件不存在检查失败");
                    MessageBox.Show("输入文件不存在", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // 更新设置
                WriteDebugLog("更新UI设置...");
                UpdateSettingsFromUI();
                WriteDebugLog("UI设置更新完成");
                
                // 开始转换
                WriteDebugLog("开始异步转换...");
                await StartConversionAsync();
                WriteDebugLog("异步转换完成");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"转换过程异常: {ex}");
                MessageBox.Show($"转换过程中发生错误: {ex.Message}\n\n详细错误信息请查看调试日志", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                ResetConversionUI();
            }
        }

        /// <summary>
        /// 打开PDF文件按钮点击事件
        /// </summary>
        private void OpenOutputButton_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_outputFilePath) && File.Exists(_outputFilePath))
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = _outputFilePath,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"无法打开PDF文件: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        /// <summary>
        /// 打开文件夹按钮点击事件
        /// </summary>
        private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_outputFilePath))
            {
                try
                {
                    var directory = Path.GetDirectoryName(_outputFilePath);
                    if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory))
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = directory,
                            UseShellExecute = true
                        });
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"无法打开文件夹: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        #endregion

        #region 转换流程

        /// <summary>
        /// 开始转换流程
        /// </summary>
        private async Task StartConversionAsync()
        {
            try
            {
                WriteDebugLog("=== 开始转换流程内部 ===");
                
                // 设置UI状态
                WriteDebugLog("设置转换UI状态...");
                SetConversionUI(true);
                _cancellationTokenSource = new CancellationTokenSource();
                WriteDebugLog("UI状态设置完成，创建取消令牌");

                UpdateStatus("开始直接转换Markdown到PDF...");
                WriteDebugLog("开始直接Markdown到PDF转换...");
                
                // 直接转换Markdown为PDF（跳过HTML步骤）
                await _directService.ConvertMarkdownToPdfAsync(_inputFilePath!, _outputFilePath!, _settings, _cancellationTokenSource.Token);
                WriteDebugLog("直接PDF转换完成");

                // 转换完成
                UpdateStatus("转换完成!");
                WriteDebugLog("显示成功消息...");
                MessageBox.Show("PDF文件生成成功!", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
                
                // 启用操作按钮
                WriteDebugLog("启用操作按钮...");
                if (OpenOutputButton != null) OpenOutputButton.IsEnabled = true;
                if (OpenFolderButton != null) OpenFolderButton.IsEnabled = true;
                WriteDebugLog("=== 转换流程完成 ===");
            }
            catch (OperationCanceledException ex)
            {
                WriteDebugLog($"转换被取消: {ex.Message}");
                UpdateStatus("转换已取消");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"转换过程中发生异常: {ex}");
                UpdateStatus("转换失败");
                throw;
            }
            finally
            {
                WriteDebugLog("重置转换UI状态...");
                ResetConversionUI();
                WriteDebugLog("转换流程清理完成");
            }
        }

        /// <summary>
        /// 设置转换期间的UI状态
        /// </summary>
        private void SetConversionUI(bool isConverting)
        {
            try
            {
                if (ConvertButton != null) ConvertButton.IsEnabled = !isConverting;
                if (BrowseButton != null) BrowseButton.IsEnabled = !isConverting;
                if (SaveAsButton != null) SaveAsButton.IsEnabled = !isConverting;
                if (ProgressPanel != null) ProgressPanel.Visibility = isConverting ? Visibility.Visible : Visibility.Collapsed;
                
                if (isConverting)
                {
                    if (ConversionProgressBar != null) ConversionProgressBar.Value = 0;
                    if (ProgressTextBlock != null) ProgressTextBlock.Text = "准备开始...";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UI状态设置错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 重置转换UI状态
        /// </summary>
        private void ResetConversionUI()
        {
            SetConversionUI(false);
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 加载文件预览
        /// </summary>
        private async void LoadFilePreview(string filePath)
        {
            try
            {
                if (PreviewTextBox == null) return;
                
                if (File.Exists(filePath))
                {
                    var content = await File.ReadAllTextAsync(filePath);
                    
                    // 限制预览内容长度
                    if (content.Length > 5000)
                    {
                        content = content.Substring(0, 5000) + "\n\n... (文件内容过长，仅显示前5000个字符)";
                    }
                    
                    PreviewTextBox.Text = content;
                }
                else
                {
                    PreviewTextBox.Text = "文件不存在";
                }
            }
            catch (Exception ex)
            {
                if (PreviewTextBox != null)
                {
                    PreviewTextBox.Text = $"无法加载文件预览: {ex.Message}";
                }
            }
        }

        /// <summary>
        /// 从界面更新设置
        /// </summary>
        private void UpdateSettingsFromUI()
        {
            try
            {
                // 页面设置
                if (PageSizeComboBox?.SelectedIndex >= 0)
                {
                    _settings.PageSize = (PageSize)PageSizeComboBox.SelectedIndex;
                }
                
                // 字体设置
                _settings.FontFamily = GetSelectedFontFamily();
                if (!string.IsNullOrEmpty(FontSizeTextBox?.Text) && 
                    double.TryParse(FontSizeTextBox.Text, out var fontSize))
                {
                    _settings.FontSize = fontSize;
                }
                
                // 图片设置
                _settings.ImageQuality = GetSelectedImageQuality();
                _settings.CompressImages = CompressImagesCheckBox?.IsChecked ?? true;
                
                // 性能设置
                _settings.EnableLargeFileOptimization = LargeFileOptimizationCheckBox?.IsChecked ?? true;
            }
            catch (Exception ex)
            {
                UpdateStatus($"更新设置时出错: {ex.Message}");
                // 使用默认设置
                _settings = new ConversionSettings();
            }
        }

        /// <summary>
        /// 获取选中的字体族
        /// </summary>
        private string GetSelectedFontFamily()
        {
            try
            {
                return (FontFamilyComboBox?.SelectedIndex ?? 0) switch
                {
                    0 => "Microsoft YaHei",
                    1 => "SimSun",
                    2 => "Arial",
                    _ => "Microsoft YaHei"
                };
            }
            catch
            {
                return "Microsoft YaHei";
            }
        }

        /// <summary>
        /// 获取选中的图片质量
        /// </summary>
        private ImageQuality GetSelectedImageQuality()
        {
            try
            {
                return (ImageQualityComboBox?.SelectedIndex ?? 2) switch
                {
                    0 => ImageQuality.Low,
                    1 => ImageQuality.Medium,
                    2 => ImageQuality.High,
                    _ => ImageQuality.High
                };
            }
            catch
            {
                return ImageQuality.High;
            }
        }

        /// <summary>
        /// 更新状态栏
        /// </summary>
        private void UpdateStatus(string message)
        {
            try
            {
                if (StatusTextBlock != null)
                {
                    StatusTextBlock.Text = message ?? "状态未知";
                }
            }
            catch (Exception ex)
            {
                // 记录错误但不影响主流程
                System.Diagnostics.Debug.WriteLine($"状态更新错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 启动内存监控
        /// </summary>
        private void StartMemoryMonitoring()
        {
            var timer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(2)
            };
            
            timer.Tick += (sender, e) =>
            {
                var memoryUsage = GC.GetTotalMemory(false) / 1024 / 1024; // MB
                // MemoryUsageTextBlock.Text = $"当前内存使用: {memoryUsage} MB";
            };
            
            timer.Start();
        }

        #endregion

        #region 进度事件处理

        /// <summary>
        /// 进度变化事件处理
        /// </summary>
        private void OnProgressChanged(object? sender, EventArgs e)
        {
            try
            {
                Dispatcher.Invoke(() =>
                {
                    try
                    {
                        int progress = 0;
                        string message = "处理中...";

                        // 处理不同类型的进度事件
                        switch (e)
                        {
                            case MarkdownProcessingProgressEventArgs markdownProgress:
                                progress = markdownProgress.Progress / 2; // Markdown处理占50%
                                message = $"Markdown处理: {markdownProgress.Message}";
                                break;
                            case PdfGenerationProgressEventArgs pdfProgress:
                                progress = 50 + pdfProgress.Progress / 2; // PDF生成占50%
                                message = $"PDF生成: {pdfProgress.Message}";
                                break;
                        }

                        if (ConversionProgressBar != null)
                        {
                            ConversionProgressBar.Value = Math.Max(0, Math.Min(100, progress));
                        }
                        
                        if (ProgressTextBlock != null)
                        {
                            ProgressTextBlock.Text = message ?? "处理中...";
                        }
                    }
                    catch (Exception uiEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"UI更新错误: {uiEx.Message}");
                    }
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"进度事件处理错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 直接转换进度变化事件处理
        /// </summary>
        private void OnDirectProgressChanged(object? sender, DirectPdfGenerationProgressEventArgs e)
        {
            try
            {
                Dispatcher.Invoke(() =>
                {
                    try
                    {
                        if (ConversionProgressBar != null)
                        {
                            ConversionProgressBar.Value = Math.Max(0, Math.Min(100, e.Progress));
                        }
                        
                        if (ProgressTextBlock != null)
                        {
                            ProgressTextBlock.Text = e.Message ?? "处理中...";
                        }

                        WriteDebugLog($"直接转换进度: {e.Progress}% - {e.Message}");
                    }
                    catch (Exception uiEx)
                    {
                        WriteDebugLog($"直接转换UI更新错误: {uiEx.Message}");
                    }
                });
            }
            catch (Exception ex)
            {
                WriteDebugLog($"直接转换进度事件处理错误: {ex.Message}");
            }
        }

        #endregion

        #region 调试和日志

        /// <summary>
        /// 写入调试日志
        /// </summary>
        private void WriteDebugLog(string message)
        {
            try
            {
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                var logMessage = $"[{timestamp}] {message}";
                
                // 输出到控制台
                Console.WriteLine(logMessage);
                
                // 输出到调试窗口
                System.Diagnostics.Debug.WriteLine(logMessage);
                
                // 可选：写入文件
                var logFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "debug.log");
                File.AppendAllText(logFile, logMessage + Environment.NewLine);
            }
            catch
            {
                // 忽略日志写入错误
            }
        }

        /// <summary>
        /// 检查控件状态
        /// </summary>
        private void CheckControlsState()
        {
            WriteDebugLog("=== 检查控件状态 ===");
            WriteDebugLog($"PageSizeComboBox: {(PageSizeComboBox == null ? "NULL" : "OK")}");
            WriteDebugLog($"FontFamilyComboBox: {(FontFamilyComboBox == null ? "NULL" : "OK")}");
            WriteDebugLog($"FontSizeTextBox: {(FontSizeTextBox == null ? "NULL" : "OK")}");
            WriteDebugLog($"ImageQualityComboBox: {(ImageQualityComboBox == null ? "NULL" : "OK")}");
            WriteDebugLog($"CompressImagesCheckBox: {(CompressImagesCheckBox == null ? "NULL" : "OK")}");
            WriteDebugLog($"LargeFileOptimizationCheckBox: {(LargeFileOptimizationCheckBox == null ? "NULL" : "OK")}");
            WriteDebugLog($"StatusTextBlock: {(StatusTextBlock == null ? "NULL" : "OK")}");
            WriteDebugLog($"ProgressPanel: {(ProgressPanel == null ? "NULL" : "OK")}");
            WriteDebugLog($"ConversionProgressBar: {(ConversionProgressBar == null ? "NULL" : "OK")}");
            WriteDebugLog($"ProgressTextBlock: {(ProgressTextBlock == null ? "NULL" : "OK")}");
            WriteDebugLog($"PreviewTextBox: {(PreviewTextBox == null ? "NULL" : "OK")}");
            WriteDebugLog("=== 控件状态检查完成 ===");
        }

        #endregion

        #region 窗口生命周期

        protected override void OnClosed(EventArgs e)
        {
            try
            {
                WriteDebugLog("开始清理资源...");
                
                // 清理资源
                _cancellationTokenSource?.Cancel();
                _cancellationTokenSource?.Dispose();
                _markdownService?.Dispose();
                
                WriteDebugLog("资源清理完成");
            }
            catch (Exception ex)
            {
                WriteDebugLog($"资源清理时出错: {ex.Message}");
            }
            
            base.OnClosed(e);
        }

        #endregion
    }
} 
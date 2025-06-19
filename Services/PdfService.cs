using System;
using System.IO;
using System.Threading.Tasks;
using System.Threading;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using MarkdownToPdf.Models;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.Collections.Concurrent;
using System.Text;

namespace MarkdownToPdf.Services
{
    /// <summary>
    /// PDF生成服务
    /// 负责将HTML转换为PDF文档
    /// 支持中文字体、图片嵌入、大文件优化等功能
    /// </summary>
    public class PdfService
    {
        private readonly ImageProcessor _imageProcessor;

        /// <summary>
        /// 转换进度事件
        /// </summary>
        public event EventHandler<PdfGenerationProgressEventArgs>? ProgressChanged;

        public PdfService()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("初始化PdfService...");
                _imageProcessor = new ImageProcessor();
                
                // 配置PdfSharp字体解析器，支持系统字体
                System.Diagnostics.Debug.WriteLine("配置字体解析器...");
                if (GlobalFontSettings.FontResolver == null)
                {
                    var fontResolver = new SystemFontResolver();
                    GlobalFontSettings.FontResolver = fontResolver;
                    System.Diagnostics.Debug.WriteLine("字体解析器配置完成");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("字体解析器已存在，跳过配置");
                }

                // 测试字体系统
                TestFontSystem();
                
                System.Diagnostics.Debug.WriteLine("PdfService初始化完成");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"PdfService初始化失败: {ex}");
                throw new InvalidOperationException($"PDF服务初始化失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 测试字体系统是否正常工作
        /// </summary>
        private void TestFontSystem()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("测试字体系统...");
                
                // 尝试创建一些常用字体来验证系统
                var testFonts = new[] { "Microsoft YaHei", "SimSun", "Arial" };
                
                foreach (var fontName in testFonts)
                {
                    try
                    {
                        var testFont = new XFont(fontName, 12, XFontStyleEx.Regular);
                        System.Diagnostics.Debug.WriteLine($"字体测试成功: {fontName}");
                        // XFont不需要手动Dispose
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"字体测试失败 {fontName}: {ex.Message}");
                    }
                }
                
                System.Diagnostics.Debug.WriteLine("字体系统测试完成");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"字体系统测试异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 将HTML内容转换为PDF文件
        /// </summary>
        /// <param name="htmlContent">HTML内容</param>
        /// <param name="outputPath">输出PDF文件路径</param>
        /// <param name="settings">转换设置</param>
        /// <param name="cancellationToken">取消令牌</param>
        /// <returns>转换任务</returns>
        public async Task ConvertToPdfAsync(string htmlContent, string outputPath, 
            ConversionSettings settings, CancellationToken cancellationToken = default)
        {
            try
            {
                OnProgressChanged(new PdfGenerationProgressEventArgs(0, "开始生成PDF..."));

                // 创建PDF文档
                using (var document = new PdfDocument())
                {
                    document.Info.Title = "Markdown转PDF";
                    document.Info.Creator = "MarkdownToPdf工具";
                    document.Info.CreationDate = DateTime.Now;

                    // 如果启用大文件优化，则分块处理
                    if (settings.EnableLargeFileOptimization && IsLargeContent(htmlContent))
                    {
                        await ProcessLargeFileAsync(document, htmlContent, settings, cancellationToken);
                    }
                    else
                    {
                        await ProcessContentAsync(document, htmlContent, settings, cancellationToken);
                    }

                    OnProgressChanged(new PdfGenerationProgressEventArgs(90, "保存PDF文件..."));

                    // 确保输出目录存在
                    var outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // 保存PDF文档
                    document.Save(outputPath);
                    
                    OnProgressChanged(new PdfGenerationProgressEventArgs(100, "PDF生成完成"));
                }
            }
            catch (OperationCanceledException)
            {
                throw; // 重新抛出取消异常
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"生成PDF时发生错误: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 判断是否为大内容（需要分块处理）
        /// </summary>
        private bool IsLargeContent(string content)
        {
            // 简单判断：超过50KB或包含大量图片
            return content.Length > 50000 || 
                   Regex.Matches(content, @"<img[^>]*>", RegexOptions.IgnoreCase).Count > 20;
        }

        /// <summary>
        /// 大文件优化转换（并行分块处理）
        /// </summary>
        private async Task ProcessLargeFileAsync(PdfDocument document, string htmlContent, 
            ConversionSettings settings, CancellationToken cancellationToken)
        {
            OnProgressChanged(new PdfGenerationProgressEventArgs(10, "分析大文件结构..."));

            // 将HTML内容按章节分块，增加分块数量以提高并行度
            var chunks = SplitHtmlIntoChunks(htmlContent, settings.EnableLargeFileOptimization ? 50 : 20);
            var totalChunks = chunks.Count;

            if (totalChunks <= 1)
            {
                await ProcessRegularFileAsync(document, htmlContent, settings, cancellationToken);
                return;
            }

            OnProgressChanged(new PdfGenerationProgressEventArgs(15, $"分割为 {totalChunks} 个处理块..."));

            // 根据系统资源和设置动态调整并行度
            var baseParallelism = settings.Performance?.MaxParallelPdfProcessing ?? Environment.ProcessorCount;
            var maxParallelism = Math.Min(baseParallelism, Math.Max(2, totalChunks / 5));
            var semaphore = new SemaphoreSlim(maxParallelism, maxParallelism);
            
            // 创建临时页面集合用于并行处理
            var pageResults = new ConcurrentDictionary<int, List<string>>();
            var tasks = new List<Task>();

            for (int i = 0; i < totalChunks; i++)
            {
                var chunkIndex = i;
                var chunk = chunks[i];
                
                var task = Task.Run(async () =>
                {
                    await semaphore.WaitAsync(cancellationToken);
                    try
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        // 创建临时文档用于处理单个块
                        using var tempDocument = new PdfDocument();
                        await ProcessChunkAsync(tempDocument, chunk, settings, cancellationToken);

                        // 将页面信息保存到结果集合（而不是直接克隆页面对象）
                        var pageContents = new List<string>();
                        for (int pageIndex = 0; pageIndex < tempDocument.PageCount; pageIndex++)
                        {
                            // 这里我们保存块内容而不是页面对象，稍后重新渲染
                            pageContents.Add(chunk);
                        }
                        pageResults[chunkIndex] = pageContents;

                        var progress = 15 + (chunkIndex * 70 / totalChunks);
                        OnProgressChanged(new PdfGenerationProgressEventArgs(progress, $"处理第 {chunkIndex + 1}/{totalChunks} 块内容..."));
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                });

                tasks.Add(task);
                
                // 控制内存使用，每启动一定数量的任务后等待部分完成
                if (tasks.Count >= maxParallelism * 2)
                {
                    await Task.WhenAny(tasks);
                    tasks.RemoveAll(t => t.IsCompleted);
                }
            }

            // 等待所有任务完成
            await Task.WhenAll(tasks);

            OnProgressChanged(new PdfGenerationProgressEventArgs(85, "合并处理结果..."));

            // 按顺序合并所有页面
            for (int i = 0; i < totalChunks; i++)
            {
                if (pageResults.TryGetValue(i, out var pageContents))
                {
                    foreach (var content in pageContents)
                    {
                        // 重新处理每个块并添加到主文档
                        await ProcessChunkAsync(document, content, settings, cancellationToken);
                    }
                }
            }

            semaphore.Dispose();
        }

        /// <summary>
        /// 处理普通内容
        /// </summary>
        private async Task ProcessContentAsync(PdfDocument document, string htmlContent, 
            ConversionSettings settings, CancellationToken cancellationToken)
        {
            OnProgressChanged(new PdfGenerationProgressEventArgs(20, "解析HTML内容..."));
            await ProcessChunkAsync(document, htmlContent, settings, cancellationToken);
        }

        /// <summary>
        /// 处理内容块
        /// </summary>
        private async Task ProcessChunkAsync(PdfDocument document, string htmlChunk, 
            ConversionSettings settings, CancellationToken cancellationToken)
        {
            // 解析HTML内容
            var elements = ParseHtmlElements(htmlChunk);
            
            // 创建页面
            var page = document.AddPage();
            SetupPage(page, settings);
            
            var gfx = XGraphics.FromPdfPage(page);
            var currentY = settings.MarginTop;
            var pageHeight = page.Height - settings.MarginBottom;

            foreach (var element in elements)
            {
                cancellationToken.ThrowIfCancellationRequested();

                // 检查是否需要新页面
                if (currentY > pageHeight - 50) // 预留50个单位的底部空间
                {
                    gfx.Dispose();
                    page = document.AddPage();
                    SetupPage(page, settings);
                    gfx = XGraphics.FromPdfPage(page);
                    currentY = settings.MarginTop;
                }

                // 渲染元素
                currentY = await RenderElementAsync(gfx, element, settings, currentY, cancellationToken);
            }

            gfx.Dispose();
        }

        /// <summary>
        /// 设置页面属性
        /// </summary>
        private void SetupPage(PdfPage page, ConversionSettings settings)
        {
            // 设置页面大小
            switch (settings.PageSize)
            {
                case PageSize.A4:
                    page.Size = PdfSharp.PageSize.A4;
                    break;
                case PageSize.A3:
                    page.Size = PdfSharp.PageSize.A3;
                    break;
                case PageSize.A5:
                    page.Size = PdfSharp.PageSize.A5;
                    break;
                case PageSize.Letter:
                    page.Size = PdfSharp.PageSize.Letter;
                    break;
                case PageSize.Legal:
                    page.Size = PdfSharp.PageSize.Legal;
                    break;
            }

            page.Orientation = PdfSharp.PageOrientation.Portrait;
        }

        /// <summary>
        /// 将HTML内容分割为更小的块以支持并行处理
        /// </summary>
        /// <param name="htmlContent">HTML内容</param>
        /// <param name="maxChunkSize">最大块大小</param>
        /// <returns>HTML块列表</returns>
        private List<string> SplitHtmlIntoChunks(string htmlContent, int maxChunkSize = 50)
        {
            var chunks = new List<string>();
            
            if (string.IsNullOrEmpty(htmlContent))
            {
                return chunks;
            }

            // 按照标题标签分割内容
            var headingPattern = @"<h[1-6][^>]*>.*?</h[1-6]>";
            var matches = System.Text.RegularExpressions.Regex.Matches(htmlContent, headingPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Singleline);
            
            if (matches.Count == 0)
            {
                // 没有标题时按段落分割
                return SplitByParagraphs(htmlContent, maxChunkSize);
            }

            var currentChunk = new StringBuilder();
            var lastIndex = 0;
            var chunkCount = 0;

            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                // 添加当前标题之前的内容
                if (match.Index > lastIndex)
                {
                    var beforeHeading = htmlContent.Substring(lastIndex, match.Index - lastIndex);
                    currentChunk.Append(beforeHeading);
                }

                // 如果当前块太大，先保存
                if (currentChunk.Length > 0 && chunkCount >= maxChunkSize)
                {
                    chunks.Add(WrapHtmlContent(currentChunk.ToString()));
                    currentChunk.Clear();
                    chunkCount = 0;
                }

                // 找到下一个标题或内容结束
                var nextMatch = match.NextMatch();
                var sectionEnd = nextMatch.Success ? nextMatch.Index : htmlContent.Length;
                var section = htmlContent.Substring(match.Index, sectionEnd - match.Index);
                
                currentChunk.Append(section);
                chunkCount++;
                
                lastIndex = sectionEnd;
            }

            // 添加剩余内容
            if (currentChunk.Length > 0)
            {
                chunks.Add(WrapHtmlContent(currentChunk.ToString()));
            }

            return chunks.Count > 0 ? chunks : new List<string> { htmlContent };
        }

        /// <summary>
        /// 按段落分割HTML内容
        /// </summary>
        private List<string> SplitByParagraphs(string htmlContent, int maxChunkSize)
        {
            var chunks = new List<string>();
            var paragraphs = htmlContent.Split(new[] { "</p>", "<br>", "<hr>" }, StringSplitOptions.RemoveEmptyEntries);
            
            var currentChunk = new StringBuilder();
            var paragraphCount = 0;

            foreach (var paragraph in paragraphs)
            {
                if (paragraphCount >= maxChunkSize && currentChunk.Length > 0)
                {
                    chunks.Add(WrapHtmlContent(currentChunk.ToString()));
                    currentChunk.Clear();
                    paragraphCount = 0;
                }

                currentChunk.Append(paragraph);
                if (!paragraph.EndsWith("</p>") && !paragraph.EndsWith("<br>") && !paragraph.EndsWith("<hr>"))
                {
                    currentChunk.Append("</p>");
                }
                paragraphCount++;
            }

            if (currentChunk.Length > 0)
            {
                chunks.Add(WrapHtmlContent(currentChunk.ToString()));
            }

            return chunks.Count > 0 ? chunks : new List<string> { htmlContent };
        }

        /// <summary>
        /// 包装HTML内容，确保有正确的HTML结构
        /// </summary>
        private string WrapHtmlContent(string content)
        {
            if (string.IsNullOrWhiteSpace(content))
                return content;

            // 如果内容已经有完整的HTML结构，直接返回
            if (content.TrimStart().StartsWith("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) || 
                content.TrimStart().StartsWith("<html", StringComparison.OrdinalIgnoreCase))
            {
                return content;
            }

            // 包装为完整的HTML
            return $@"<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <style>
        body {{ margin: 40px; font-family: 'Microsoft YaHei', SimSun, Arial, sans-serif; line-height: 1.6; }}
        img {{ max-width: 100%; height: auto; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; }}
    </style>
</head>
<body>
{content}
</body>
</html>";
        }

        /// <summary>
        /// 解析HTML元素
        /// </summary>
        private List<HtmlElement> ParseHtmlElements(string html)
        {
            var elements = new List<HtmlElement>();
            
            // 简单的HTML解析（实际项目中建议使用HtmlAgilityPack）
            var textContent = Regex.Replace(html, @"<[^>]+>", "", RegexOptions.IgnoreCase);
            
            // 按行分割文本
            var lines = textContent.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var line in lines)
            {
                if (!string.IsNullOrWhiteSpace(line))
                {
                    elements.Add(new HtmlElement 
                    { 
                        Type = "text", 
                        Content = line.Trim() 
                    });
                }
            }

            return elements;
        }

        /// <summary>
        /// 渲染HTML元素
        /// </summary>
        private async Task<double> RenderElementAsync(XGraphics gfx, HtmlElement element, 
            ConversionSettings settings, double currentY, CancellationToken cancellationToken)
        {
            // 创建字体，添加错误处理和回退机制
            var font = CreateSafeFont(settings.FontFamily, settings.FontSize, XFontStyleEx.Regular);
            var brush = XBrushes.Black;
            var rect = new XRect(settings.MarginLeft, currentY, 
                gfx.PageSize.Width - settings.MarginLeft - settings.MarginRight, 
                settings.FontSize * settings.LineHeight);

            switch (element.Type.ToLower())
            {
                case "text":
                    // 渲染文本，支持自动换行
                    var textHeight = RenderTextWithWrapping(gfx, element.Content, font, brush, rect, settings);
                    return currentY + textHeight + (settings.FontSize * 0.5); // 添加段落间距

                case "image":
                    // 渲染图片
                    var imageHeight = await RenderImageAsync(gfx, element.Content, rect, settings, cancellationToken);
                    return currentY + imageHeight + 10; // 添加图片间距

                default:
                    return currentY;
            }
        }

        /// <summary>
        /// 安全创建字体，包含回退机制
        /// </summary>
        private XFont CreateSafeFont(string fontFamily, double fontSize, XFontStyleEx style)
        {
            var fallbackFonts = new[]
            {
                fontFamily,
                "Microsoft YaHei",  // 微软雅黑
                "SimSun",           // 宋体  
                "SimHei",           // 黑体
                "Arial Unicode MS", // Arial Unicode (支持中文)
                "Noto Sans CJK SC", // Google Noto中文字体
                "PingFang SC",      // 苹果中文字体
                "Microsoft JhengHei", // 微软正黑体
                "Arial",            // Arial
                "Calibri",          // Calibri
                "Verdana",          // Verdana
                "Times New Roman"   // Times New Roman
            };

            foreach (var font in fallbackFonts.Distinct())
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine($"尝试创建字体: {font}, 大小: {fontSize}, 样式: {style}");
                    var xFont = new XFont(font, fontSize, style);
                    
                    // 验证字体是否能正常工作
                    if (ValidateFont(xFont))
                    {
                        System.Diagnostics.Debug.WriteLine($"字体创建并验证成功: {font}");
                        return xFont;
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"字体验证失败: {font}");
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"字体创建失败 {font}: {ex.Message}");
                    continue;
                }
            }

            // 最后的回退：使用最基本的系统字体
            var emergencyFonts = new[] { "Tahoma", "Segoe UI", "Consolas", "Courier New" };
            
            foreach (var emergencyFont in emergencyFonts)
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine($"使用紧急回退字体: {emergencyFont}");
                    var xFont = new XFont(emergencyFont, fontSize, style);
                    
                    if (ValidateFont(xFont))
                    {
                        System.Diagnostics.Debug.WriteLine($"紧急字体验证成功: {emergencyFont}");
                        return xFont;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"紧急字体失败 {emergencyFont}: {ex.Message}");
                }
            }

            // 最后的最后：使用固定参数
            try
            {
                System.Diagnostics.Debug.WriteLine("使用最小回退字体: Arial 12pt Regular");
                var finalFont = new XFont("Arial", 12, XFontStyleEx.Regular);
                System.Diagnostics.Debug.WriteLine("最小回退字体创建成功");
                return finalFont;
            }
            catch (Exception finalEx)
            {
                throw new FontNotFoundException($"无法创建任何字体。请检查系统字体配置。最后错误: {finalEx.Message}", finalEx);
            }
        }

        /// <summary>
        /// 验证字体是否能正常工作
        /// </summary>
        private bool ValidateFont(XFont font)
        {
            try
            {
                // 创建一个小的测试文档来验证字体
                using (var testDoc = new PdfDocument())
                {
                    var testPage = testDoc.AddPage();
                    testPage.Width = 100;
                    testPage.Height = 100;
                    
                    using (var testGfx = XGraphics.FromPdfPage(testPage))
                    {
                        // 测试基本字符和中文字符
                        var testStrings = new[] 
                        { 
                            "A", "a", "1", " ",     // 基本ASCII字符
                            "中", "文", "测", "试",   // 中文字符
                            "。", "，", "？", "！"    // 中文标点
                        };
                        
                        foreach (var testStr in testStrings)
                        {
                            try
                            {
                                var size = testGfx.MeasureString(testStr, font);
                                if (size.Width <= 0 || size.Height <= 0 || double.IsNaN(size.Width) || double.IsNaN(size.Height))
                                {
                                    System.Diagnostics.Debug.WriteLine($"字体测量结果无效: '{testStr}' -> {size.Width}x{size.Height}");
                                    return false;
                                }
                                
                                // 尝试绘制到一个位置
                                testGfx.DrawString(testStr, font, XBrushes.Black, 10, 10);
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"字体测试字符失败: '{testStr}' - {ex.Message}");
                                
                                // 对于中文字符，如果失败就跳过这个字体
                                if (IsCJKCharacter(testStr.FirstOrDefault()))
                                {
                                    return false;
                                }
                                // 对于基本字符，如果失败也跳过
                                return false;
                            }
                        }
                    }
                }
                
                System.Diagnostics.Debug.WriteLine("字体验证通过");
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"字体验证过程失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 渲染带换行的文本（增强版，支持中文和特殊字符处理）
        /// </summary>
        private double RenderTextWithWrapping(XGraphics gfx, string text, XFont font, XBrush brush, XRect rect, ConversionSettings settings)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"开始渲染文本，长度: {text.Length}");
                
                // 清理和预处理文本
                var cleanText = CleanTextForRendering(text);
                System.Diagnostics.Debug.WriteLine($"清理后文本长度: {cleanText.Length}");
                
                // 按照中文友好的方式分词
                var segments = SplitTextForChinese(cleanText);
                
                var currentLine = string.Empty;
                var lineHeight = settings.FontSize * settings.LineHeight;
                var totalHeight = 0.0;
                var y = rect.Top;

                foreach (var segment in segments)
                {
                    if (string.IsNullOrWhiteSpace(segment)) continue;
                    
                    var testLine = string.IsNullOrEmpty(currentLine) ? segment : $"{currentLine}{segment}";
                    
                    // 安全测量字符串
                    var testSize = SafeMeasureString(gfx, testLine, font);
                    
                    if (testSize.Width <= rect.Width && testSize.Width > 0)
                    {
                        currentLine = testLine;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(currentLine))
                        {
                            // 安全绘制当前行
                            SafeDrawString(gfx, currentLine, font, brush, rect.Left, y);
                            y += lineHeight;
                            totalHeight += lineHeight;
                        }
                        
                        // 处理超长单词/句子
                        if (SafeMeasureString(gfx, segment, font).Width > rect.Width)
                        {
                            var brokenLines = BreakLongText(gfx, segment, font, rect.Width);
                            foreach (var line in brokenLines)
                            {
                                SafeDrawString(gfx, line, font, brush, rect.Left, y);
                                y += lineHeight;
                                totalHeight += lineHeight;
                            }
                            currentLine = string.Empty;
                        }
                        else
                        {
                            currentLine = segment;
                        }
                    }
                }

                // 渲染最后一行
                if (!string.IsNullOrEmpty(currentLine))
                {
                    SafeDrawString(gfx, currentLine, font, brush, rect.Left, y);
                    totalHeight += lineHeight;
                }

                System.Diagnostics.Debug.WriteLine($"文本渲染完成，总高度: {totalHeight}");
                return Math.Max(totalHeight, lineHeight);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"文本渲染失败: {ex.Message}");
                // 回退到简单文本渲染
                try
                {
                    var lineHeight = settings.FontSize * settings.LineHeight;
                    SafeDrawString(gfx, "[文本渲染失败]", font, brush, rect.Left, rect.Top);
                    return lineHeight;
                }
                catch
                {
                    return settings.FontSize * settings.LineHeight;
                }
            }
        }

        /// <summary>
        /// 清理文本中可能导致问题的字符
        /// </summary>
        private string CleanTextForRendering(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            
            try
            {
                // 移除或替换问题字符
                var cleanText = text
                    // 移除控制字符（除了换行和制表符）
                    .Where(c => !char.IsControl(c) || c == '\n' || c == '\t' || c == '\r')
                    // 替换某些可能有问题的Unicode字符
                    .Select(c => c switch
                    {
                        '\u00A0' => ' ',  // 不间断空格替换为普通空格
                        '\u2028' => '\n', // 行分隔符
                        '\u2029' => '\n', // 段落分隔符
                        '\uFEFF' => ' ',  // 零宽度无断空格
                        _ => c
                    })
                    .ToArray();
                
                return new string(cleanText);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"文本清理失败: {ex.Message}");
                // 如果清理失败，尝试基本处理
                return text?.Replace('\0', ' ') ?? string.Empty;
            }
        }

        /// <summary>
        /// 中文友好的文本分割
        /// </summary>
        private List<string> SplitTextForChinese(string text)
        {
            var segments = new List<string>();
            
            if (string.IsNullOrEmpty(text)) return segments;
            
            try
            {
                // 对于中文文本，我们按字符和标点符号分割
                var currentSegment = new StringBuilder();
                
                for (int i = 0; i < text.Length; i++)
                {
                    var c = text[i];
                    
                    if (char.IsWhiteSpace(c))
                    {
                        if (currentSegment.Length > 0)
                        {
                            segments.Add(currentSegment.ToString());
                            currentSegment.Clear();
                        }
                        segments.Add(" "); // 保留空格
                    }
                    else if (IsCJKCharacter(c) || char.IsPunctuation(c))
                    {
                        // 中文字符和标点符号单独处理
                        if (currentSegment.Length > 0)
                        {
                            segments.Add(currentSegment.ToString());
                            currentSegment.Clear();
                        }
                        segments.Add(c.ToString());
                    }
                    else
                    {
                        currentSegment.Append(c);
                        
                        // 英文单词达到一定长度时分割
                        if (currentSegment.Length >= 10 && char.IsLetter(c))
                        {
                            segments.Add(currentSegment.ToString());
                            currentSegment.Clear();
                        }
                    }
                }
                
                if (currentSegment.Length > 0)
                {
                    segments.Add(currentSegment.ToString());
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"文本分割失败: {ex.Message}");
                // 回退到简单的空格分割
                segments.AddRange(text.Split(' ', StringSplitOptions.None));
            }
            
            return segments;
        }

        /// <summary>
        /// 判断是否为中日韩字符
        /// </summary>
        private bool IsCJKCharacter(char c)
        {
            var code = (int)c;
            return (code >= 0x4E00 && code <= 0x9FFF) ||   // CJK统一汉字
                   (code >= 0x3400 && code <= 0x4DBF) ||   // CJK扩展A
                   (code >= 0x20000 && code <= 0x2A6DF) || // CJK扩展B
                   (code >= 0x3000 && code <= 0x303F) ||   // CJK符号和标点
                   (code >= 0xFF00 && code <= 0xFFEF);     // 全角ASCII
        }

        /// <summary>
        /// 安全测量字符串大小
        /// </summary>
        private XSize SafeMeasureString(XGraphics gfx, string text, XFont font)
        {
            try
            {
                if (string.IsNullOrEmpty(text))
                    return new XSize(0, font.Height);
                
                return gfx.MeasureString(text, font);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"测量字符串失败: {text} - {ex.Message}");
                // 返回估算大小
                return new XSize(text.Length * font.Size * 0.6, font.Height);
            }
        }

        /// <summary>
        /// 安全绘制字符串
        /// </summary>
        private void SafeDrawString(XGraphics gfx, string text, XFont font, XBrush brush, double x, double y)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    gfx.DrawString(text, font, brush, x, y);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"绘制字符串失败: {text} - {ex.Message}");
                // 尝试绘制替代文本
                try
                {
                    gfx.DrawString("[?]", font, brush, x, y);
                }
                catch
                {
                    // 如果连替代文本都失败，就忽略
                }
            }
        }

        /// <summary>
        /// 拆分超长文本
        /// </summary>
        private List<string> BreakLongText(XGraphics gfx, string text, XFont font, double maxWidth)
        {
            var lines = new List<string>();
            
            if (string.IsNullOrEmpty(text)) return lines;
            
            try
            {
                var currentLine = new StringBuilder();
                
                foreach (var c in text)
                {
                    var testLine = currentLine.ToString() + c;
                    var size = SafeMeasureString(gfx, testLine, font);
                    
                    if (size.Width <= maxWidth)
                    {
                        currentLine.Append(c);
                    }
                    else
                    {
                        if (currentLine.Length > 0)
                        {
                            lines.Add(currentLine.ToString());
                            currentLine.Clear();
                        }
                        currentLine.Append(c);
                    }
                }
                
                if (currentLine.Length > 0)
                {
                    lines.Add(currentLine.ToString());
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"拆分长文本失败: {ex.Message}");
                lines.Add(text);
            }
            
            return lines;
        }

        /// <summary>
        /// 渲染图片
        /// </summary>
        private async Task<double> RenderImageAsync(XGraphics gfx, string imagePath, XRect rect, 
            ConversionSettings settings, CancellationToken cancellationToken)
        {
            try
            {
                // 处理图片
                var processedPath = await _imageProcessor.ProcessImageAsync(imagePath, settings);
                
                if (File.Exists(processedPath))
                {
                    using (var image = XImage.FromFile(processedPath))
                    {
                        // 计算图片显示尺寸，保持宽高比
                        var aspectRatio = (double)image.PixelHeight / image.PixelWidth;
                        var displayWidth = Math.Min(rect.Width, settings.MaxImageWidth);
                        var displayHeight = displayWidth * aspectRatio;

                        // 居中显示图片
                        var x = rect.Left + (rect.Width - displayWidth) / 2;
                        var imageRect = new XRect(x, rect.Top, displayWidth, displayHeight);
                        
                        gfx.DrawImage(image, imageRect);
                        return displayHeight;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"渲染图片失败 {imagePath}: {ex.Message}");
            }

            return 0;
        }

        /// <summary>
        /// 常规文件处理
        /// </summary>
        private async Task ProcessRegularFileAsync(PdfDocument document, string htmlContent, 
            ConversionSettings settings, CancellationToken cancellationToken)
        {
            OnProgressChanged(new PdfGenerationProgressEventArgs(20, "解析HTML内容..."));
            await ProcessChunkAsync(document, htmlContent, settings, cancellationToken);
        }

        /// <summary>
        /// 触发进度变化事件
        /// </summary>
        protected virtual void OnProgressChanged(PdfGenerationProgressEventArgs e)
        {
            ProgressChanged?.Invoke(this, e);
        }
    }

    /// <summary>
    /// HTML元素
    /// </summary>
    public class HtmlElement
    {
        public string Type { get; set; } = string.Empty;
        public string Content { get; set; } = string.Empty;
        public Dictionary<string, string> Attributes { get; set; } = new();
    }

    /// <summary>
    /// 增强的系统字体解析器
    /// 完整支持中文字体，包括字体文件查找和加载
    /// </summary>
    public class SystemFontResolver : IFontResolver
    {
        private readonly Dictionary<string, string> _fontCache = new();
        private readonly string[] _fontDirectories;

        public SystemFontResolver()
        {
            // Windows系统字体目录
            _fontDirectories = new[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts"),
                @"C:\Windows\Fonts"
            };
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            try
            {
                // 生成字体键
                var style = isBold && isItalic ? "BoldItalic" : 
                           isBold ? "Bold" : 
                           isItalic ? "Italic" : "Regular";
                
                var fontKey = $"{familyName}#{style}";

                // 映射中文字体到实际字体文件
                var actualFontName = familyName switch
                {
                    "Microsoft YaHei" or "微软雅黑" => GetChineseFontName("Microsoft YaHei", isBold),
                    "SimSun" or "宋体" => GetChineseFontName("SimSun", isBold),
                    "SimHei" or "黑体" => GetChineseFontName("SimHei", isBold),
                    "Arial" => GetEnglishFontName("Arial", isBold, isItalic),
                    _ => GetChineseFontName("Microsoft YaHei", isBold) // 默认使用微软雅黑
                };

                return new FontResolverInfo(actualFontName, isBold, isItalic);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"字体解析失败 {familyName}: {ex.Message}");
                // 返回安全的默认字体
                return new FontResolverInfo("Arial", false, false);
            }
        }

        public byte[] GetFont(string faceName)
        {
            try
            {
                // 检查缓存
                if (_fontCache.TryGetValue(faceName, out var cachedPath))
                {
                    if (File.Exists(cachedPath))
                    {
                        return File.ReadAllBytes(cachedPath);
                    }
                }

                // 查找字体文件
                var fontPath = FindFontFile(faceName);
                if (!string.IsNullOrEmpty(fontPath) && File.Exists(fontPath))
                {
                    _fontCache[faceName] = fontPath;
                    return File.ReadAllBytes(fontPath);
                }

                // 如果找不到指定字体，尝试使用系统默认字体
                var fallbackPath = GetFallbackFont();
                if (!string.IsNullOrEmpty(fallbackPath))
                {
                    return File.ReadAllBytes(fallbackPath);
                }

                throw new FileNotFoundException($"无法找到字体文件: {faceName}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"加载字体文件失败 {faceName}: {ex.Message}");
                
                // 最后的回退：使用系统内置字体
                try
                {
                    var arialPath = Path.Combine(_fontDirectories[0], "arial.ttf");
                    if (File.Exists(arialPath))
                    {
                        return File.ReadAllBytes(arialPath);
                    }
                }
                catch { }

                // 如果所有方法都失败，返回空数组（这可能导致异常，但至少不会崩溃）
                throw new FontNotFoundException($"无法加载任何可用字体，请检查系统字体配置。原始错误: {ex.Message}");
            }
        }

        private string GetChineseFontName(string baseName, bool isBold)
        {
            return baseName switch
            {
                "Microsoft YaHei" => isBold ? "Microsoft YaHei Bold" : "Microsoft YaHei",
                "SimSun" => isBold ? "SimSun Bold" : "SimSun",
                "SimHei" => "SimHei", // 黑体通常没有Bold变体
                _ => "Microsoft YaHei"
            };
        }

        private string GetEnglishFontName(string baseName, bool isBold, bool isItalic)
        {
            return baseName switch
            {
                "Arial" when isBold && isItalic => "Arial Bold Italic",
                "Arial" when isBold => "Arial Bold",
                "Arial" when isItalic => "Arial Italic",
                "Arial" => "Arial",
                _ => "Arial"
            };
        }

        private string? FindFontFile(string fontName)
        {
            var possibleNames = new[]
            {
                fontName.ToLowerInvariant(),
                fontName.Replace(" ", "").ToLowerInvariant(),
                fontName.Replace(" ", "_").ToLowerInvariant(),
                fontName.Replace(" ", "-").ToLowerInvariant()
            };

            var extensions = new[] { ".ttf", ".otf", ".ttc" };

            foreach (var directory in _fontDirectories.Where(Directory.Exists))
            {
                foreach (var name in possibleNames)
                {
                    foreach (var ext in extensions)
                    {
                        var path = Path.Combine(directory, name + ext);
                        if (File.Exists(path))
                        {
                            return path;
                        }
                    }
                }

                // 特殊处理常见的中文字体文件名
                var chineseFontMappings = new Dictionary<string, string[]>
                {
                    { "microsoft yahei", new[] { "msyh.ttc", "msyhbd.ttc", "msyhl.ttc" } },
                    { "simsun", new[] { "simsun.ttc", "simsunb.ttf" } },
                    { "simhei", new[] { "simhei.ttf" } },
                    { "arial", new[] { "arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf" } }
                };

                foreach (var mapping in chineseFontMappings)
                {
                    if (possibleNames.Any(n => n.Contains(mapping.Key.Replace(" ", ""))))
                    {
                        foreach (var fileName in mapping.Value)
                        {
                            var path = Path.Combine(directory, fileName);
                            if (File.Exists(path))
                            {
                                return path;
                            }
                        }
                    }
                }
            }

            return null;
        }

        private string? GetFallbackFont()
        {
            // 按优先级尝试回退字体
            var fallbackFonts = new[]
            {
                "msyh.ttc",      // 微软雅黑
                "simsun.ttc",    // 宋体
                "arial.ttf",     // Arial
                "calibri.ttf",   // Calibri
                "tahoma.ttf"     // Tahoma
            };

            foreach (var directory in _fontDirectories.Where(Directory.Exists))
            {
                foreach (var fontFile in fallbackFonts)
                {
                    var path = Path.Combine(directory, fontFile);
                    if (File.Exists(path))
                    {
                        return path;
                    }
                }
            }

            return null;
        }
    }

    /// <summary>
    /// 字体未找到异常
    /// </summary>
    public class FontNotFoundException : Exception
    {
        public FontNotFoundException(string message) : base(message) { }
        public FontNotFoundException(string message, Exception innerException) : base(message, innerException) { }
    }

    /// <summary>
    /// PDF生成进度事件参数
    /// </summary>
    public class PdfGenerationProgressEventArgs : EventArgs
    {
        public int Progress { get; }
        public string Message { get; }

        public PdfGenerationProgressEventArgs(int progress, string message)
        {
            Progress = Math.Max(0, Math.Min(100, progress));
            Message = message;
        }
    }
} 
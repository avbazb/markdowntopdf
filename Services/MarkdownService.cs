using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Markdig;
using Markdig.Extensions.MediaLinks;
using System.Net.Http;
using System.Collections.Generic;
using System.Linq;
using MarkdownToPdf.Models;

namespace MarkdownToPdf.Services
{
    /// <summary>
    /// Markdown解析服务
    /// 负责解析Markdown文档并转换为HTML格式
    /// 支持图片处理、表格、代码高亮等扩展功能
    /// </summary>
    public class MarkdownService
    {
        private readonly MarkdownPipeline _pipeline;
        private readonly HttpClient _httpClient;
        private readonly ImageProcessor _imageProcessor;

        /// <summary>
        /// 转换进度事件
        /// </summary>
        public event EventHandler<MarkdownProcessingProgressEventArgs>? ProgressChanged;

        public MarkdownService()
        {
            // 配置Markdown解析管道，支持各种扩展
            _pipeline = new MarkdownPipelineBuilder()
                .UseAdvancedExtensions() // 启用高级扩展（表格、脚注等）
                .UseMediaLinks() // 支持媒体链接
                .UsePipeTables() // 支持管道表格
                .UseGridTables() // 支持网格表格
                .UseEmphasisExtras() // 支持额外的强调语法
                .UseTaskLists() // 支持任务列表
                .UseMathematics() // 支持数学公式
                .UseAutoIdentifiers() // 自动生成标题ID
                .UseAutoLinks() // 自动链接识别
                .Build();

            _httpClient = new HttpClient();
            _imageProcessor = new ImageProcessor();

            // 订阅图片处理进度事件
            _imageProcessor.ProgressChanged += OnImageProcessingProgress;
        }

        /// <summary>
        /// 处理 Markdown 文本，包括图片处理和优化（高性能版本）
        /// </summary>
        /// <param name="markdownText">Markdown 文本</param>
        /// <param name="settings">转换设置</param>
        /// <param name="basePath">基础路径，用于解析相对路径的图片</param>
        /// <returns>处理后的 HTML 内容</returns>
        public async Task<string> ProcessMarkdownAsync(string markdownText, ConversionSettings settings, string? basePath = null)
        {
            OnProgressChanged(new MarkdownProcessingProgressEventArgs(0, "开始解析 Markdown..."));

            try
            {
                // 第一步：预处理和分析（10%）
                OnProgressChanged(new MarkdownProcessingProgressEventArgs(10, "分析 Markdown 结构..."));
                var imageUrls = ExtractImageUrls(markdownText);
                
                // 第二步：并行处理图片和预处理HTML（20-70%）
                var imageProcessingTask = ProcessImagesInParallel(imageUrls, settings, basePath);
                var htmlPreprocessingTask = PreprocessMarkdownAsync(markdownText);
                
                OnProgressChanged(new MarkdownProcessingProgressEventArgs(20, "开始并行处理..."));
                
                // 等待两个任务完成
                var processedImages = await imageProcessingTask;
                var preprocessedMarkdown = await htmlPreprocessingTask;
                
                OnProgressChanged(new MarkdownProcessingProgressEventArgs(70, "合并处理结果..."));
                
                // 第三步：更新图片路径（80%）
                OnProgressChanged(new MarkdownProcessingProgressEventArgs(80, "更新图片路径..."));
                var updatedMarkdown = UpdateImagePaths(preprocessedMarkdown, processedImages);

                // 第四步：转换为HTML（90%）
                OnProgressChanged(new MarkdownProcessingProgressEventArgs(90, "转换为 HTML..."));
                var html = await ConvertToHtmlAsync(updatedMarkdown);

                // 第五步：后处理优化（100%）
                OnProgressChanged(new MarkdownProcessingProgressEventArgs(100, "优化 HTML 结构..."));
                var optimizedHtml = await OptimizeHtmlAsync(html, settings);

                return optimizedHtml;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"处理 Markdown 时发生错误: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 并行处理图片
        /// </summary>
        private async Task<Dictionary<string, string>> ProcessImagesInParallel(List<string> imageUrls, ConversionSettings settings, string? basePath)
        {
            var processedImages = new Dictionary<string, string>();

            if (imageUrls.Count == 0)
            {
                return processedImages;
            }

            // 解析图片路径
            var imagePaths = new List<string>();
            foreach (var imageUrl in imageUrls)
            {
                var resolvedPath = ResolveImagePath(imageUrl, basePath);
                if (!string.IsNullOrEmpty(resolvedPath) && File.Exists(resolvedPath))
                {
                    imagePaths.Add(resolvedPath);
                }
            }

            // 并行处理图片
            if (imagePaths.Count > 0)
            {
                var imageResults = await _imageProcessor.ProcessImagesAsync(imagePaths, settings);
                
                // 构建URL到处理后路径的映射
                foreach (var imageUrl in imageUrls)
                {
                    var resolvedPath = ResolveImagePath(imageUrl, basePath);
                    if (!string.IsNullOrEmpty(resolvedPath) && imageResults.ContainsKey(resolvedPath))
                    {
                        processedImages[imageUrl] = imageResults[resolvedPath];
                    }
                }
            }

            return processedImages;
        }

        /// <summary>
        /// 预处理Markdown文本（异步进行语法分析等）
        /// </summary>
        private async Task<string> PreprocessMarkdownAsync(string markdownText)
        {
            return await Task.Run(() =>
            {
                // 在这里可以进行一些预处理工作
                // 比如语法高亮预处理、链接验证等
                
                // 移除多余的空行
                var lines = markdownText.Split('\n');
                var processedLines = new List<string>();
                var previousLineEmpty = false;

                foreach (var line in lines)
                {
                    var isCurrentLineEmpty = string.IsNullOrWhiteSpace(line);
                    
                    if (!isCurrentLineEmpty || !previousLineEmpty)
                    {
                        processedLines.Add(line);
                    }
                    
                    previousLineEmpty = isCurrentLineEmpty;
                }

                return string.Join('\n', processedLines);
            });
        }

        /// <summary>
        /// 异步转换为HTML
        /// </summary>
        private async Task<string> ConvertToHtmlAsync(string markdownText)
        {
            return await Task.Run(() =>
            {
                return Markdown.ToHtml(markdownText, _pipeline);
            });
        }

        /// <summary>
        /// 异步优化HTML
        /// </summary>
        private async Task<string> OptimizeHtmlAsync(string html, ConversionSettings settings)
        {
            return await Task.Run(() =>
            {
                return OptimizeHtml(html, settings);
            });
        }

        /// <summary>
        /// 处理 Markdown 文本，包括图片处理和优化
        /// </summary>
        /// <param name="markdownText">Markdown 文本</param>
        /// <param name="settings">转换设置</param>
        /// <param name="basePath">基础路径，用于解析相对路径的图片</param>
        /// <returns>处理后的 HTML 内容</returns>
        public async Task<string> ParseMarkdownAsync(string markdownText, string? basePath = null)
        {
            try
            {
                // 报告进度
                OnProgressChanged(new MarkdownProcessingProgressEventArgs(0, "开始解析Markdown..."));

                // 预处理：处理图片路径
                var processedMarkdown = await ProcessImagePathsAsync(markdownText, basePath);
                OnProgressChanged(new MarkdownProcessingProgressEventArgs(30, "处理图片路径完成"));

                // 转换为HTML
                var html = Markdown.ToHtml(processedMarkdown, _pipeline);
                OnProgressChanged(new MarkdownProcessingProgressEventArgs(70, "Markdown转HTML完成"));

                // 后处理：添加样式和优化
                var styledHtml = ApplyDefaultStyling(html);
                OnProgressChanged(new MarkdownProcessingProgressEventArgs(100, "解析完成"));

                return styledHtml;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"解析Markdown时发生错误: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 处理Markdown中的图片路径
        /// 支持本地图片和网络图片
        /// </summary>
        private async Task<string> ProcessImagePathsAsync(string markdown, string? basePath)
        {
            // 匹配Markdown图片语法: ![alt](path "title")
            var imageRegex = new Regex(@"!\[([^\]]*)\]\(([^)]+)(?:\s+""([^""]*)"")?\)", RegexOptions.Multiline);
            var matches = imageRegex.Matches(markdown).Cast<Match>().ToList();
            
            var processedMarkdown = markdown;
            var processedCount = 0;

            foreach (var match in matches)
            {
                var altText = match.Groups[1].Value;
                var imagePath = match.Groups[2].Value.Trim();
                var titleText = match.Groups[3].Value;

                try
                {
                    var processedPath = await ProcessSingleImageAsync(imagePath, basePath);
                    
                    // 构建新的图片标记
                    var newImageTag = string.IsNullOrEmpty(titleText) 
                        ? $"![{altText}]({processedPath})"
                        : $"![{altText}]({processedPath} \"{titleText}\")";
                    
                    processedMarkdown = processedMarkdown.Replace(match.Value, newImageTag);
                }
                catch (Exception ex)
                {
                    // 图片处理失败时保留原始标记并记录警告
                    System.Diagnostics.Debug.WriteLine($"处理图片失败 {imagePath}: {ex.Message}");
                }

                processedCount++;
                // 报告图片处理进度
                if (matches.Count > 0)
                {
                    var progress = 30 + (processedCount * 40 / matches.Count);
                    OnProgressChanged(new MarkdownProcessingProgressEventArgs(progress, $"处理图片 {processedCount}/{matches.Count}"));
                }
            }

            return processedMarkdown;
        }

        /// <summary>
        /// 处理单个图片
        /// </summary>
        private async Task<string> ProcessSingleImageAsync(string imagePath, string? basePath)
        {
            // 处理网络图片
            if (imagePath.StartsWith("http://") || imagePath.StartsWith("https://"))
            {
                return await ProcessNetworkImageAsync(imagePath);
            }

            // 处理本地图片
            return ProcessLocalImage(imagePath, basePath);
        }

        /// <summary>
        /// 处理网络图片（下载并缓存）
        /// </summary>
        private async Task<string> ProcessNetworkImageAsync(string imageUrl)
        {
            try
            {
                // 创建缓存目录
                var cacheDir = Path.Combine(Path.GetTempPath(), "MarkdownToPdf", "ImageCache");
                Directory.CreateDirectory(cacheDir);

                // 生成缓存文件名
                var fileName = Path.GetFileName(new Uri(imageUrl).LocalPath);
                if (string.IsNullOrEmpty(fileName) || !Path.HasExtension(fileName))
                {
                    fileName = $"image_{imageUrl.GetHashCode():X8}.jpg";
                }

                var cachedPath = Path.Combine(cacheDir, fileName);

                // 如果缓存文件不存在，则下载
                if (!File.Exists(cachedPath))
                {
                    var imageData = await _httpClient.GetByteArrayAsync(imageUrl);
                    await File.WriteAllBytesAsync(cachedPath, imageData);
                }

                return cachedPath;
            }
            catch (Exception ex)
            {
                // 网络图片下载失败时返回原始URL
                System.Diagnostics.Debug.WriteLine($"下载网络图片失败 {imageUrl}: {ex.Message}");
                return imageUrl;
            }
        }

        /// <summary>
        /// 处理本地图片
        /// </summary>
        private string ProcessLocalImage(string imagePath, string? basePath)
        {
            // 如果是绝对路径且文件存在，直接返回
            if (Path.IsPathRooted(imagePath) && File.Exists(imagePath))
            {
                return imagePath;
            }

            // 处理相对路径
            if (!string.IsNullOrEmpty(basePath))
            {
                var fullPath = Path.Combine(basePath, imagePath);
                if (File.Exists(fullPath))
                {
                    return Path.GetFullPath(fullPath);
                }
            }

            // 如果找不到文件，返回原始路径
            return imagePath;
        }

        /// <summary>
        /// 为HTML应用默认样式
        /// 确保良好的PDF渲染效果和中文字体支持
        /// </summary>
        private string ApplyDefaultStyling(string html)
        {
            var styledHtml = $@"
<!DOCTYPE html>
<html lang=""zh-CN"">
<head>
    <meta charset=""UTF-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
    <style>
        body {{
            font-family: 'Microsoft YaHei', 'PingFang SC', 'Hiragino Sans GB', 'SimSun', sans-serif;
            font-size: 14px;
            line-height: 1.6;
            color: #333;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            background: white;
        }}
        
        h1, h2, h3, h4, h5, h6 {{
            color: #2c3e50;
            margin-top: 24px;
            margin-bottom: 16px;
            font-weight: 600;
            line-height: 1.25;
        }}
        
        h1 {{ font-size: 2em; border-bottom: 2px solid #eee; padding-bottom: 10px; }}
        h2 {{ font-size: 1.5em; border-bottom: 1px solid #eee; padding-bottom: 8px; }}
        h3 {{ font-size: 1.25em; }}
        
        p {{ margin-bottom: 16px; }}
        
        code {{
            background-color: #f6f8fa;
            padding: 2px 4px;
            border-radius: 3px;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            font-size: 85%;
        }}
        
        pre {{
            background-color: #f6f8fa;
            padding: 16px;
            border-radius: 6px;
            overflow: auto;
            margin-bottom: 16px;
        }}
        
        pre code {{
            background: none;
            padding: 0;
        }}
        
        blockquote {{
            border-left: 4px solid #dfe2e5;
            padding-left: 16px;
            margin-left: 0;
            color: #6a737d;
        }}
        
        table {{
            border-collapse: collapse;
            width: 100%;
            margin-bottom: 16px;
        }}
        
        th, td {{
            border: 1px solid #dfe2e5;
            padding: 6px 13px;
            text-align: left;
        }}
        
        th {{
            background-color: #f6f8fa;
            font-weight: 600;
        }}
        
        img {{
            max-width: 100%;
            height: auto;
            display: block;
            margin: 16px auto;
            border-radius: 4px;
        }}
        
        ul, ol {{
            padding-left: 2em;
            margin-bottom: 16px;
        }}
        
        li {{
            margin-bottom: 4px;
        }}
        
        a {{
            color: #0366d6;
            text-decoration: none;
        }}
        
        a:hover {{
            text-decoration: underline;
        }}
        
        .task-list-item {{
            list-style-type: none;
        }}
        
        .task-list-item-checkbox {{
            margin-right: 4px;
        }}
    </style>
</head>
<body>
{html}
</body>
</html>";

            return styledHtml;
        }

        /// <summary>
        /// 图片处理进度事件处理
        /// </summary>
        private void OnImageProcessingProgress(object? sender, ImageProcessingProgressEventArgs e)
        {
            // 将图片处理进度映射到 Markdown 处理进度（40-80%）
            var markdownProgress = 40 + (int)(e.Progress * 0.4);
            OnProgressChanged(new MarkdownProcessingProgressEventArgs(markdownProgress, e.Message));
        }

        /// <summary>
        /// 触发进度变化事件
        /// </summary>
        protected virtual void OnProgressChanged(MarkdownProcessingProgressEventArgs e)
        {
            ProgressChanged?.Invoke(this, e);
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose()
        {
            _httpClient?.Dispose();
            _imageProcessor?.Dispose();
        }

        /// <summary>
        /// 从 Markdown 文本中提取所有图片 URL
        /// </summary>
        /// <param name="markdownText">Markdown 文本</param>
        /// <returns>图片 URL 列表</returns>
        private List<string> ExtractImageUrls(string markdownText)
        {
            var imageUrls = new List<string>();

            // 匹配 Markdown 图片语法: ![alt](url) 或 ![alt](url "title")
            var markdownImageRegex = new Regex(@"!\[.*?\]\(([^)]+)\)", RegexOptions.IgnoreCase);
            var markdownMatches = markdownImageRegex.Matches(markdownText);
            
            foreach (Match match in markdownMatches)
            {
                var url = match.Groups[1].Value.Trim();
                // 移除可能的标题部分
                var spaceIndex = url.IndexOf(' ');
                if (spaceIndex > 0)
                {
                    url = url.Substring(0, spaceIndex);
                }
                if (!imageUrls.Contains(url))
                {
                    imageUrls.Add(url);
                }
            }

            // 匹配 HTML img 标签: <img src="url">
            var htmlImageRegex = new Regex(@"<img[^>]+src\s*=\s*[""']([^""']+)[""'][^>]*>", RegexOptions.IgnoreCase);
            var htmlMatches = htmlImageRegex.Matches(markdownText);
            
            foreach (Match match in htmlMatches)
            {
                var url = match.Groups[1].Value.Trim();
                if (!imageUrls.Contains(url))
                {
                    imageUrls.Add(url);
                }
            }

            return imageUrls;
        }

        /// <summary>
        /// 解析图片路径（支持相对路径和绝对路径）
        /// </summary>
        /// <param name="imageUrl">图片 URL</param>
        /// <param name="basePath">基础路径</param>
        /// <returns>解析后的完整路径</returns>
        private string? ResolveImagePath(string imageUrl, string? basePath)
        {
            try
            {
                // 跳过网络图片
                if (imageUrl.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                    imageUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                    imageUrl.StartsWith("ftp://", StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                // 处理绝对路径
                if (Path.IsPathRooted(imageUrl))
                {
                    return File.Exists(imageUrl) ? imageUrl : null;
                }

                // 处理相对路径
                if (!string.IsNullOrEmpty(basePath))
                {
                    var baseDirectory = Path.GetDirectoryName(basePath);
                    if (!string.IsNullOrEmpty(baseDirectory))
                    {
                        var fullPath = Path.GetFullPath(Path.Combine(baseDirectory, imageUrl));
                        return File.Exists(fullPath) ? fullPath : null;
                    }
                }

                // 尝试当前工作目录
                var workingDirPath = Path.GetFullPath(imageUrl);
                return File.Exists(workingDirPath) ? workingDirPath : null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 更新 Markdown 中的图片路径
        /// </summary>
        /// <param name="markdownText">原始 Markdown 文本</param>
        /// <param name="processedImages">处理后的图片路径映射</param>
        /// <returns>更新后的 Markdown 文本</returns>
        private string UpdateImagePaths(string markdownText, Dictionary<string, string> processedImages)
        {
            if (processedImages.Count == 0)
            {
                return markdownText;
            }

            var result = markdownText;

            foreach (var (originalUrl, processedPath) in processedImages)
            {
                // 转换为 file:// URL，确保路径正确
                var fileUrl = new Uri(processedPath).AbsoluteUri;

                // 替换 Markdown 图片语法
                var markdownPattern = $@"!\[([^\]]*)\]\(\s*{Regex.Escape(originalUrl)}(\s+[^)]*)?\)";
                result = Regex.Replace(result, markdownPattern, $"![$1]({fileUrl})", RegexOptions.IgnoreCase);

                // 替换 HTML img 标签
                var htmlPattern = $@"(<img[^>]+src\s*=\s*[""']){Regex.Escape(originalUrl)}([""'][^>]*>)";
                result = Regex.Replace(result, htmlPattern, $"$1{fileUrl}$2", RegexOptions.IgnoreCase);
            }

            return result;
        }

        /// <summary>
        /// 优化 HTML 内容
        /// </summary>
        /// <param name="html">原始 HTML</param>
        /// <param name="settings">转换设置</param>
        /// <returns>优化后的 HTML</returns>
        private string OptimizeHtml(string html, ConversionSettings settings)
        {
            // 添加响应式图片样式
            html = html.Replace("<img ", "<img style=\"max-width: 100%; height: auto; display: block; margin: 10px 0;\" ");

            // 添加表格样式
            html = html.Replace("<table>", 
                "<table style=\"border-collapse: collapse; width: 100%; margin: 15px 0; border: 1px solid #ddd;\">");
            
            html = html.Replace("<th>", 
                "<th style=\"border: 1px solid #ddd; padding: 8px; background-color: #f2f2f2; text-align: left;\">");
            
            html = html.Replace("<td>", 
                "<td style=\"border: 1px solid #ddd; padding: 8px;\">");

            // 添加代码块样式
            html = html.Replace("<pre><code>", 
                "<pre style=\"background-color: #f8f8f8; border: 1px solid #ddd; border-radius: 4px; padding: 10px; overflow-x: auto; font-family: 'Consolas', 'Monaco', monospace;\"><code>");

            // 添加行内代码样式
            html = Regex.Replace(html, @"<code>(?!</pre>)", 
                "<code style=\"background-color: #f8f8f8; padding: 2px 4px; border-radius: 3px; font-family: 'Consolas', 'Monaco', monospace;\">");

            // 添加引用样式
            html = html.Replace("<blockquote>", 
                "<blockquote style=\"border-left: 4px solid #ddd; margin: 15px 0; padding: 10px 15px; background-color: #f9f9f9; color: #666;\">");

            // 添加段落间距
            html = html.Replace("<p>", "<p style=\"margin: 10px 0; line-height: 1.6;\">");

            // 添加标题样式
            for (int i = 1; i <= 6; i++)
            {
                var margin = Math.Max(20 - i * 2, 10);
                html = html.Replace($"<h{i}>", 
                    $"<h{i} style=\"margin: {margin}px 0 10px 0; color: #333; font-weight: bold;\">");
            }

            // 添加列表样式
            html = html.Replace("<ul>", "<ul style=\"margin: 10px 0; padding-left: 20px;\">");
            html = html.Replace("<ol>", "<ol style=\"margin: 10px 0; padding-left: 20px;\">");
            html = html.Replace("<li>", "<li style=\"margin: 5px 0; line-height: 1.5;\">");

            return html;
        }
    }

    /// <summary>
    /// Markdown 处理进度事件参数
    /// </summary>
    public class MarkdownProcessingProgressEventArgs : EventArgs
    {
        public int Progress { get; }
        public string Message { get; }

        public MarkdownProcessingProgressEventArgs(int progress, string message)
        {
            Progress = Math.Max(0, Math.Min(100, progress));
            Message = message;
        }
    }
} 
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Markdig.Extensions.Tables;
using MarkdownToPdf.Models;
using HtmlAgilityPack;

namespace MarkdownToPdf.Services
{
    /// <summary>
    /// 增强的直接Markdown到PDF转换服务
    /// 支持多种图片格式、智能图片布局和并排显示
    /// </summary>
    public class DirectMarkdownToPdfService
    {
        private readonly MarkdownPipeline _pipeline;
        private readonly ImageProcessor _imageProcessor;
        private double _currentY;
        private double _pageWidth;
        private double _pageHeight;
        private double _leftMargin;
        private double _rightMargin;
        private double _topMargin;
        private double _bottomMargin;
        private string? _basePath;
        
        // 页码和目录相关字段
        private int _currentPageNumber;
        private readonly List<TocItem> _tocItems = new();
        private readonly Dictionary<int, int> _blockToPageMapping = new();
        private readonly List<TocItem> _actualTocItems = new(); // 实际渲染时收集的目录项

        /// <summary>
        /// 转换进度事件
        /// </summary>
        public event EventHandler<DirectPdfGenerationProgressEventArgs>? ProgressChanged;

        public DirectMarkdownToPdfService()
        {
            // 配置Markdown解析管道
            _pipeline = new MarkdownPipelineBuilder()
                .UseAdvancedExtensions()
                .UsePipeTables()
                .UseGridTables() 
                .UseGenericAttributes() // 支持HTML属性
                .UseEmphasisExtras()
                .UseTaskLists()
                .UseAutoIdentifiers()
                .UseAutoLinks()
                .UseSoftlineBreakAsHardlineBreak() // 更好的换行支持
                .Build();

            // 初始化增强的图片处理器
            _imageProcessor = new ImageProcessor();
        }

        /// <summary>
        /// 直接将Markdown转换为PDF
        /// </summary>
        public async Task ConvertMarkdownToPdfAsync(string markdownFilePath, string outputPath, 
            ConversionSettings settings, CancellationToken cancellationToken = default)
        {
            try
            {
                OnProgressChanged(new DirectPdfGenerationProgressEventArgs(0, "开始读取Markdown文件..."));

                // 读取Markdown文件
                var markdownText = await File.ReadAllTextAsync(markdownFilePath, Encoding.UTF8, cancellationToken);
                _basePath = Path.GetDirectoryName(markdownFilePath);

                OnProgressChanged(new DirectPdfGenerationProgressEventArgs(10, "解析Markdown文档..."));

                // 解析Markdown为语法树
                var document = Markdown.Parse(markdownText, _pipeline);

                OnProgressChanged(new DirectPdfGenerationProgressEventArgs(15, "预处理图片..."));

                // 预处理所有图片，实现智能分组和并排显示
                var imageGroups = await PreprocessImagesAsync(document, settings, cancellationToken);

                OnProgressChanged(new DirectPdfGenerationProgressEventArgs(25, "创建PDF文档..."));

                // 创建PDF文档
                using (var pdfDocument = new PdfDocument())
                {
                    pdfDocument.Info.Title = Path.GetFileNameWithoutExtension(markdownFilePath);
                    pdfDocument.Info.Creator = "Enhanced Markdown To PDF Converter";
                    pdfDocument.Info.CreationDate = DateTime.Now;

                    // 设置页面参数
                    SetupPageParameters(settings);

                    OnProgressChanged(new DirectPdfGenerationProgressEventArgs(30, "渲染文档内容..."));

                    // 第一步：先渲染文档内容，收集实际的标题和页码信息
                    await RenderDocumentContentAndCollectToc(pdfDocument, document, settings, imageGroups, cancellationToken);

                    OnProgressChanged(new DirectPdfGenerationProgressEventArgs(70, "生成目录页..."));

                    // 第二步：基于实际页码信息，在第一页插入目录
                    await InsertTableOfContentsPage(pdfDocument, settings);

                    OnProgressChanged(new DirectPdfGenerationProgressEventArgs(85, "更新页码..."));

                    // 第三步：更新所有页面的页码（因为插入了目录页，所有页码需要+1）
                    UpdateAllPageNumbers(pdfDocument, settings);

                    OnProgressChanged(new DirectPdfGenerationProgressEventArgs(90, "保存PDF文件..."));

                    // 确保输出目录存在
                    var outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // 保存PDF
                    pdfDocument.Save(outputPath);

                    OnProgressChanged(new DirectPdfGenerationProgressEventArgs(100, "转换完成"));
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Markdown转PDF失败: {ex.Message}", ex);
            }
            finally
            {
                // 清理资源
                _imageProcessor?.Dispose();
            }
        }

        /// <summary>
        /// 预处理图片，实现智能分组和并排显示
        /// </summary>
        private async Task<Dictionary<string, (string processedPath, bool shouldRender)>> PreprocessImagesAsync(MarkdownDocument document, 
            ConversionSettings settings, CancellationToken cancellationToken)
        {
            var allImages = ExtractAllImages(document);
            var imageGroups = new Dictionary<string, (string processedPath, bool shouldRender)>();

            if (allImages.Count == 0) return imageGroups;

            // 按段落分组图片
            var paragraphImages = GroupImagesByParagraph(document);
            
            foreach (var paragraphGroup in paragraphImages)
            {
                if (paragraphGroup.Value.Count > 1)
                {
                    // 多张图片，创建并排显示
                    var resolvedPaths = paragraphGroup.Value
                        .Select(ResolveImagePath)
                        .Where(path => !string.IsNullOrEmpty(path) && File.Exists(path))
                        .ToList();

                    if (resolvedPaths.Count > 1)
                    {
                        var combinedImagePath = await _imageProcessor.CreateSideBySideImageAsync(
                            resolvedPaths!, settings);
                        
                        // 只有第一张图片需要渲染（显示合并后的图片），其他的跳过
                        for (int i = 0; i < paragraphGroup.Value.Count; i++)
                        {
                            var originalUrl = paragraphGroup.Value[i];
                            var shouldRender = i == 0; // 只渲染第一张
                            imageGroups[originalUrl] = (combinedImagePath, shouldRender);
                        }
                    }
                    else
                    {
                        // 如果只有一张有效图片，正常处理
                        foreach (var imageUrl in paragraphGroup.Value)
                        {
                            var resolvedPath = ResolveImagePath(imageUrl);
                            if (!string.IsNullOrEmpty(resolvedPath))
                            {
                                var processedPath = await _imageProcessor.ProcessImageAsync(resolvedPath, settings);
                                imageGroups[imageUrl] = (processedPath, true);
                            }
                        }
                    }
                }
                else if (paragraphGroup.Value.Count == 1)
                {
                    // 单张图片，正常处理
                    var imageUrl = paragraphGroup.Value[0];
                    var resolvedPath = ResolveImagePath(imageUrl);
                    if (!string.IsNullOrEmpty(resolvedPath))
                    {
                        var processedPath = await _imageProcessor.ProcessImageAsync(resolvedPath, settings);
                        imageGroups[imageUrl] = (processedPath, true);
                    }
                }
            }

            return imageGroups;
        }

        /// <summary>
        /// 提取文档中的所有图片
        /// </summary>
        private List<string> ExtractAllImages(MarkdownDocument document)
        {
            var images = new List<string>();
            
            foreach (var block in document.Descendants())
            {
                if (block is ParagraphBlock paragraph)
                {
                    foreach (var inline in paragraph.Inline?.Descendants() ?? Enumerable.Empty<Inline>())
                    {
                        if (inline is LinkInline { IsImage: true } imageLink && !string.IsNullOrEmpty(imageLink.Url))
                        {
                            images.Add(imageLink.Url);
                        }
                    }
                }
            }

            return images;
        }

        /// <summary>
        /// 智能识别连续图片进行分组
        /// </summary>
        private Dictionary<int, List<string>> GroupImagesByParagraph(MarkdownDocument document)
        {
            var imageGroups = new Dictionary<int, List<string>>();
            var currentGroup = new List<string>();
            var groupIndex = 0;

            foreach (var block in document)
            {
                if (block is ParagraphBlock paragraph && ContainsImage(paragraph.Inline))
                {
                    // 检查段落是否只包含图片（没有其他文字内容）
                    var isImageOnlyParagraph = IsImageOnlyParagraph(paragraph);
                    
                    if (isImageOnlyParagraph)
                    {
                        // 提取图片URL
                        var images = ExtractImagesFromParagraph(paragraph);
                        currentGroup.AddRange(images);
                    }
                    else
                    {
                        // 如果有文字内容，结束当前组并开始新组
                        if (currentGroup.Count > 0)
                        {
                            imageGroups[groupIndex++] = new List<string>(currentGroup);
                            currentGroup.Clear();
                        }
                        
                        // 单独处理含文字的图片段落
                        var images = ExtractImagesFromParagraph(paragraph);
                        if (images.Count > 0)
                        {
                            imageGroups[groupIndex++] = images;
                        }
                    }
                }
                else
                {
                    // 非图片段落，结束当前图片组
                    if (currentGroup.Count > 0)
                    {
                        imageGroups[groupIndex++] = new List<string>(currentGroup);
                        currentGroup.Clear();
                    }
                }
            }

            // 处理最后一组
            if (currentGroup.Count > 0)
            {
                imageGroups[groupIndex] = currentGroup;
            }

            return imageGroups;
        }

        /// <summary>
        /// 检查段落是否只包含图片（没有文字内容）
        /// </summary>
        private bool IsImageOnlyParagraph(ParagraphBlock paragraph)
        {
            if (paragraph.Inline == null) return false;

            foreach (var inline in paragraph.Inline)
            {
                switch (inline)
                {
                    case LinkInline { IsImage: true }:
                        // 图片链接，继续检查
                        continue;
                    case LiteralInline literal:
                        // 检查是否只包含空白字符
                        if (!string.IsNullOrWhiteSpace(literal.Content.ToString()))
                            return false;
                        break;
                    case LineBreakInline:
                        // 换行符，忽略
                        continue;
                    default:
                        // 其他内联元素表示有文字内容
                        return false;
                }
            }

            return true;
        }

        /// <summary>
        /// 从段落中提取图片URL
        /// </summary>
        private List<string> ExtractImagesFromParagraph(ParagraphBlock paragraph)
        {
            var images = new List<string>();
            
            foreach (var inline in paragraph.Inline?.Descendants() ?? Enumerable.Empty<Inline>())
            {
                if (inline is LinkInline { IsImage: true } imageLink && !string.IsNullOrEmpty(imageLink.Url))
                {
                    images.Add(imageLink.Url);
                }
            }

            return images;
        }

        /// <summary>
        /// 渲染文档内容并收集实际的目录信息
        /// </summary>
        private async Task RenderDocumentContentAndCollectToc(PdfDocument pdfDocument, MarkdownDocument document, 
            ConversionSettings settings, Dictionary<string, (string processedPath, bool shouldRender)> imageGroups, CancellationToken cancellationToken)
        {
            _actualTocItems.Clear();
            var totalBlocks = CountBlocks(document);
            var processedBlocks = 0;
            
            // 创建第一页（后续将成为第二页，第一页留给目录）
            var page = pdfDocument.AddPage();
            page.Width = _pageWidth;
            page.Height = _pageHeight;
            var gfx = XGraphics.FromPdfPage(page);
            _currentY = _topMargin;
            _currentPageNumber = 1; // 内容页从1开始，后续插入目录后变成2

            try
            {
                foreach (var block in document)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // 检查是否需要新页面
                    var estimatedHeight = await EstimateBlockHeightAsync(block, settings, imageGroups);
                    if (_currentY + estimatedHeight > _pageHeight - _bottomMargin - 30) // 预留页码空间
                    {
                        // 创建新页面
                        gfx.Dispose();
                        page = pdfDocument.AddPage();
                        page.Width = _pageWidth;
                        page.Height = _pageHeight;
                        gfx = XGraphics.FromPdfPage(page);
                        _currentY = _topMargin;
                        _currentPageNumber++;
                    }

                    // 收集标题信息（在渲染前记录当前页码）
                    if (block is HeadingBlock heading && (heading.Level == 1 || heading.Level == 2))
                    {
                        var title = ExtractTextFromInlines(heading.Inline);
                        _actualTocItems.Add(new TocItem
                        {
                            Level = heading.Level,
                            Title = title,
                            PageNumber = _currentPageNumber + 1 // +1因为后续会插入目录页
                        });
                    }

                    // 渲染块
                    await RenderBlockAsync(gfx, block, settings, imageGroups, cancellationToken);

                    processedBlocks++;
                    var progress = 30 + (processedBlocks * 40 / totalBlocks);
                    OnProgressChanged(new DirectPdfGenerationProgressEventArgs(progress, 
                        $"渲染内容... ({processedBlocks}/{totalBlocks})"));
                }
            }
            finally
            {
                gfx.Dispose();
            }
        }

        /// <summary>
        /// 在第一页插入目录页
        /// </summary>
        private async Task InsertTableOfContentsPage(PdfDocument pdfDocument, ConversionSettings settings)
        {
            // 在第一页位置插入目录页
            var tocPage = pdfDocument.InsertPage(0);
            tocPage.Width = _pageWidth;
            tocPage.Height = _pageHeight;

            var gfx = XGraphics.FromPdfPage(tocPage);
            
            try
            {
                var fontSize = settings.FontSize * 0.75;
                var titleFont = CreateSafeFont(settings.FontFamily, fontSize + 6, XFontStyleEx.Bold);
                var h1Font = CreateSafeFont(settings.FontFamily, fontSize + 2, XFontStyleEx.Bold);
                var h2Font = CreateSafeFont(settings.FontFamily, fontSize, XFontStyleEx.Regular);

                var currentY = _topMargin;

                // 标题
                SafeDrawString(gfx, "目录", titleFont, XBrushes.Black, _leftMargin, currentY);
                currentY += (fontSize + 6) * 1.5 + 20;

                // 渲染目录项（支持自动换行）
                foreach (var item in _actualTocItems)
                {
                    var font = item.Level == 1 ? h1Font : h2Font;
                    var indent = item.Level == 1 ? 0 : 20;
                    
                    // 计算可用宽度（预留页码空间）
                    var pageText = item.PageNumber.ToString();
                    var pageSize = SafeMeasureString(gfx, pageText, font);
                    var pageNumberWidth = pageSize.Width + 20; // 页码宽度 + 边距
                    var availableWidth = _pageWidth - _leftMargin - _rightMargin - indent - pageNumberWidth;

                    // 渲染标题（支持自动换行）
                    var titleX = _leftMargin + indent;
                    var titleRect = new XRect(titleX, currentY, availableWidth, font.Height * 3);
                    var titleHeight = RenderTextWithWrapping(gfx, item.Title, font, XBrushes.Black, titleRect, settings);

                    // 绘制页码（对齐到标题的最后一行）
                    var pageX = _pageWidth - _rightMargin - pageSize.Width;
                    var pageY = currentY + titleHeight - font.Height; // 对齐最后一行
                    SafeDrawString(gfx, pageText, font, XBrushes.Black, pageX, pageY);

                    // 绘制点线（连接标题末尾和页码）
                    var lastLine = GetLastLine(item.Title, font, availableWidth, gfx);
                    var titleLastLineWidth = SafeMeasureString(gfx, lastLine, font).Width;
                    var dotStartX = titleX + titleLastLineWidth + 8;
                    var dotEndX = pageX - 8;
                    var dotY = pageY + font.Height * 0.4; // 调整垂直位置到行中央

                    if (dotEndX > dotStartX + 20) // 确保有足够空间绘制点线
                    {
                        var pen = new XPen(XColors.Gray, 0.8) { DashStyle = XDashStyle.Dot };
                        gfx.DrawLine(pen, dotStartX, dotY, dotEndX, dotY);
                    }

                    currentY += titleHeight + font.Height * 0.3; // 项目间距

                    // 检查是否需要分页（目录页也可能需要分页）
                    if (currentY > _pageHeight - _bottomMargin - 50)
                    {
                        // 创建新的目录页
                        gfx.Dispose();
                        
                        // 查找当前目录页的索引
                        var currentTocPageIndex = -1;
                        for (int i = 0; i < pdfDocument.PageCount; i++)
                        {
                            if (pdfDocument.Pages[i] == tocPage)
                            {
                                currentTocPageIndex = i;
                                break;
                            }
                        }
                        
                        var newTocPage = pdfDocument.InsertPage(currentTocPageIndex + 1);
                        newTocPage.Width = _pageWidth;
                        newTocPage.Height = _pageHeight;
                        gfx = XGraphics.FromPdfPage(newTocPage);
                        tocPage = newTocPage;
                        currentY = _topMargin;
                        
                        // 在新页面也渲染"目录"标题
                        SafeDrawString(gfx, "目录 (续)", titleFont, XBrushes.Black, _leftMargin, currentY);
                        currentY += (fontSize + 6) * 1.5 + 20;
                    }
                }

                // 渲染目录页页码
                RenderPageNumber(gfx, 1, settings);
            }
            finally
            {
                gfx.Dispose();
            }
        }

        /// <summary>
        /// 获取文本换行后的最后一行内容
        /// </summary>
        private string GetLastLine(string text, XFont font, double maxWidth, XGraphics gfx)
        {
            if (string.IsNullOrEmpty(text)) return "";

            // 使用与RenderTextWithWrapping相同的逻辑来计算换行
            var lines = WrapTextToLines(text, font, maxWidth, gfx);
            
            return lines.Count > 0 ? lines.Last() : text;
        }

        /// <summary>
        /// 更新所有页面的页码（因为插入了目录页）
        /// </summary>
        private void UpdateAllPageNumbers(PdfDocument pdfDocument, ConversionSettings settings)
        {
            // 计算目录页数量
            var tocPageCount = 0;
            for (int i = 0; i < pdfDocument.PageCount; i++)
            {
                // 通过检查页面内容来确定是否为目录页
                // 简化判断：前面连续的页面都是目录页，直到遇到内容页
                if (i == 0 || IsTocPage(pdfDocument.Pages[i]))
                {
                    tocPageCount++;
                }
                else
                {
                    break;
                }
            }

            // 为目录页设置页码
            for (int i = 0; i < tocPageCount; i++)
            {
                var page = pdfDocument.Pages[i];
                var gfx = XGraphics.FromPdfPage(page);
                
                try
                {
                    RenderPageNumber(gfx, i + 1, settings);
                }
                finally
                {
                    gfx.Dispose();
                }
            }

            // 为内容页更新页码（从目录页数量+1开始）
            for (int i = tocPageCount; i < pdfDocument.PageCount; i++)
            {
                var page = pdfDocument.Pages[i];
                var gfx = XGraphics.FromPdfPage(page);
                
                try
                {
                    // 清除旧页码区域（用白色矩形覆盖）
                    var clearRect = new XRect(_pageWidth - _rightMargin - 50, _pageHeight - _bottomMargin, 50, _bottomMargin);
                    gfx.DrawRectangle(XBrushes.White, clearRect);
                    
                    // 渲染新页码（内容页从目录页数量+1开始）
                    RenderPageNumber(gfx, i + 1, settings);
                }
                finally
                {
                    gfx.Dispose();
                }
            }
        }

        /// <summary>
        /// 简单判断是否为目录页（基于页面位置）
        /// </summary>
        private bool IsTocPage(PdfPage page)
        {
            // 简化判断：前面的页面通常是目录页
            // 这里可以根据需要添加更复杂的判断逻辑
            return true; // 暂时简化处理
        }

        /// <summary>
        /// 设置页面参数
        /// </summary>
        private void SetupPageParameters(ConversionSettings settings)
        {
            // 根据页面大小设置尺寸（点单位）
            switch (settings.PageSize)
            {
                case PageSize.A3:
                    _pageWidth = 842; _pageHeight = 1191;
                    break;
                case PageSize.A5:
                    _pageWidth = 420; _pageHeight = 595;
                    break;
                case PageSize.Letter:
                    _pageWidth = 612; _pageHeight = 792;
                    break;
                case PageSize.Legal:
                    _pageWidth = 612; _pageHeight = 1008;
                    break;
                default: // A4
                    _pageWidth = 595; _pageHeight = 842;
                    break;
            }

            // 设置边距 (转换毫米到点，1mm ≈ 2.83点)
            _leftMargin = settings.MarginLeft * 2.83;
            _rightMargin = settings.MarginRight * 2.83;
            _topMargin = settings.MarginTop * 2.83;
            _bottomMargin = settings.MarginBottom * 2.83;
        }

        /// <summary>
        /// 渲染包含目录的完整文档（旧方法，保留备用）
        /// </summary>
        private async Task RenderDocumentWithTocAsync(PdfDocument pdfDocument, MarkdownDocument document, 
            ConversionSettings settings, Dictionary<string, (string processedPath, bool shouldRender)> imageGroups, CancellationToken cancellationToken)
        {
            // 这个方法已被新的三步流程替代
            // 保留以备兼容性需要
            await Task.CompletedTask;
        }

        /// <summary>
        /// 渲染目录页
        /// </summary>
        private async Task RenderTableOfContentsPageAsync(PdfDocument pdfDocument, ConversionSettings settings)
        {
            var page = pdfDocument.AddPage();
            page.Width = _pageWidth;
            page.Height = _pageHeight;

            var gfx = XGraphics.FromPdfPage(page);
            
            try
            {
                var fontSize = settings.FontSize * 0.75;
                var titleFont = CreateSafeFont(settings.FontFamily, fontSize + 6, XFontStyleEx.Bold);
                var h1Font = CreateSafeFont(settings.FontFamily, fontSize + 2, XFontStyleEx.Bold);
                var h2Font = CreateSafeFont(settings.FontFamily, fontSize, XFontStyleEx.Regular);

                var currentY = _topMargin;

                // 标题
                SafeDrawString(gfx, "目录", titleFont, XBrushes.Black, _leftMargin, currentY);
                currentY += (fontSize + 6) * 1.5 + 20;

                // 渲染目录项
                foreach (var item in _tocItems)
                {
                    var font = item.Level == 1 ? h1Font : h2Font;
                    var indent = item.Level == 1 ? 0 : 20;
                    
                    // 绘制标题
                    var titleX = _leftMargin + indent;
                    SafeDrawString(gfx, item.Title, font, XBrushes.Black, titleX, currentY);

                    // 绘制页码
                    var pageText = item.PageNumber.ToString();
                    var pageSize = SafeMeasureString(gfx, pageText, font);
                    var pageX = _pageWidth - _rightMargin - pageSize.Width;
                    SafeDrawString(gfx, pageText, font, XBrushes.Black, pageX, currentY);

                    // 绘制点线
                    var titleSize = SafeMeasureString(gfx, item.Title, font);
                    var dotStartX = titleX + titleSize.Width + 10;
                    var dotEndX = pageX - 10;
                    var dotY = currentY + pageSize.Height / 2;

                    if (dotEndX > dotStartX)
                    {
                        var pen = new XPen(XColors.Gray, 0.5) { DashStyle = XDashStyle.Dot };
                        gfx.DrawLine(pen, dotStartX, dotY, dotEndX, dotY);
                    }

                    currentY += font.Height * 1.2;
                }

                // 渲染页码
                RenderPageNumber(gfx, 1, settings);
            }
            finally
            {
                gfx.Dispose();
            }
        }

        /// <summary>
        /// 渲染文档内容
        /// </summary>
        private async Task RenderDocumentContentAsync(PdfDocument pdfDocument, MarkdownDocument document, 
            ConversionSettings settings, Dictionary<string, (string processedPath, bool shouldRender)> imageGroups, CancellationToken cancellationToken)
        {
            var totalBlocks = CountBlocks(document);
            var processedBlocks = 0;
            var blockIndex = 0;

            // 创建第一个内容页
            var page = pdfDocument.AddPage();
            page.Width = _pageWidth;
            page.Height = _pageHeight;
            var gfx = XGraphics.FromPdfPage(page);
            _currentY = _topMargin;
            _currentPageNumber = 2; // 内容从第2页开始

            try
            {
                foreach (var block in document)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // 检查是否需要新页面
                    var estimatedHeight = await EstimateBlockHeightAsync(block, settings, imageGroups);
                    if (_currentY + estimatedHeight > _pageHeight - _bottomMargin - 30) // 预留页码空间
                    {
                        // 渲染当前页的页码
                        RenderPageNumber(gfx, _currentPageNumber, settings);
                        
                        // 创建新页面
                        gfx.Dispose();
                        page = pdfDocument.AddPage();
                        page.Width = _pageWidth;
                        page.Height = _pageHeight;
                        gfx = XGraphics.FromPdfPage(page);
                        _currentY = _topMargin;
                        _currentPageNumber++;
                    }

                    // 渲染块
                    await RenderBlockAsync(gfx, block, settings, imageGroups, cancellationToken);

                    processedBlocks++;
                    blockIndex++;
                    var progress = 40 + (processedBlocks * 45 / totalBlocks);
                    OnProgressChanged(new DirectPdfGenerationProgressEventArgs(progress, 
                        $"渲染内容... ({processedBlocks}/{totalBlocks})"));
                }

                // 渲染最后一页的页码
                RenderPageNumber(gfx, _currentPageNumber, settings);
            }
            finally
            {
                gfx.Dispose();
            }
        }

        /// <summary>
        /// 渲染页码（右下角）
        /// </summary>
        private void RenderPageNumber(XGraphics gfx, int pageNumber, ConversionSettings settings)
        {
            var font = CreateSafeFont(settings.FontFamily, settings.FontSize * 0.6, XFontStyleEx.Regular);
            var pageText = pageNumber.ToString();
            var textSize = SafeMeasureString(gfx, pageText, font);
            
            var x = _pageWidth - _rightMargin - textSize.Width;
            var y = _pageHeight - _bottomMargin / 2;
            
            SafeDrawString(gfx, pageText, font, XBrushes.Gray, x, y);
        }

        /// <summary>
        /// 旧的渲染整个文档方法（保留以备用）
        /// </summary>
        private async Task RenderDocumentAsync(PdfDocument pdfDocument, MarkdownDocument document, 
            ConversionSettings settings, Dictionary<string, (string processedPath, bool shouldRender)> imageGroups, CancellationToken cancellationToken)
        {
            var totalBlocks = CountBlocks(document);
            var processedBlocks = 0;

            // 创建第一页
            var page = pdfDocument.AddPage();
            page.Width = _pageWidth;
            page.Height = _pageHeight;

            var gfx = XGraphics.FromPdfPage(page);
            _currentY = _topMargin;

            try
            {
                foreach (var block in document)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // 检查是否需要新页面
                    var estimatedHeight = EstimateBlockHeight(block, settings);
                    if (_currentY + estimatedHeight > _pageHeight - _bottomMargin)
                    {
                        // 需要新页面
                        gfx.Dispose();
                        page = pdfDocument.AddPage();
                        page.Width = _pageWidth;
                        page.Height = _pageHeight;
                        gfx = XGraphics.FromPdfPage(page);
                        _currentY = _topMargin;
                    }

                    // 渲染块
                    await RenderBlockAsync(gfx, block, settings, imageGroups, cancellationToken);

                    processedBlocks++;
                    var progress = 30 + (processedBlocks * 60 / totalBlocks);
                    OnProgressChanged(new DirectPdfGenerationProgressEventArgs(progress, 
                        $"渲染内容... ({processedBlocks}/{totalBlocks})"));
                }
            }
            finally
            {
                gfx.Dispose();
            }
        }

        /// <summary>
        /// 渲染单个块
        /// </summary>
        private async Task RenderBlockAsync(XGraphics gfx, Block block, ConversionSettings settings, 
            Dictionary<string, (string processedPath, bool shouldRender)> imageGroups, CancellationToken cancellationToken)
        {
            switch (block)
            {
                case HeadingBlock heading:
                    RenderHeading(gfx, heading, settings);
                    break;

                case ParagraphBlock paragraph:
                    await RenderParagraphAsync(gfx, paragraph, settings, imageGroups, cancellationToken);
                    break;

                case CodeBlock codeBlock:
                    RenderCodeBlock(gfx, codeBlock, settings);
                    break;

                case QuoteBlock quote:
                    await RenderQuoteBlockAsync(gfx, quote, settings, imageGroups, cancellationToken);
                    break;

                case ListBlock list:
                    await RenderListAsync(gfx, list, settings, imageGroups, cancellationToken);
                    break;

                case Table table:
                    await RenderTableAsync(gfx, table, settings, imageGroups, cancellationToken);
                    break;

                case ThematicBreakBlock:
                    RenderHorizontalRule(gfx, settings);
                    break;

                default:
                    // 检查是否为HTML块
                    if (block.GetType().Name == "HtmlBlock")
                    {
                        await RenderHtmlBlockAsync(gfx, (dynamic)block, settings, imageGroups, cancellationToken);
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"未处理的块类型: {block.GetType().Name}");
                        _currentY += settings.FontSize;
                    }
                    break;
            }

            // 添加块之间的间距
            _currentY += settings.FontSize * 0.5;
        }

        /// <summary>
        /// 渲染标题
        /// </summary>
        private void RenderHeading(XGraphics gfx, HeadingBlock heading, ConversionSettings settings)
        {
            var level = heading.Level;
            // 调整字体大小：让标题更合理，H1最大，逐级递减
            var baseFontSize = settings.FontSize * 0.75; // 缩小基础字体
            var fontSize = level switch
            {
                1 => baseFontSize + 6,
                2 => baseFontSize + 4,
                3 => baseFontSize + 2,
                4 => baseFontSize + 1,
                _ => baseFontSize
            };
            
            var font = CreateSafeFont(settings.FontFamily, fontSize, XFontStyleEx.Bold);

            // 添加标题前的间距
            if (_currentY > _topMargin)
            {
                _currentY += fontSize * 0.6;
            }

            var text = ExtractTextFromInlines(heading.Inline);
            var rect = new XRect(_leftMargin, _currentY, _pageWidth - _leftMargin - _rightMargin, fontSize * 3);
            
            var height = RenderTextWithWrapping(gfx, text, font, XBrushes.Black, rect, settings);
            _currentY += height;

            // 为H1和H2添加下划线
            if (level <= 2)
            {
                var lineY = _currentY + 3;
                var lineWidth = level == 1 ? 1.5 : 1;
                var pen = new XPen(XColors.Gray, lineWidth);
                gfx.DrawLine(pen, _leftMargin, lineY, _pageWidth - _rightMargin, lineY);
                _currentY += 8;
            }

            _currentY += fontSize * 0.2; // 标题后的间距
        }

        /// <summary>
        /// 渲染段落
        /// </summary>
        private async Task RenderParagraphAsync(XGraphics gfx, ParagraphBlock paragraph, ConversionSettings settings, 
            Dictionary<string, (string processedPath, bool shouldRender)> imageGroups, CancellationToken cancellationToken)
        {
            if (paragraph.Inline == null) return;

            // 缩小正文字体大小
            var fontSize = settings.FontSize * 0.75;
            var font = CreateSafeFont(settings.FontFamily, fontSize, XFontStyleEx.Regular);
            
            // 处理图片
            if (ContainsImage(paragraph.Inline))
            {
                await RenderImageParagraphAsync(gfx, paragraph, settings, imageGroups, cancellationToken);
                return;
            }

            // 普通文本段落
            var text = ExtractTextFromInlines(paragraph.Inline);
            var rect = new XRect(_leftMargin, _currentY, _pageWidth - _leftMargin - _rightMargin, fontSize * 15);
            var height = RenderTextWithWrapping(gfx, text, font, XBrushes.Black, rect, settings);
            _currentY += height + fontSize * 0.3; // 段落间距
        }

        /// <summary>
        /// 检查段落是否包含图片
        /// </summary>
        private bool ContainsImage(ContainerInline inlines)
        {
            foreach (var inline in inlines)
            {
                if (inline is LinkInline link && link.IsImage)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// 渲染包含图片的段落
        /// </summary>
        private async Task RenderImageParagraphAsync(XGraphics gfx, ParagraphBlock paragraph, ConversionSettings settings, 
            Dictionary<string, (string processedPath, bool shouldRender)> imageGroups, CancellationToken cancellationToken)
        {
            foreach (var inline in paragraph.Inline!)
            {
                if (inline is LinkInline link && link.IsImage)
                {
                    await RenderImageAsync(gfx, link, settings, imageGroups, cancellationToken);
                }
                else
                {
                    // 渲染非图片文本
                    var text = ExtractTextFromInline(inline);
                    if (!string.IsNullOrEmpty(text))
                    {
                        var fontSize = settings.FontSize * 0.75;
                        var font = CreateSafeFont(settings.FontFamily, fontSize, XFontStyleEx.Regular);
                        var rect = new XRect(_leftMargin, _currentY, _pageWidth - _leftMargin - _rightMargin, fontSize * 3);
                        var height = RenderTextWithWrapping(gfx, text, font, XBrushes.Black, rect, settings);
                        _currentY += height;
                    }
                }
            }
        }

        /// <summary>
        /// 渲染图片（增强版，支持智能压缩和并排显示）
        /// </summary>
        private async Task RenderImageAsync(XGraphics gfx, LinkInline imageLink, ConversionSettings settings, 
            Dictionary<string, (string processedPath, bool shouldRender)> imageGroups, CancellationToken cancellationToken)
        {
            try
            {
                var imageUrl = imageLink.Url ?? "";
                string imagePath;
                bool shouldRender = true;

                // 检查是否有预处理的图片（可能是合并后的并排图片）
                if (imageGroups.TryGetValue(imageUrl, out var imageInfo))
                {
                    imagePath = imageInfo.processedPath;
                    shouldRender = imageInfo.shouldRender;
                }
                else
                {
                    // fallback到原来的逻辑
                    imagePath = ResolveImagePath(imageUrl);
                    if (!string.IsNullOrEmpty(imagePath))
                    {
                        imagePath = await _imageProcessor.ProcessImageAsync(imagePath, settings);
                    }
                }

                // 如果标记为不需要渲染，直接跳过
                if (!shouldRender)
                {
                    return;
                }

                if (string.IsNullOrEmpty(imagePath) || !File.Exists(imagePath))
                {
                    // 图片不存在，渲染占位符
                    RenderImagePlaceholder(gfx, imageLink.Title ?? "图片", settings);
                    return;
                }

                // 加载图片
                try
                {
                    using (var image = XImage.FromFile(imagePath))
                    {
                        // 计算可用空间
                        var maxWidth = _pageWidth - _leftMargin - _rightMargin;
                        var availableHeight = _pageHeight - _currentY - _bottomMargin - 50; // 预留页码空间
                        
                        var imageWidth = image.PixelWidth;
                        var imageHeight = image.PixelHeight;
                        
                        // 计算基础缩放比例（默认缩小50%以节约空间）
                        var baseScale = 0.7;
                        var scaleX = (maxWidth * baseScale) / imageWidth;
                        var scaleY = (availableHeight * baseScale) / imageHeight;
                        var idealScale = Math.Min(Math.Min(scaleX, scaleY), baseScale);
                        
                        var idealWidth = imageWidth * idealScale;
                        var idealHeight = imageHeight * idealScale;
                        
                        // 智能压缩算法：如果理想尺寸装不下，尝试压缩
                        var finalScale = idealScale;
                        var finalWidth = idealWidth;
                        var finalHeight = idealHeight;
                        
                        if (idealHeight > availableHeight)
                        {
                            // 计算需要的压缩比例
                            var requiredScale = availableHeight / imageHeight;
                            
                            // 检查压缩比例是否在可接受范围内（最多压缩50%）
                            var minAllowedScale = Math.Max(baseScale * 0.5, 0.25); // 最小25%
                            
                            if (requiredScale >= minAllowedScale)
                            {
                                // 可以通过压缩在当前页面显示
                                finalScale = requiredScale;
                                finalWidth = imageWidth * finalScale;
                                finalHeight = imageHeight * finalScale;
                                
                                System.Diagnostics.Debug.WriteLine($"智能压缩图片: 从 {idealScale:P0} 压缩到 {finalScale:P0}");
                            }
                            else
                            {
                                // 即使最大压缩也装不下，需要换页
                                // 这里应该触发换页逻辑
                                return; // 暂时跳过，等待换页逻辑完善
                            }
                        }
                        
                        // 确保图片不会超出页面宽度
                        if (finalWidth > maxWidth)
                        {
                            var widthScale = maxWidth / finalWidth;
                            finalWidth *= widthScale;
                            finalHeight *= widthScale;
                            finalScale *= widthScale;
                        }
                        
                        // 居中显示图片
                        var imageX = _leftMargin + (maxWidth - finalWidth) / 2;
                        var imageRect = new XRect(imageX, _currentY, finalWidth, finalHeight);
                        
                        gfx.DrawImage(image, imageRect);
                        _currentY += finalHeight + 10; // 图片后的间距
                        
                        // 渲染图片标题（如果有）
                        if (!string.IsNullOrEmpty(imageLink.Title))
                        {
                            var captionFontSize = settings.FontSize * 0.65;
                            var captionFont = CreateSafeFont(settings.FontFamily, captionFontSize, XFontStyleEx.Italic);
                            var captionRect = new XRect(_leftMargin, _currentY, maxWidth, captionFontSize * 3);
                            var captionHeight = RenderTextWithWrapping(gfx, imageLink.Title, captionFont, XBrushes.Gray, captionRect, settings);
                            _currentY += captionHeight + 5;
                        }
                        
                        // 输出压缩信息（调试用）
                        if (finalScale < idealScale)
                        {
                            System.Diagnostics.Debug.WriteLine($"图片智能压缩: {Path.GetFileName(imagePath)} " +
                                $"原始尺寸={imageWidth}x{imageHeight}, " +
                                $"理想尺寸={idealWidth:F0}x{idealHeight:F0}, " +
                                $"最终尺寸={finalWidth:F0}x{finalHeight:F0}, " +
                                $"压缩比例={finalScale:P1}");
                        }
                    }
                }
                catch (Exception imageEx)
                {
                    System.Diagnostics.Debug.WriteLine($"加载图片失败: {imagePath} - {imageEx.Message}");
                    RenderImagePlaceholder(gfx, $"图片格式不支持: {Path.GetFileName(imageUrl)}", settings);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"渲染图片失败: {imageLink.Url} - {ex.Message}");
                RenderImagePlaceholder(gfx, $"图片加载失败: {Path.GetFileName(imageLink.Url ?? "")}", settings);
            }
        }

        /// <summary>
        /// 解析图片路径
        /// </summary>
        private string? ResolveImagePath(string imagePath)
        {
            if (string.IsNullOrEmpty(imagePath)) return null;

            // 网络图片暂不支持
            if (imagePath.StartsWith("http://") || imagePath.StartsWith("https://"))
                return null;

            // 绝对路径
            if (Path.IsPathRooted(imagePath) && File.Exists(imagePath))
                return imagePath;

            // 相对路径
            if (!string.IsNullOrEmpty(_basePath))
            {
                var fullPath = Path.Combine(_basePath, imagePath);
                if (File.Exists(fullPath))
                    return Path.GetFullPath(fullPath);
            }

            return null;
        }

        /// <summary>
        /// 渲染图片占位符
        /// </summary>
        private void RenderImagePlaceholder(XGraphics gfx, string text, ConversionSettings settings)
        {
            var placeholderHeight = 80;
            var rect = new XRect(_leftMargin, _currentY, _pageWidth - _leftMargin - _rightMargin, placeholderHeight);
            
            // 绘制背景和边框
            gfx.DrawRectangle(XBrushes.LightGray, rect);
            gfx.DrawRectangle(new XPen(XColors.Gray, 0.8), rect);
            
            // 绘制文本
            var fontSize = settings.FontSize * 0.7;
            var font = CreateSafeFont(settings.FontFamily, fontSize, XFontStyleEx.Italic);
            var textRect = new XRect(rect.X + 8, rect.Y + (placeholderHeight - fontSize) / 2, 
                rect.Width - 16, fontSize * 3);
            RenderTextWithWrapping(gfx, text, font, XBrushes.Gray, textRect, settings);
            
            _currentY += placeholderHeight + 8;
        }

        /// <summary>
        /// 渲染代码块
        /// </summary>
        private void RenderCodeBlock(XGraphics gfx, CodeBlock codeBlock, ConversionSettings settings)
        {
            var fontSize = settings.FontSize * 0.65; // 代码字体更小一些
            var font = CreateSafeFont("Consolas", fontSize, XFontStyleEx.Regular);
            var code = codeBlock.Lines.ToString();
            
            // 绘制背景
            var lines = code.Split('\n');
            var lineHeight = fontSize * settings.LineHeight;
            var totalHeight = lines.Length * lineHeight + 16; // 8px上下padding
            
            var bgRect = new XRect(_leftMargin - 8, _currentY - 4, 
                _pageWidth - _leftMargin - _rightMargin + 16, totalHeight);
            gfx.DrawRectangle(XBrushes.LightGray, bgRect);
            gfx.DrawRectangle(new XPen(XColors.Gray, 0.5), bgRect);

            _currentY += 8; // 上padding

            // 渲染代码行
            foreach (var line in lines)
            {
                SafeDrawString(gfx, line, font, XBrushes.Black, _leftMargin, _currentY);
                _currentY += lineHeight;
            }

            _currentY += 8; // 下padding
        }

        /// <summary>
        /// 渲染引用块
        /// </summary>
        private async Task RenderQuoteBlockAsync(XGraphics gfx, QuoteBlock quote, ConversionSettings settings, 
            Dictionary<string, (string processedPath, bool shouldRender)> imageGroups, CancellationToken cancellationToken)
        {
            // 绘制左边框
            var pen = new XPen(XColors.Gray, 3);
            var startY = _currentY;

            var oldLeftMargin = _leftMargin;
            _leftMargin += 20; // 增加左边距

            // 渲染引用内容
            foreach (var block in quote)
            {
                await RenderBlockAsync(gfx, block, settings, imageGroups, cancellationToken);
            }

            // 绘制左边框的实际高度
            gfx.DrawLine(pen, oldLeftMargin, startY, oldLeftMargin, _currentY);

            _leftMargin = oldLeftMargin; // 恢复左边距
        }

        /// <summary>
        /// 渲染列表
        /// </summary>
        private async Task RenderListAsync(XGraphics gfx, ListBlock list, ConversionSettings settings, 
            Dictionary<string, (string processedPath, bool shouldRender)> imageGroups, CancellationToken cancellationToken)
        {
            var itemNumber = 1;
            var oldLeftMargin = _leftMargin;
            _leftMargin += 15; // 缩进

            foreach (ListItemBlock item in list)
            {
                // 绘制列表标记
                var marker = list.IsOrdered ? $"{itemNumber}." : "•";
                var fontSize = settings.FontSize * 0.75; // 使用和正文一样的字体大小
                var font = CreateSafeFont(settings.FontFamily, fontSize, XFontStyleEx.Regular);
                SafeDrawString(gfx, marker, font, XBrushes.Black, oldLeftMargin, _currentY);

                // 渲染列表项内容
                foreach (var block in item)
                {
                    await RenderBlockAsync(gfx, block, settings, imageGroups, cancellationToken);
                }

                itemNumber++;
                _currentY += fontSize * 0.2; // 列表项间距
            }

            _leftMargin = oldLeftMargin; // 恢复缩进
        }

        /// <summary>
        /// 渲染表格
        /// </summary>
        private async Task RenderTableAsync(XGraphics gfx, Table table, ConversionSettings settings, 
            Dictionary<string, (string processedPath, bool shouldRender)> imageGroups, CancellationToken cancellationToken)
        {
            Console.WriteLine("=== 开始渲染Markdown表格 ===");
            
            var fontSize = settings.FontSize * 0.8; // 表格字体适中
            var font = CreateSafeFont(settings.FontFamily, fontSize, XFontStyleEx.Regular);
            var headerFont = CreateSafeFont(settings.FontFamily, fontSize, XFontStyleEx.Bold);
            
            var tableWidth = _pageWidth - _leftMargin - _rightMargin;
            var columnCount = Math.Max(1, table.ColumnDefinitions?.Count ?? 0);
            
            // 如果没有列定义，通过第一行确定列数
            if (columnCount == 1 && table.Count() > 0 && table.First() is TableRow firstRow)
            {
                columnCount = firstRow.Count();
                Console.WriteLine($"通过第一行确定列数: {columnCount}");
            }
            
            var columnWidth = tableWidth / columnCount;
            var cellPadding = 6; // 单元格内边距
            var minRowHeight = fontSize * settings.LineHeight + cellPadding * 2;

            var currentTableY = _currentY;
            var pen = new XPen(XColors.Black, 0.8);

            // 渲染表格行
            int rowIndex = 0;
            foreach (TableRow row in table)
            {
                Console.WriteLine($"处理表格第{++rowIndex}行, 列数: {row.Count()}");
                
                var isHeader = row.IsHeader;
                var currentFont = isHeader ? headerFont : font;
                var bgBrush = isHeader ? new XSolidBrush(XColor.FromArgb(240, 240, 240)) : XBrushes.White;

                // 计算这一行的实际高度（根据单元格内容）
                var actualRowHeight = minRowHeight;
                var cellX = _leftMargin;
                var cellContents = new List<string>();
                
                int cellIndex = 0;
                foreach (TableCell cell in row)
                {
                    cellIndex++;
                    Console.WriteLine($"处理单元格[{rowIndex},{cellIndex}]...");
                    
                    var cellText = "";
                    if (cell.Count > 0)
                    {
                        if (cell[0] is ParagraphBlock paragraph)
                        {
                            // 记录单元格内所有内联元素类型，用于调试
                            if (paragraph.Inline != null)
                            {
                                var typesInfo = string.Join(", ", paragraph.Inline.Select(i => $"{i.GetType().Name}"));
                                Console.WriteLine($"单元格内联元素类型: {typesInfo}");
                                
                                // 特别检查是否包含HTML内联元素
                                var htmlInlines = paragraph.Inline.Descendants().OfType<HtmlInline>().ToList();
                                if (htmlInlines.Any())
                                {
                                    Console.WriteLine($"发现HTML内联元素: {htmlInlines.Count}个");
                                    foreach (var html in htmlInlines)
                                    {
                                        Console.WriteLine($"HTML标签内容: '{html.Tag}'");
                                    }
                                }
                            }
                            
                            // 增强单元格文本提取逻辑，确保处理HTML内联元素
                            var sb = new StringBuilder();
                            ExtractCellTextRecursively(paragraph.Inline, sb);
                            cellText = sb.ToString();
                            
                            Console.WriteLine($"单元格[{rowIndex},{cellIndex}]文本提取结果: '{cellText}'");
                        }
                        else
                        {
                            // 处理其他类型的单元格内容
                            cellText = cell[0].ToString() ?? "";
                            Console.WriteLine($"非段落单元格类型: {cell[0].GetType().Name}, 内容: '{cellText}'");
                        }
                    }
                    cellContents.Add(cellText);
                    
                    // 计算文本高度
                    if (!string.IsNullOrEmpty(cellText))
                    {
                        var textWidth = columnWidth - cellPadding * 2;
                        var lines = WrapTextToLines(cellText, currentFont, textWidth, gfx);
                        var textHeight = lines.Count * currentFont.Height * settings.LineHeight + cellPadding * 2;
                        actualRowHeight = Math.Max(actualRowHeight, textHeight);
                    }
                }

                // 绘制行背景
                var rowRect = new XRect(_leftMargin, currentTableY, tableWidth, actualRowHeight);
                gfx.DrawRectangle(bgBrush, rowRect);

                // 渲染单元格内容和边框
                cellX = _leftMargin;
                for (int i = 0; i < cellContents.Count && i < columnCount; i++)
                {
                    var cellText = cellContents[i];
                    
                    // 绘制单元格边框
                    var cellRect = new XRect(cellX, currentTableY, columnWidth, actualRowHeight);
                    gfx.DrawRectangle(pen, cellRect);
                    
                    // 渲染单元格文本（居中对齐）
                    if (!string.IsNullOrEmpty(cellText))
                    {
                        var textRect = new XRect(cellX + cellPadding, currentTableY + cellPadding, 
                                                columnWidth - cellPadding * 2, actualRowHeight - cellPadding * 2);
                        Console.WriteLine($"渲染单元格[{rowIndex},{i+1}]文本: '{cellText}'");
                        RenderCellText(gfx, cellText, currentFont, XBrushes.Black, textRect, settings);
                    }

                    cellX += columnWidth;
                }

                currentTableY += actualRowHeight;
            }

            Console.WriteLine($"=== 表格渲染完成: {rowIndex}行 ===");
            _currentY = currentTableY + fontSize * 0.5; // 表格后留一点空间
        }

        /// <summary>
        /// 递归提取单元格内文本，包括处理HTML内联元素
        /// </summary>
        private void ExtractCellTextRecursively(ContainerInline inlines, StringBuilder sb)
        {
            if (inlines == null) return;
            
            // 创建一个字符串构建器来存储HTML内容
            var htmlContent = new StringBuilder();
            var inHtmlTag = false;
            var fontContentBuffer = new StringBuilder(); // 存储font标签间的内容
            
            foreach (var inline in inlines)
            {
                if (inline is LiteralInline literal)
                {
                    var content = literal.Content.ToString();
                    
                    // 检查是否包含font标签
                    if (content.Contains("<font") || content.Contains("</font>"))
                    {
                        Console.WriteLine($"在LiteralInline中发现font标签: '{content}'");
                        
                        if (content.Contains("<font") && content.Contains("</font>"))
                        {
                            // 完整的font标签和内容
                            var fontContent = ExtractFontTagContent(content);
                            sb.Append(fontContent);
                        }
                        else if (content.Contains("<font"))
                        {
                            // 只有开始标签
                            inHtmlTag = true;
                            htmlContent.Append(content);
                            
                            // 提取开始标签后的文本
                            var tagEndPos = content.IndexOf('>');
                            if (tagEndPos >= 0 && tagEndPos < content.Length - 1)
                            {
                                var textAfterTag = content.Substring(tagEndPos + 1);
                                fontContentBuffer.Append(textAfterTag);
                                Console.WriteLine($"缓存font内开始标签后内容: '{textAfterTag}'");
                            }
                        }
                        else if (content.Contains("</font>"))
                        {
                            // 只有结束标签
                            var textBeforeTag = content.Split(new[] { "</font>" }, StringSplitOptions.None)[0];
                            fontContentBuffer.Append(textBeforeTag);
                            Console.WriteLine($"缓存font内结束标签前内容: '{textBeforeTag}'");
                            
                            // 完成整个font标签处理
                            inHtmlTag = false;
                            htmlContent.Append(content);
                            var completeTag = htmlContent.ToString();
                            
                            // 如果没有完整标签但有缓存的内容
                            if (fontContentBuffer.Length > 0)
                            {
                                Console.WriteLine($"使用缓存内容: '{fontContentBuffer}'");
                                sb.Append(fontContentBuffer.ToString());
                                fontContentBuffer.Clear();
                            }
                            else
                            {
                                // 尝试从完整标签中提取
                                var fontContent = ExtractFontTagContent(completeTag);
                                sb.Append(fontContent);
                            }
                            
                            htmlContent.Clear();
                        }
                    }
                    else if (inHtmlTag)
                    {
                        // 在font标签内的内容需要缓存
                        fontContentBuffer.Append(content);
                        htmlContent.Append(content);
                        Console.WriteLine($"缓存font标签内文本: '{content}'");
                    }
                    else 
                    {
                        sb.Append(content);
                        Console.WriteLine($"提取文本内容: '{content}'");
                    }
                }
                else if (inline is EmphasisInline emphasis)
                {
                    // 可以在这里添加强调效果（如加粗、斜体），但目前只提取文本
                    Console.WriteLine("处理强调文本...");
                    ExtractCellTextRecursively(emphasis, sb);
                }
                else if (inline is LinkInline link && !link.IsImage)
                {
                    // 处理链接文本
                    Console.WriteLine("处理链接文本...");
                    ExtractCellTextRecursively(link, sb);
                }
                else if (inline is LineBreakInline)
                {
                    if (inHtmlTag)
                    {
                        // 在font标签内的换行
                        fontContentBuffer.Append(" ");
                        htmlContent.Append(" ");
                    }
                    else
                    {
                        sb.Append(" ");
                    }
                    Console.WriteLine("处理换行符 -> 空格");
                }
                else if (inline is HtmlInline html)
                {
                    try 
                    {
                        Console.WriteLine($"处理HTML标签: {html.Tag}");
                        
                        // 收集HTML标签内容用于后续处理
                        if (html.Tag.StartsWith("<font"))
                        {
                            // 开始收集HTML内容
                            inHtmlTag = true;
                            htmlContent.Append(html.Tag);
                        }
                        else if (html.Tag == "</font>" && inHtmlTag)
                        {
                            // 结束收集，处理完整标签
                            htmlContent.Append(html.Tag);
                            var completeTag = htmlContent.ToString();
                            
                            Console.WriteLine($"处理完整font标签: {completeTag}");
                            
                            // 优先使用缓存的内容
                            if (fontContentBuffer.Length > 0)
                            {
                                var bufferContent = fontContentBuffer.ToString();
                                Console.WriteLine($"使用已缓存的font内容: '{bufferContent}'");
                                sb.Append(bufferContent);
                            }
                            else
                            {
                                // 从font标签中提取文本内容
                                var content = new HtmlDocument();
                                content.LoadHtml(completeTag);
                                var text = content.DocumentNode.InnerText;
                                
                                if (!string.IsNullOrEmpty(text))
                                {
                                    Console.WriteLine($"提取到font标签内容: '{text}'");
                                    sb.Append(text);
                                }
                                else
                                {
                                    // 尝试使用正则表达式提取
                                    var fontContent = ExtractFontTagContent(completeTag);
                                    sb.Append(fontContent);
                                }
                            }
                            
                            // 重置状态
                            inHtmlTag = false;
                            htmlContent.Clear();
                            fontContentBuffer.Clear();
                        }
                        else if (inHtmlTag)
                        {
                            // 如果在HTML标签内，收集内容
                            htmlContent.Append(html.Tag);
                            
                            // 尝试提取可能的文本内容
                            var strippedTag = System.Text.RegularExpressions.Regex.Replace(html.Tag, "<[^>]*>", "");
                            if (!string.IsNullOrEmpty(strippedTag))
                            {
                                fontContentBuffer.Append(strippedTag);
                                Console.WriteLine($"向缓存添加的内容: '{strippedTag}'");
                            }
                        }
                        else
                        {
                            // 直接处理HTML标签
                            string htmlTextContent = ExtractTextFromHtmlInline(html.Tag);
                            sb.Append(htmlTextContent);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"提取HTML内容时出错: {ex.Message}");
                        // 出错时回退到正则提取
                        var fallbackRegex = "<[^>]*>";
                        var fallbackText = System.Text.RegularExpressions.Regex.Replace(html.Tag, fallbackRegex, "");
                        Console.WriteLine($"回退策略提取结果: '{fallbackText}'");
                        sb.Append(fallbackText);
                    }
                }
                else if (inline is ContainerInline container)
                {
                    Console.WriteLine("处理嵌套容器...");
                    ExtractCellTextRecursively(container, sb);
                }
                else
                {
                    Console.WriteLine($"未处理的内联类型: {inline.GetType().Name}");
                }
            }
            
            // 如果收集了HTML内容但未处理，尝试最后处理
            if (inHtmlTag && htmlContent.Length > 0)
            {
                var remainingHtml = htmlContent.ToString();
                Console.WriteLine($"处理未完成的HTML内容: {remainingHtml}");
                
                // 优先使用已缓存的内容
                if (fontContentBuffer.Length > 0)
                {
                    var bufferContent = fontContentBuffer.ToString();
                    Console.WriteLine($"使用未闭合标签中的缓存内容: '{bufferContent}'");
                    sb.Append(bufferContent);
                }
                else
                {
                    // 尝试解析内容
                    try {
                        var doc = new HtmlDocument();
                        doc.LoadHtml(remainingHtml);
                        var text = doc.DocumentNode.InnerText;
                        if (!string.IsNullOrEmpty(text))
                        {
                            sb.Append(text);
                        }
                        else
                        {
                            // 使用正则表达式移除HTML标签
                            var stripped = System.Text.RegularExpressions.Regex.Replace(remainingHtml, "<[^>]*>", "");
                            sb.Append(stripped);
                        }
                    }
                    catch {
                        var stripped = System.Text.RegularExpressions.Regex.Replace(remainingHtml, "<[^>]*>", "");
                        sb.Append(stripped);
                    }
                }
            }
        }

        /// <summary>
        /// 将文本按行分割（用于计算表格单元格高度）
        /// </summary>
        private List<string> WrapTextToLines(string text, XFont font, double maxWidth, XGraphics gfx)
        {
            var lines = new List<string>();
            if (string.IsNullOrEmpty(text)) return lines;

            var segments = SplitTextForWrapping(text);
            var currentLine = string.Empty;

            foreach (var segment in segments)
            {
                if (segment == "\n")
                {
                    lines.Add(currentLine);
                    currentLine = string.Empty;
                    continue;
                }

                var testLine = string.IsNullOrEmpty(currentLine) ? segment : currentLine + segment;
                var testSize = SafeMeasureString(gfx, testLine, font);

                if (testSize.Width <= maxWidth)
                {
                    currentLine = testLine;
                }
                else
                {
                    if (!string.IsNullOrEmpty(currentLine))
                    {
                        lines.Add(currentLine);
                    }
                    currentLine = segment;
                }
            }

            if (!string.IsNullOrEmpty(currentLine))
            {
                lines.Add(currentLine);
            }

            return lines.Count > 0 ? lines : new List<string> { text };
        }

        /// <summary>
        /// 渲染单元格文本（支持垂直居中）
        /// </summary>
        private void RenderCellText(XGraphics gfx, string text, XFont font, XBrush brush, XRect rect, ConversionSettings settings)
        {
            var lines = WrapTextToLines(text, font, rect.Width, gfx);
            var lineHeight = font.Height * settings.LineHeight;
            var totalTextHeight = lines.Count * lineHeight;
            
            // 垂直居中
            var startY = rect.Y + (rect.Height - totalTextHeight) / 2 + font.Height * 0.8;
            
            for (int i = 0; i < lines.Count; i++)
            {
                var line = lines[i];
                var y = startY + i * lineHeight;
                
                // 确保不超出单元格范围
                if (y + font.Height <= rect.Y + rect.Height)
                {
                    SafeDrawString(gfx, line, font, brush, rect.X, y);
                }
            }
        }

        /// <summary>
        /// 渲染HTML块（支持HTML表格）
        /// </summary>
        private async Task RenderHtmlBlockAsync(XGraphics gfx, dynamic htmlBlock, ConversionSettings settings, 
            Dictionary<string, (string processedPath, bool shouldRender)> imageGroups, CancellationToken cancellationToken)
        {
            try
            {
                var htmlContent = htmlBlock.Lines?.ToString() ?? "";
                if (string.IsNullOrEmpty(htmlContent)) return;

                System.Diagnostics.Debug.WriteLine($"正在处理HTML块: {htmlContent.Substring(0, Math.Min(100, htmlContent.Length))}...");

                // 解析HTML内容
                var htmlDoc = new HtmlDocument();
                htmlDoc.LoadHtml(htmlContent);

                // 查找表格元素
                var tableNodes = htmlDoc.DocumentNode.SelectNodes("//table");
                if (tableNodes != null)
                {
                    foreach (var tableNode in tableNodes)
                    {
                        await RenderHtmlTableAsync(gfx, tableNode, settings);
                    }
                }
                else
                {
                    // 如果不是表格，尝试提取纯文本内容
                    var textContent = htmlDoc.DocumentNode.InnerText;
                    if (!string.IsNullOrEmpty(textContent))
                    {
                        var fontSize = settings.FontSize * 0.75;
                        var font = CreateSafeFont(settings.FontFamily, fontSize, XFontStyleEx.Regular);
                        var rect = new XRect(_leftMargin, _currentY, _pageWidth - _leftMargin - _rightMargin, fontSize * 10);
                        var height = RenderTextWithWrapping(gfx, textContent, font, XBrushes.Black, rect, settings);
                        _currentY += height;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"HTML块渲染失败: {ex.Message}");
                _currentY += settings.FontSize;
            }
        }

        /// <summary>
        /// 渲染HTML表格
        /// </summary>
        private async Task RenderHtmlTableAsync(XGraphics gfx, HtmlNode tableNode, ConversionSettings settings)
        {
            try
            {
                var fontSize = settings.FontSize * 0.7;
                var font = CreateSafeFont(settings.FontFamily, fontSize, XFontStyleEx.Regular);
                var headerFont = CreateSafeFont(settings.FontFamily, fontSize, XFontStyleEx.Bold);

                // 获取所有行
                var rows = tableNode.SelectNodes(".//tr")?.ToList() ?? new List<HtmlNode>();
                if (rows.Count == 0) 
                {
                    System.Diagnostics.Debug.WriteLine("HTML表格：未找到行元素");
                    return;
                }

                // 计算真实列数（考虑所有行，找到最大列数）
                var columnCount = 0;
                foreach (var row in rows)
                {
                    var cellsInRow = row.SelectNodes(".//td|.//th")?.Count ?? 0;
                    columnCount = Math.Max(columnCount, cellsInRow);
                }

                if (columnCount == 0) 
                {
                    System.Diagnostics.Debug.WriteLine("HTML表格：未找到单元格");
                    return;
                }

                var tableWidth = _pageWidth - _leftMargin - _rightMargin;
                var columnWidth = tableWidth / columnCount;

                // 检查是否有表头
                var hasHeader = tableNode.SelectSingleNode(".//thead") != null || 
                               rows[0].SelectNodes(".//th")?.Count > 0;

                var pen = new XPen(XColors.Black, 0.8);
                var currentTableY = _currentY;

                System.Diagnostics.Debug.WriteLine($"开始渲染HTML表格: {rows.Count}行 x {columnCount}列, 表头: {hasHeader}");

                for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
                {
                    var row = rows[rowIndex];
                    var cells = row.SelectNodes(".//td|.//th")?.ToList() ?? new List<HtmlNode>();
                    
                    System.Diagnostics.Debug.WriteLine($"第{rowIndex + 1}行: 找到{cells.Count}个单元格");

                    // 计算行高（根据最高单元格内容）
                    var maxRowHeight = fontSize * 1.8; // 增加最小行高
                    
                    for (int colIndex = 0; colIndex < Math.Min(cells.Count, columnCount); colIndex++)
                    {
                        var cellText = ExtractCellText(cells[colIndex]);
                        System.Diagnostics.Debug.WriteLine($"  单元格[{rowIndex},{colIndex}]: '{cellText}'");
                        
                        var textFont = hasHeader && rowIndex == 0 ? headerFont : font;
                        var lines = WrapTextToLines(cellText, textFont, columnWidth - 12, gfx);
                        var cellHeight = lines.Count * fontSize * 1.2 + 12; // 6px上下内边距
                        maxRowHeight = Math.Max(maxRowHeight, cellHeight);
                    }

                    // 绘制行背景（表头使用浅灰色）
                    if (hasHeader && rowIndex == 0)
                    {
                        var headerBrush = new XSolidBrush(XColor.FromArgb(240, 240, 240));
                        var headerRect = new XRect(_leftMargin, currentTableY, tableWidth, maxRowHeight);
                        gfx.DrawRectangle(headerBrush, headerRect);
                    }
                    
                    // 绘制单元格边框和内容
                    for (int colIndex = 0; colIndex < columnCount; colIndex++)
                    {
                        var cellX = _leftMargin + colIndex * columnWidth;
                        var cellRect = new XRect(cellX, currentTableY, columnWidth, maxRowHeight);
                        
                        // 绘制单元格边框
                        gfx.DrawRectangle(pen, cellRect);
                        
                        // 渲染单元格文本（如果单元格存在）
                        if (colIndex < cells.Count)
                        {
                            var cellText = ExtractCellText(cells[colIndex]);
                            if (!string.IsNullOrEmpty(cellText))
                            {
                                var textFont = hasHeader && rowIndex == 0 ? headerFont : font;
                                var textBrush = XBrushes.Black;
                                var textRect = new XRect(cellX + 6, currentTableY + 6, columnWidth - 12, maxRowHeight - 12);
                                RenderCellText(gfx, cellText, textFont, textBrush, textRect, settings);
                            }
                        }
                    }

                    currentTableY += maxRowHeight;
                }

                _currentY = currentTableY + fontSize * 0.5; // 表格后的间距
                System.Diagnostics.Debug.WriteLine($"HTML表格渲染完成，总高度: {currentTableY - _currentY + fontSize * 0.5}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"HTML表格渲染失败: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"异常堆栈: {ex.StackTrace}");
                _currentY += settings.FontSize;
            }
        }

        /// <summary>
        /// 提取单元格文本（处理嵌套HTML元素）
        /// </summary>
        private string ExtractCellText(HtmlNode cellNode)
        {
            if (cellNode == null) return "";

            var text = new StringBuilder();
            ExtractTextRecursive(cellNode, text);
            return text.ToString().Trim();
        }

        /// <summary>
        /// 递归提取HTML节点中的文本
        /// </summary>
        private void ExtractTextRecursive(HtmlNode node, StringBuilder text)
        {
            if (node == null) return;

            foreach (var child in node.ChildNodes)
            {
                if (child.NodeType == HtmlNodeType.Text)
                {
                    // 直接文本节点
                    var nodeText = HtmlEntity.DeEntitize(child.InnerText);
                    text.Append(nodeText);
                    Console.WriteLine($"提取到文本节点: '{nodeText}'");
                }
                else if (child.NodeType == HtmlNodeType.Element)
                {
                    // HTML元素节点，递归处理
                    Console.WriteLine($"处理HTML元素: <{child.Name}>");
                    switch (child.Name.ToLower())
                    {
                        case "br":
                            text.Append(" "); // 换行替换为空格
                            Console.WriteLine("处理<br>标签 -> 空格");
                            break;
                        case "font":
                            Console.WriteLine($"处理<font>标签，属性: {string.Join(", ", child.Attributes.Select(a => $"{a.Name}='{a.Value}'"))}");
                            ExtractTextRecursive(child, text);
                            break;
                        case "span":
                        case "b":
                        case "strong":
                        case "i":
                        case "em":
                        case "u":
                        case "a":
                        case "small":
                        case "sub":
                        case "sup":
                            // 对于文本格式化标签，直接提取内容
                            Console.WriteLine($"处理<{child.Name}>标签");
                            ExtractTextRecursive(child, text);
                            break;
                        default:
                            // 尝试提取内部文本
                            if (!string.IsNullOrEmpty(child.InnerText))
                            {
                                var innerText = HtmlEntity.DeEntitize(child.InnerText);
                                text.Append(innerText);
                                Console.WriteLine($"从<{child.Name}>提取内部文本: '{innerText}'");
                            }
                            else
                            {
                                // 递归处理所有嵌套内容
                                Console.WriteLine($"递归处理<{child.Name}>的子节点");
                                ExtractTextRecursive(child, text);
                            }
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// 渲染水平分割线
        /// </summary>
        private void RenderHorizontalRule(XGraphics gfx, ConversionSettings settings)
        {
            var fontSize = settings.FontSize * 0.75;
            _currentY += fontSize * 0.5;
            var pen = new XPen(XColors.Gray, 0.8);
            gfx.DrawLine(pen, _leftMargin, _currentY, _pageWidth - _rightMargin, _currentY);
            _currentY += fontSize * 0.5;
        }

        #region 辅助方法

        /// <summary>
        /// 统计块数量
        /// </summary>
        private int CountBlocks(MarkdownDocument document)
        {
            return document.Count();
        }

        /// <summary>
        /// 增强的块高度估算（支持图片）
        /// </summary>
        private async Task<double> EstimateBlockHeightAsync(Block block, ConversionSettings settings, 
            Dictionary<string, (string processedPath, bool shouldRender)> imageGroups)
        {
            return await Task.Run(() => EstimateBlockHeight(block, settings, imageGroups));
        }

        /// <summary>
        /// 估算块的高度（包含图片处理）
        /// </summary>
        private double EstimateBlockHeight(Block block, ConversionSettings settings, 
            Dictionary<string, (string processedPath, bool shouldRender)>? imageGroups = null)
        {
            return block switch
            {
                HeadingBlock => settings.FontSize * 3,
                ParagraphBlock paragraph => EstimateParagraphHeight(paragraph, settings, imageGroups),
                CodeBlock code => code.Lines.Count * settings.FontSize * settings.LineHeight + 20,
                Table table => table.Count() * settings.FontSize * settings.LineHeight + 20,
                ListBlock list => list.Count() * settings.FontSize * settings.LineHeight * 2,
                _ => settings.FontSize * settings.LineHeight * 2
            };
        }

        /// <summary>
        /// 估算段落高度（包含图片智能压缩）
        /// </summary>
        private double EstimateParagraphHeight(ParagraphBlock paragraph, ConversionSettings settings, 
            Dictionary<string, (string processedPath, bool shouldRender)>? imageGroups = null)
        {
            if (paragraph.Inline == null) 
                return settings.FontSize * settings.LineHeight;

            // 检查是否包含图片
            if (ContainsImage(paragraph.Inline))
            {
                // 有图片的段落，使用智能压缩估算图片高度
                var hasRenderableImage = false;
                if (imageGroups != null)
                {
                    foreach (var inline in paragraph.Inline.Descendants())
                    {
                        if (inline is LinkInline { IsImage: true } imageLink && !string.IsNullOrEmpty(imageLink.Url))
                        {
                            if (!imageGroups.TryGetValue(imageLink.Url, out var imageInfo) || imageInfo.shouldRender)
                            {
                                hasRenderableImage = true;
                                break;
                            }
                        }
                    }
                }
                else
                {
                    hasRenderableImage = true; // 保守估计
                }

                if (hasRenderableImage)
                {
                    // 智能估算图片高度：考虑默认50%缩放和可能的智能压缩
                    var availableHeight = _pageHeight - _currentY - _bottomMargin - 50;
                    var maxWidth = _pageWidth - _leftMargin - _rightMargin;
                    
                    // 假设图片为常见比例（16:9 或 4:3），估算高度
                    var estimatedImageWidth = maxWidth * 0.5; // 默认50%缩放
                    var estimatedImageHeight = estimatedImageWidth * 0.6; // 假设宽高比为1.67:1
                    
                    // 如果估算高度超过可用空间，应用智能压缩
                    if (estimatedImageHeight > availableHeight)
                    {
                        var compressionRatio = availableHeight / estimatedImageHeight;
                        var minCompressionRatio = 0.5; // 最多压缩50%
                        
                        if (compressionRatio >= minCompressionRatio)
                        {
                            estimatedImageHeight = availableHeight; // 压缩后能放下
                        }
                        else
                        {
                            estimatedImageHeight = estimatedImageWidth * 0.6 * minCompressionRatio; // 最小尺寸
                        }
                    }
                    
                    return Math.Min(estimatedImageHeight + 20, Math.Max(200, settings.FontSize * 8)); // 至少200像素
                }
            }

            // 普通文本段落
            var text = ExtractTextFromInlines(paragraph.Inline);
            var lineCount = Math.Max(1, text.Length / 50); // 估算行数：每行约50字符
            return lineCount * settings.FontSize * settings.LineHeight + settings.FontSize * 0.5;
        }

        /// <summary>
        /// 从内联元素提取文本
        /// </summary>
        private string ExtractTextFromInlines(ContainerInline? inlines)
        {
            if (inlines == null) return string.Empty;

            var text = new StringBuilder();
            foreach (var inline in inlines)
            {
                text.Append(ExtractTextFromInline(inline));
            }
            return text.ToString();
        }

        /// <summary>
        /// 从单个内联元素提取文本（增强支持HTML内联元素）
        /// </summary>
        private string ExtractTextFromInline(Inline inline)
        {
            return inline switch
            {
                LiteralInline literal => literal.Content.ToString(),
                CodeInline code => code.Content,
                EmphasisInline emphasis => ExtractTextFromInlines(emphasis),
                LinkInline link when !link.IsImage => ExtractTextFromInlines(link),
                LineBreakInline => " ",
                HtmlInline html => ExtractTextFromHtmlInline(html.Tag),
                _ => ""
            };
        }

        /// <summary>
        /// 专门处理font标签，确保能提取到内容
        /// </summary>
        private string ExtractFontTagContent(string htmlTag)
        {
            Console.WriteLine($"特殊处理font标签: '{htmlTag}'");
            
            try
            {
                // 如果是结束标签，直接返回空
                if (htmlTag == "</font>")
                {
                    return "";
                }
                
                // 分析标签内容是否为开始标签
                if (htmlTag.StartsWith("<font") && !htmlTag.Contains("</font>"))
                {
                    // 检查是否包含内部文本内容
                    var tagEndPos = htmlTag.IndexOf('>');
                    if (tagEndPos >= 0 && tagEndPos < htmlTag.Length - 1)
                    {
                        // 可能包含标签后的文本内容
                        var textContent = htmlTag.Substring(tagEndPos + 1);
                        Console.WriteLine($"从开始标签中提取内容: '{textContent}'");
                        return textContent;
                    }
                    
                    // 这只是开始标签，没有文本内容
                    Console.WriteLine("这只是font开始标签，没有文本内容");
                    return "";
                }
                
                // 直接提取font标签内容的正则表达式
                string fontPattern = @"<font\s+[^>]*>(.*?)</font>";
                var fontMatch = System.Text.RegularExpressions.Regex.Match(
                    htmlTag, 
                    fontPattern, 
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase | 
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                
                if (fontMatch.Success && fontMatch.Groups.Count > 1)
                {
                    var content = fontMatch.Groups[1].Value;
                    Console.WriteLine($"font标签内容提取成功: '{content}'");
                    return content;
                }
                
                // 检查是否有嵌套的font标签
                string nestedFontPattern = @"<font[^>]*>.*?<font[^>]*>(.*?)</font>.*?</font>";
                var nestedFontMatch = System.Text.RegularExpressions.Regex.Match(
                    htmlTag,
                    nestedFontPattern,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase | 
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                
                if (nestedFontMatch.Success && nestedFontMatch.Groups.Count > 1)
                {
                    var content = nestedFontMatch.Groups[1].Value;
                    Console.WriteLine($"嵌套font标签内容提取成功: '{content}'");
                    return content;
                }
                
                // 尝试解析完整的HTML字符串
                try 
                {
                    var htmlDoc = new HtmlDocument();
                    htmlDoc.LoadHtml(htmlTag);
                    var fontNode = htmlDoc.DocumentNode.SelectSingleNode("//font");
                    if (fontNode != null)
                    {
                        var innerText = fontNode.InnerText;
                        Console.WriteLine($"通过HtmlAgilityPack提取font内容: '{innerText}'");
                        return innerText;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HtmlAgilityPack解析font标签失败: {ex.Message}");
                }
                
                // 尝试提取颜色信息
                string colorPattern = @"<font\s+[^>]*color=""?([^"">]+)""?[^>]*>";
                var colorMatch = System.Text.RegularExpressions.Regex.Match(
                    htmlTag, colorPattern, 
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                
                if (colorMatch.Success)
                {
                    Console.WriteLine($"检测到font标签颜色: {colorMatch.Groups[1].Value}");
                }
                
                // 如果上面的方法失败，检查是否有文本内容夹在标签之间
                var startTagEndPos = htmlTag.IndexOf('>');
                if (startTagEndPos >= 0 && startTagEndPos < htmlTag.Length - 1)
                {
                    var possibleText = htmlTag.Substring(startTagEndPos + 1);
                    // 如果文本包含结束标签，则移除
                    var endTagPos = possibleText.IndexOf("</font>");
                    if (endTagPos >= 0)
                    {
                        possibleText = possibleText.Substring(0, endTagPos);
                    }
                    
                    if (!string.IsNullOrEmpty(possibleText))
                    {
                        Console.WriteLine($"标签间直接提取文本: '{possibleText}'");
                        return possibleText;
                    }
                }
                
                // 如果上面的方法失败，尝试直接移除所有HTML标签
                var strippedContent = System.Text.RegularExpressions.Regex.Replace(
                    htmlTag, "<[^>]*>", "");
                
                Console.WriteLine($"最终font处理结果: '{strippedContent}'");
                return strippedContent;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"font标签处理出错: {ex.Message}");
                // 简单回退：移除所有HTML标签
                return System.Text.RegularExpressions.Regex.Replace(htmlTag, "<[^>]*>", "");
            }
        }

        /// <summary>
        /// 从HTML内联元素提取文本
        /// </summary>
        private string ExtractTextFromHtmlInline(string htmlTag)
        {
            try
            {
                if (string.IsNullOrEmpty(htmlTag)) return "";

                Console.WriteLine($"HTML内联文本提取开始: '{htmlTag}'");
                
                // 特殊处理font标签
                if (htmlTag.Contains("<font"))
                {
                    return ExtractFontTagContent(htmlTag);
                }

                // 处理其他常见的HTML内联标签
                if (htmlTag.Contains("<span") || 
                    htmlTag.Contains("<b") || htmlTag.Contains("<i") ||
                    htmlTag.Contains("<strong") || htmlTag.Contains("<em"))
                {
                    Console.WriteLine("检测到特殊HTML标签，使用HtmlAgilityPack解析");
                    // 使用HtmlAgilityPack解析HTML标签
                    var htmlDoc = new HtmlDocument();
                    htmlDoc.LoadHtml(htmlTag);

                    var textBuilder = new StringBuilder();
                    ExtractTextRecursive(htmlDoc.DocumentNode, textBuilder);
                    
                    var result = textBuilder.ToString().Trim();
                    Console.WriteLine($"HtmlAgilityPack提取结果: '{result}'");
                    
                    // 如果结果为空，尝试使用正则表达式提取
                    if (string.IsNullOrEmpty(result))
                    {
                        Console.WriteLine("HtmlAgilityPack提取为空，尝试正则表达式");
                        var stripTagsRegex = "<[^>]*>";
                        result = System.Text.RegularExpressions.Regex.Replace(htmlTag, stripTagsRegex, "");
                        Console.WriteLine($"正则表达式提取结果: '{result}'");
                    }
                    
                    return result;
                }
                else
                {
                    // 对于其他简单HTML标签，使用正则表达式直接提取
                    Console.WriteLine("使用正则表达式直接提取HTML内容");
                    var stripTagsRegex = "<[^>]*>";
                    var result = System.Text.RegularExpressions.Regex.Replace(htmlTag, stripTagsRegex, "");
                    Console.WriteLine($"提取结果: '{result}'");
                    return result;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"HTML内联文本提取失败: {ex.Message}");
                // 如果解析失败，尝试简单的正则提取
                var stripTagsRegex = "<[^>]*>";
                var result = System.Text.RegularExpressions.Regex.Replace(htmlTag, stripTagsRegex, "");
                Console.WriteLine($"异常后正则提取结果: '{result}'");
                return result;
            }
        }

        /// <summary>
        /// 创建安全字体
        /// </summary>
        private XFont CreateSafeFont(string fontFamily, double fontSize, XFontStyleEx style)
        {
            var fallbackFonts = new[]
            {
                fontFamily,
                "Microsoft YaHei",
                "SimSun", 
                "SimHei",
                "Arial Unicode MS",
                "Arial",
                "Calibri",
                "Verdana",
                "Times New Roman"
            };

            foreach (var font in fallbackFonts.Distinct())
            {
                try
                {
                    var xFont = new XFont(font, fontSize, style);
                    return xFont;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"字体创建失败 {font}: {ex.Message}");
                    continue;
                }
            }

            // 最后回退
            try
            {
                return new XFont("Arial", fontSize, XFontStyleEx.Regular);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"无法创建任何字体: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 安全测量字符串
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
            }
        }

        /// <summary>
        /// 带换行的文本渲染，支持中文自动换行
        /// </summary>
        private double RenderTextWithWrapping(XGraphics gfx, string text, XFont font, XBrush brush, XRect rect, ConversionSettings settings)
        {
            if (string.IsNullOrEmpty(text))
                return 0;

            var segments = SplitTextForWrapping(text);
            var currentY = rect.Y;
            var currentLine = string.Empty;
            var lineHeight = font.Height * settings.LineHeight;
            var totalHeight = 0.0;

            foreach (var segment in segments)
            {
                if (segment == "\n")
                {
                    // 换行
                    if (!string.IsNullOrEmpty(currentLine))
                    {
                        SafeDrawString(gfx, currentLine, font, brush, rect.X, currentY);
                        totalHeight += lineHeight;
                        currentY += lineHeight;
                        currentLine = string.Empty;
                    }
                    continue;
                }

                var testLine = string.IsNullOrEmpty(currentLine) ? segment : currentLine + segment;
                var testSize = SafeMeasureString(gfx, testLine, font);

                if (testSize.Width <= rect.Width)
                {
                    currentLine = testLine;
                }
                else
                {
                    // 当前行装不下，需要换行
                    if (!string.IsNullOrEmpty(currentLine))
                    {
                        SafeDrawString(gfx, currentLine, font, brush, rect.X, currentY);
                        totalHeight += lineHeight;
                        currentY += lineHeight;
                    }
                    currentLine = segment;
                }
            }

            // 绘制最后一行
            if (!string.IsNullOrEmpty(currentLine))
            {
                SafeDrawString(gfx, currentLine, font, brush, rect.X, currentY);
                totalHeight += lineHeight;
            }

                         return Math.Max(totalHeight, font.Height);
         }

        /// <summary>
        /// 分割文本以支持更好的换行，特别是中文
        /// </summary>
        private List<string> SplitTextForWrapping(string text)
        {
            var segments = new List<string>();
            var currentSegment = string.Empty;

            for (int i = 0; i < text.Length; i++)
            {
                var ch = text[i];

                // 换行符单独处理
                if (ch == '\n' || ch == '\r')
                {
                    if (!string.IsNullOrEmpty(currentSegment))
                    {
                        segments.Add(currentSegment);
                        currentSegment = string.Empty;
                    }
                    if (ch == '\n')
                        segments.Add("\n");
                    continue;
                }

                // 英文单词：遇到空格就分段
                if (ch == ' ')
                {
                    if (!string.IsNullOrEmpty(currentSegment))
                    {
                        segments.Add(currentSegment);
                        currentSegment = string.Empty;
                    }
                    segments.Add(" ");
                    continue;
                }

                // 中文字符：每个字符都可以作为分割点
                if (IsChinese(ch))
                {
                    if (!string.IsNullOrEmpty(currentSegment))
                    {
                        segments.Add(currentSegment);
                        currentSegment = string.Empty;
                    }
                    segments.Add(ch.ToString());
                    continue;
                }

                // 标点符号和其他字符
                if (IsPunctuation(ch))
                {
                    if (!string.IsNullOrEmpty(currentSegment))
                    {
                        segments.Add(currentSegment);
                        currentSegment = string.Empty;
                    }
                    segments.Add(ch.ToString());
                    continue;
                }

                // 普通字符（英文字母、数字等）
                currentSegment += ch;
            }

            if (!string.IsNullOrEmpty(currentSegment))
            {
                segments.Add(currentSegment);
            }

            return segments;
        }

        /// <summary>
        /// 渲染过长的文本段，强制按字符拆分
        /// </summary>
        private double RenderLongSegment(XGraphics gfx, string segment, XFont font, XBrush brush, double x, double startY, double maxWidth, double lineHeight)
        {
            var totalHeight = 0.0;
            var y = startY;
            var currentLine = string.Empty;

            foreach (char ch in segment)
            {
                var testLine = currentLine + ch;
                var testSize = SafeMeasureString(gfx, testLine, font);

                if (testSize.Width <= maxWidth)
                {
                    currentLine = testLine;
                }
                else
                {
                    if (!string.IsNullOrEmpty(currentLine))
                    {
                        SafeDrawString(gfx, currentLine, font, brush, x, y);
                        y += lineHeight;
                        totalHeight += lineHeight;
                    }
                    currentLine = ch.ToString();
                }
            }

            if (!string.IsNullOrEmpty(currentLine))
            {
                SafeDrawString(gfx, currentLine, font, brush, x, y);
                totalHeight += lineHeight;
            }

            return totalHeight;
        }

        /// <summary>
        /// 判断是否为中文字符
        /// </summary>
        private bool IsChinese(char ch)
        {
            return (ch >= 0x4e00 && ch <= 0x9fff) ||  // 基本汉字
                   (ch >= 0x3400 && ch <= 0x4dbf) ||  // 扩展A
                   (ch >= 0x20000 && ch <= 0x2a6df) || // 扩展B
                   (ch >= 0x2a700 && ch <= 0x2b73f) || // 扩展C
                   (ch >= 0x2b740 && ch <= 0x2b81f) || // 扩展D
                   (ch >= 0x2b820 && ch <= 0x2ceaf);   // 扩展E
        }

        /// <summary>
        /// 判断是否为标点符号
        /// </summary>
        private bool IsPunctuation(char ch)
        {
            return char.IsPunctuation(ch) ||
                   (ch >= 0x3000 && ch <= 0x303f) ||  // CJK符号
                   (ch >= 0xff00 && ch <= 0xffef);    // 全角符号
        }

        /// <summary>
        /// 触发进度事件
        /// </summary>
        protected virtual void OnProgressChanged(DirectPdfGenerationProgressEventArgs e)
        {
            ProgressChanged?.Invoke(this, e);
        }

        #endregion
    }

    /// <summary>
    /// 目录项
    /// </summary>
    public class TocItem
    {
        public int Level { get; set; }
        public string Title { get; set; } = string.Empty;
        public int PageNumber { get; set; }
    }

    /// <summary>
    /// 直接PDF生成进度事件参数
    /// </summary>
    public class DirectPdfGenerationProgressEventArgs : EventArgs
    {
        public int Progress { get; }
        public string Message { get; }

        public DirectPdfGenerationProgressEventArgs(int progress, string message)
        {
            Progress = Math.Max(0, Math.Min(100, progress));
            Message = message;
        }
    }
} 
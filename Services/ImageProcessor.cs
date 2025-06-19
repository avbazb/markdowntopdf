using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text;
using MarkdownToPdf.Models;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.Formats.Jpeg;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.Formats.Webp;
using System.Xml;
using System.Text.RegularExpressions;
using SkiaSharp;
using Svg.Skia;

namespace MarkdownToPdf.Services
{
    /// <summary>
    /// 增强的高性能图片处理服务
    /// 支持WebP、SVG、JPEG等多种格式，优化图片布局和并排显示
    /// </summary>
    public class ImageProcessor : IDisposable
    {
        private bool _disposed = false;
        
        // 图片缓存：使用文件哈希作为键，避免重复处理
        private readonly ConcurrentDictionary<string, string> _imageCache = new();
        
        // 处理中的图片：避免并发处理同一图片
        private readonly ConcurrentDictionary<string, Task<string>> _processingTasks = new();
        
        // 缓存目录
        private readonly string _cacheDirectory;
        
        // 支持的图片格式
        private readonly HashSet<string> _supportedFormats = new(StringComparer.OrdinalIgnoreCase)
        {
            ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif", ".webp", ".svg"
        };
        
        // 进度报告
        public event EventHandler<ImageProcessingProgressEventArgs>? ProgressChanged;

        public ImageProcessor()
        {
            _cacheDirectory = Path.Combine(Path.GetTempPath(), "MarkdownToPdf", "ImageCache");
            Directory.CreateDirectory(_cacheDirectory);
            
            // 启动时清理过期缓存
            _ = Task.Run(CleanupExpiredCacheAsync);
        }

        /// <summary>
        /// 批量处理图片（高性能并行处理）
        /// 支持智能布局和并排显示
        /// </summary>
        /// <param name="imagePaths">图片路径列表</param>
        /// <param name="settings">转换设置</param>
        /// <returns>处理后的图片路径字典</returns>
        public async Task<Dictionary<string, string>> ProcessImagesAsync(IEnumerable<string> imagePaths, ConversionSettings settings)
        {
            var imageList = imagePaths.ToList();
            var results = new ConcurrentDictionary<string, string>();
            var totalImages = imageList.Count;
            var processedCount = 0;

            OnProgressChanged(new ImageProcessingProgressEventArgs(0, totalImages, "开始批量处理图片..."));

            if (totalImages == 0)
            {
                OnProgressChanged(new ImageProcessingProgressEventArgs(100, totalImages, "没有图片需要处理"));
                return results.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
            }

            // 智能图片分组，为并排显示做准备
            var imageGroups = GroupImagesForSideBySideDisplay(imageList, settings);

            // 根据系统配置和图片数量动态调整并行度
            var baseParallelism = settings.Performance?.MaxParallelImageProcessing ?? Environment.ProcessorCount * 2;
            var maxParallelism = Math.Max(baseParallelism, 8); // 至少8个并发
            
            if (totalImages > 50) maxParallelism = Math.Min(maxParallelism * 2, Environment.ProcessorCount * 6);
            if (totalImages > 200) maxParallelism = Math.Min(maxParallelism * 2, Environment.ProcessorCount * 8);

            // 分批处理以优化内存使用
            var baseBatchSize = settings.Performance?.ImageBatchSize ?? 20;
            var batchSize = Math.Max(baseBatchSize, maxParallelism * 2);
            var batches = imageList.Chunk(batchSize).ToList();

            for (int batchIndex = 0; batchIndex < batches.Count; batchIndex++)
            {
                var batch = batches[batchIndex];
                var batchProgress = (batchIndex * 100) / batches.Count;
                
                OnProgressChanged(new ImageProcessingProgressEventArgs(batchProgress, totalImages, 
                    $"处理第 {batchIndex + 1}/{batches.Count} 批图片..."));

                // 批内并行处理
                var parallelOptions = new ParallelOptions
                {
                    MaxDegreeOfParallelism = maxParallelism
                };

                await Parallel.ForEachAsync(batch, parallelOptions, async (imagePath, cancellationToken) =>
                {
                    try
                    {
                        var processedPath = await ProcessImageAsync(imagePath, settings);
                        results[imagePath] = processedPath;
                        
                        var completed = Interlocked.Increment(ref processedCount);
                        var progress = (int)(completed * 100.0 / totalImages);
                        
                        // 降低进度报告频率以减少UI更新开销
                        if (completed % Math.Max(1, totalImages / 50) == 0 || completed == totalImages)
                        {
                            OnProgressChanged(new ImageProcessingProgressEventArgs(progress, totalImages, 
                                $"已处理 {completed}/{totalImages} 张图片"));
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"处理图片失败 {imagePath}: {ex.Message}");
                        results[imagePath] = imagePath; // 失败时使用原路径
                    }
                });

                // 批处理完成后进行内存清理
                if (batchIndex % 3 == 0) // 每3批清理一次
                {
                    GC.Collect(0, GCCollectionMode.Optimized);
                }
            }

            OnProgressChanged(new ImageProcessingProgressEventArgs(100, totalImages, "图片处理完成"));
            return results.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
        }

        /// <summary>
        /// 智能分组图片，为并排显示做准备
        /// </summary>
        private List<List<string>> GroupImagesForSideBySideDisplay(List<string> imagePaths, ConversionSettings settings)
        {
            var groups = new List<List<string>>();
            var currentGroup = new List<string>();
            var maxGroupWidth = settings.MaxImageWidth * 0.8; // 留出一些边距
            var currentGroupWidth = 0.0;

            foreach (var imagePath in imagePaths)
            {
                try
                {
                    var imageInfo = GetImageInfo(imagePath);
                    if (imageInfo.Error != null) continue;

                    // 计算图片在页面中的显示宽度
                    var displayWidth = Math.Min(imageInfo.Width, maxGroupWidth / 2); // 最多两张并排

                    // 如果当前组为空或者加入此图片不会超出限制，则加入当前组
                    if (currentGroup.Count == 0 || 
                        (currentGroup.Count < 2 && currentGroupWidth + displayWidth <= maxGroupWidth))
                    {
                        currentGroup.Add(imagePath);
                        currentGroupWidth += displayWidth;
                    }
                    else
                    {
                        // 开始新组
                        if (currentGroup.Count > 0)
                        {
                            groups.Add(new List<string>(currentGroup));
                        }
                        currentGroup.Clear();
                        currentGroup.Add(imagePath);
                        currentGroupWidth = displayWidth;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"分析图片失败 {imagePath}: {ex.Message}");
                    // 出错的图片单独成组
                    if (currentGroup.Count > 0)
                    {
                        groups.Add(new List<string>(currentGroup));
                        currentGroup.Clear();
                        currentGroupWidth = 0;
                    }
                    groups.Add(new List<string> { imagePath });
                }
            }

            // 添加最后一组
            if (currentGroup.Count > 0)
            {
                groups.Add(currentGroup);
            }

            return groups;
        }

        /// <summary>
        /// 处理单个图片，使用缓存机制提高性能
        /// 支持多种格式包括WebP和SVG
        /// </summary>
        /// <param name="imagePath">图片路径</param>
        /// <param name="settings">转换设置</param>
        /// <returns>处理后的图片路径</returns>
        public async Task<string> ProcessImageAsync(string imagePath, ConversionSettings settings)
        {
            try
            {
                if (!File.Exists(imagePath))
                {
                    return imagePath; // 文件不存在时返回原路径
                }

                // 检查是否为支持的格式
                var extension = Path.GetExtension(imagePath).ToLowerInvariant();
                if (!_supportedFormats.Contains(extension))
                {
                    System.Diagnostics.Debug.WriteLine($"不支持的图片格式: {extension}");
                    return imagePath;
                }

                // 生成缓存键（基于文件内容和设置的哈希）
                var cacheKey = await GenerateCacheKeyAsync(imagePath, settings);
                
                // 检查缓存
                if (_imageCache.TryGetValue(cacheKey, out var cachedPath) && File.Exists(cachedPath))
                {
                    return cachedPath;
                }

                // 检查是否正在处理中
                if (_processingTasks.TryGetValue(cacheKey, out var existingTask))
                {
                    return await existingTask;
                }

                // 开始处理
                var processingTask = ProcessImageInternalAsync(imagePath, settings, cacheKey);
                _processingTasks[cacheKey] = processingTask;

                try
                {
                    var result = await processingTask;
                    _imageCache[cacheKey] = result;
                    return result;
                }
                finally
                {
                    _processingTasks.TryRemove(cacheKey, out _);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"处理图片失败 {imagePath}: {ex.Message}");
                return imagePath; // 处理失败时返回原路径
            }
        }

        /// <summary>
        /// 生成缓存键
        /// </summary>
        private Task<string> GenerateCacheKeyAsync(string imagePath, ConversionSettings settings)
        {
            return Task.Run(() =>
            {
                using var sha256 = SHA256.Create();
                
                // 组合文件信息和设置信息
                var fileInfo = new FileInfo(imagePath);
                var keyData = $"{imagePath}|{fileInfo.Length}|{fileInfo.LastWriteTime.Ticks}|" +
                             $"{settings.MaxImageWidth}|{settings.ImageQuality}|{settings.CompressImages}";
                
                var hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(keyData));
                return Convert.ToHexString(hashBytes);
            });
        }

        /// <summary>
        /// 内部图片处理方法（增强版，支持多种格式）
        /// </summary>
        private async Task<string> ProcessImageInternalAsync(string imagePath, ConversionSettings settings, string cacheKey)
        {
            return await Task.Run(() =>
            {
                var extension = Path.GetExtension(imagePath).ToLowerInvariant();
                var processedPath = Path.Combine(_cacheDirectory, $"{cacheKey}.jpg"); // 统一输出为JPEG

                // 快速检查是否需要处理
                if (!NeedsProcessing(imagePath, settings) && extension != ".svg" && extension != ".webp")
                {
                    return imagePath;
                }

                try
                {
                    if (extension == ".svg")
                    {
                        // SVG特殊处理
                        return ProcessSvgImage(imagePath, processedPath, settings);
                    }
                    else
                    {
                        // 使用ImageSharp处理其他格式（包括WebP）
                        return ProcessImageWithImageSharp(imagePath, processedPath, settings);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"图片处理失败 {imagePath}: {ex.Message}");
                    // 如果ImageSharp处理失败，尝试用System.Drawing
                    return ProcessImageWithSystemDrawing(imagePath, processedPath, settings);
                }
            });
        }

        /// <summary>
        /// 使用ImageSharp处理图片（支持WebP等现代格式）
        /// </summary>
        private string ProcessImageWithImageSharp(string inputPath, string outputPath, ConversionSettings settings)
        {
            using var image = SixLabors.ImageSharp.Image.Load(inputPath);
            
            // 计算新的尺寸
            var newSize = CalculateNewSize(new System.Drawing.Size(image.Width, image.Height), settings.MaxImageWidth);
            
            // 如果尺寸相同且不需要压缩，检查是否需要格式转换
            if (newSize.Width == image.Width && newSize.Height == image.Height && !settings.CompressImages)
            {
                var extension = Path.GetExtension(inputPath).ToLowerInvariant();
                if (extension == ".jpg" || extension == ".jpeg")
                {
                    return inputPath; // 不需要处理
                }
            }

            // 调整尺寸
            if (newSize.Width != image.Width || newSize.Height != image.Height)
            {
                image.Mutate(x => x.Resize(newSize.Width, newSize.Height));
            }

            // 保存为JPEG
            var quality = GetJpegQuality(settings.ImageQuality);
            var encoder = new JpegEncoder { Quality = quality };
            image.Save(outputPath, encoder);

            return outputPath;
        }

        /// <summary>
        /// 使用System.Drawing处理图片（后备方案）
        /// </summary>
        private string ProcessImageWithSystemDrawing(string inputPath, string outputPath, ConversionSettings settings)
        {
            try
            {
                using var originalImage = System.Drawing.Image.FromFile(inputPath);
                
                // 计算新的尺寸
                var newSize = CalculateNewSize(originalImage.Size, settings.MaxImageWidth);
                
                // 如果尺寸相同且不需要压缩，直接复制文件
                if (newSize.Width == originalImage.Width && 
                    newSize.Height == originalImage.Height && 
                    !settings.CompressImages)
                {
                    var extension = Path.GetExtension(inputPath).ToLowerInvariant();
                    if (extension == ".jpg" || extension == ".jpeg")
                    {
                        return inputPath;
                    }
                }

                // 使用优化的图片处理
                ProcessImageOptimized(originalImage, outputPath, newSize, settings);
                return outputPath;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"System.Drawing处理失败 {inputPath}: {ex.Message}");
                return inputPath;
            }
        }

        /// <summary>
        /// 处理SVG图片
        /// </summary>
        private string ProcessSvgImage(string inputPath, string outputPath, ConversionSettings settings)
        {
            try
            {
                // 使用Svg.Skia来真正渲染SVG
                using var svg = new SKSvg();
                
                // 加载SVG文件
                var svgPicture = svg.Load(inputPath);
                if (svgPicture == null)
                {
                    System.Diagnostics.Debug.WriteLine($"无法加载SVG文件: {inputPath}");
                    return CreateSvgFallback(inputPath, outputPath, settings);
                }

                // 获取SVG的原始尺寸
                var svgBounds = svgPicture.CullRect;
                var originalWidth = (int)svgBounds.Width;
                var originalHeight = (int)svgBounds.Height;
                
                // 如果尺寸为0，使用默认值
                if (originalWidth <= 0 || originalHeight <= 0)
                {
                    originalWidth = 800;
                    originalHeight = 600;
                }

                // 计算新的尺寸（应用50%缩放）
                var newSize = CalculateNewSize(new System.Drawing.Size(originalWidth, originalHeight), settings.MaxImageWidth);
                
                // 创建SkiaSharp表面来渲染
                var imageInfo = new SKImageInfo(newSize.Width, newSize.Height, SKColorType.Rgba8888, SKAlphaType.Premul);
                using var surface = SKSurface.Create(imageInfo);
                using var canvas = surface.Canvas;
                
                // 设置白色背景
                canvas.Clear(SKColors.White);
                
                // 计算缩放比例
                var scaleX = (float)newSize.Width / originalWidth;
                var scaleY = (float)newSize.Height / originalHeight;
                var scale = Math.Min(scaleX, scaleY); // 保持纵横比
                
                // 居中绘制
                canvas.Save();
                var offsetX = (newSize.Width - originalWidth * scale) / 2;
                var offsetY = (newSize.Height - originalHeight * scale) / 2;
                canvas.Translate(offsetX, offsetY);
                canvas.Scale(scale);
                
                // 绘制SVG
                canvas.DrawPicture(svgPicture);
                canvas.Restore();
                
                // 获取图像并保存
                using var image = surface.Snapshot();
                using var data = image.Encode(SKEncodedImageFormat.Png, GetJpegQuality(settings.ImageQuality));
                using var fileStream = File.OpenWrite(outputPath);
                data.SaveTo(fileStream);

                System.Diagnostics.Debug.WriteLine($"SVG成功转换: {inputPath} -> {outputPath} ({newSize.Width}x{newSize.Height})");
                return outputPath;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SVG处理失败 {inputPath}: {ex.Message}");
                return CreateSvgFallback(inputPath, outputPath, settings);
            }
        }

        /// <summary>
        /// 创建SVG回退图片（当SVG处理失败时）
        /// </summary>
        private string CreateSvgFallback(string inputPath, string outputPath, ConversionSettings settings)
        {
            try
            {
                // 读取SVG内容以尝试解析尺寸
                var svgContent = File.ReadAllText(inputPath);
                var (width, height) = ParseSvgDimensions(svgContent);
                
                if (width <= 0 || height <= 0)
                {
                    width = 800;
                    height = 600;
                }

                var newSize = CalculateNewSize(new System.Drawing.Size(width, height), settings.MaxImageWidth);

                // 创建回退图片
                using var bitmap = new Bitmap(newSize.Width, newSize.Height);
                using var graphics = Graphics.FromImage(bitmap);
                
                graphics.Clear(System.Drawing.Color.White);
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

                DrawSvgPlaceholder(graphics, newSize.Width, newSize.Height, Path.GetFileName(inputPath));

                var quality = GetJpegQuality(settings.ImageQuality);
                var encoderParameters = new EncoderParameters(1);
                encoderParameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);
                
                var jpegCodec = GetImageCodec(ImageFormat.Jpeg);
                bitmap.Save(outputPath, jpegCodec, encoderParameters);

                return outputPath;
            }
            catch
            {
                return inputPath;
            }
        }

        /// <summary>
        /// 解析SVG尺寸
        /// </summary>
        private (int width, int height) ParseSvgDimensions(string svgContent)
        {
            try
            {
                var widthMatch = Regex.Match(svgContent, @"width\s*=\s*[""']?(\d+)", RegexOptions.IgnoreCase);
                var heightMatch = Regex.Match(svgContent, @"height\s*=\s*[""']?(\d+)", RegexOptions.IgnoreCase);

                var width = widthMatch.Success ? int.Parse(widthMatch.Groups[1].Value) : 0;
                var height = heightMatch.Success ? int.Parse(heightMatch.Groups[1].Value) : 0;

                return (width, height);
            }
            catch
            {
                return (0, 0);
            }
        }

        /// <summary>
        /// 绘制SVG占位符
        /// </summary>
        private void DrawSvgPlaceholder(Graphics graphics, int width, int height, string fileName)
        {
            // 绘制背景
            graphics.FillRectangle(Brushes.WhiteSmoke, 0, 0, width, height);
            graphics.DrawRectangle(Pens.Gray, 0, 0, width - 1, height - 1);

            // 绘制SVG图标
            var iconSize = Math.Min(width, height) / 4;
            var iconX = (width - iconSize) / 2;
            var iconY = (height - iconSize) / 2 - 20;

            graphics.FillEllipse(Brushes.LightBlue, iconX, iconY, iconSize, iconSize);
            
            // 绘制文本
            var font = new Font("Microsoft YaHei", Math.Min(width / 20, 12));
            var textBrush = Brushes.DarkGray;
            var textRect = new System.Drawing.RectangleF(10, iconY + iconSize + 10, width - 20, 40);
            var stringFormat = new StringFormat 
            { 
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };
            
            graphics.DrawString($"SVG: {fileName}", font, textBrush, textRect, stringFormat);
        }

        /// <summary>
        /// 获取JPEG质量参数
        /// </summary>
        private int GetJpegQuality(ImageQuality quality)
        {
            return quality switch
            {
                ImageQuality.Low => 60,
                ImageQuality.Medium => 75,
                ImageQuality.High => 90,
                _ => 80
            };
        }

        /// <summary>
        /// 创建多图片合并的并排显示图片
        /// </summary>
        public Task<string> CreateSideBySideImageAsync(List<string> imagePaths, ConversionSettings settings)
        {
            return Task.Run(() =>
            {
                if (imagePaths.Count <= 1) 
                {
                    return imagePaths.FirstOrDefault() ?? "";
                }

            try
            {
                var cacheKey = string.Join("|", imagePaths) + $"|{settings.MaxImageWidth}|sidebyside";
                var hashedKey = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(cacheKey)));
                var outputPath = Path.Combine(_cacheDirectory, $"{hashedKey}_combined.jpg");

                if (File.Exists(outputPath))
                {
                    return outputPath;
                }

                // 加载所有图片
                var images = new List<SixLabors.ImageSharp.Image>();
                try
                {
                    foreach (var path in imagePaths)
                    {
                        if (File.Exists(path))
                        {
                            images.Add(SixLabors.ImageSharp.Image.Load(path));
                        }
                    }

                    if (images.Count == 0) return imagePaths.FirstOrDefault() ?? "";

                    // 计算合并后的尺寸
                    var maxHeight = images.Max(img => img.Height);
                    var totalWidth = images.Sum(img => img.Width);
                    var maxWidth = (int)settings.MaxImageWidth;

                    // 如果总宽度超过限制，按比例缩放
                    if (totalWidth > maxWidth)
                    {
                        var scale = (double)maxWidth / totalWidth;
                        maxHeight = (int)(maxHeight * scale);
                        for (int i = 0; i < images.Count; i++)
                        {
                            var newWidth = (int)(images[i].Width * scale);
                            var newHeight = (int)(images[i].Height * scale);
                            images[i].Mutate(x => x.Resize(newWidth, newHeight));
                        }
                        totalWidth = maxWidth;
                    }

                    // 创建合并图片
                    using var combinedImage = new SixLabors.ImageSharp.Image<SixLabors.ImageSharp.PixelFormats.Rgb24>(totalWidth, maxHeight);
                    combinedImage.Mutate(ctx => ctx.BackgroundColor(SixLabors.ImageSharp.Color.White));

                    // 将图片并排放置
                    int currentX = 0;
                    for (int i = 0; i < images.Count; i++)
                    {
                        var img = images[i];
                        var y = (maxHeight - img.Height) / 2; // 垂直居中
                        combinedImage.Mutate(ctx => ctx.DrawImage(img, new SixLabors.ImageSharp.Point(currentX, y), 1f));
                        currentX += img.Width;
                    }

                    // 保存合并后的图片
                    var quality = GetJpegQuality(settings.ImageQuality);
                    var encoder = new JpegEncoder { Quality = quality };
                    combinedImage.Save(outputPath, encoder);

                    return outputPath;
                }
                finally
                {
                    // 清理资源
                    foreach (var img in images)
                    {
                        img.Dispose();
                    }
                }
            }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"创建并排图片失败: {ex.Message}");
                    return imagePaths.FirstOrDefault() ?? "";
                }
            });
        }

        /// <summary>
        /// 优化的图片处理方法（保持原有System.Drawing逻辑作为后备）
        /// </summary>
        private void ProcessImageOptimized(System.Drawing.Image originalImage, string outputPath, System.Drawing.Size newSize, ConversionSettings settings)
        {
            // 对于小图片使用高质量模式，大图片使用快速模式
            var useHighQuality = originalImage.Width <= 1000 && originalImage.Height <= 1000;
            
            using var processedImage = new Bitmap(newSize.Width, newSize.Height, PixelFormat.Format24bppRgb);
            using var graphics = Graphics.FromImage(processedImage);
            
            // 根据图片大小选择不同的渲染质量
            if (useHighQuality)
            {
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
            }
            else
            {
                // 大图片使用快速模式
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.Low;
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighSpeed;
                graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighSpeed;
                graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighSpeed;
            }

            // 绘制调整后的图片
            graphics.DrawImage(originalImage, 0, 0, newSize.Width, newSize.Height);

            // 保存处理后的图片
            var imageFormat = GetImageFormat(Path.GetExtension(outputPath));
            var encoderParameters = GetEncoderParameters(settings.ImageQuality);
            
            if (encoderParameters != null && imageFormat == ImageFormat.Jpeg)
            {
                var codec = GetImageCodec(imageFormat);
                processedImage.Save(outputPath, codec, encoderParameters);
            }
            else
            {
                processedImage.Save(outputPath, imageFormat);
            }
        }

        /// <summary>
        /// 快速检查图片是否需要处理
        /// </summary>
        private bool NeedsProcessing(string imagePath, ConversionSettings settings)
        {
            try
            {
                using var image = System.Drawing.Image.FromFile(imagePath);
                
                // 检查是否需要调整尺寸
                if (image.Width > settings.MaxImageWidth)
                    return true;
                
                // 检查是否需要压缩
                if (settings.CompressImages)
                {
                    var fileInfo = new FileInfo(imagePath);
                    // 如果文件大于1MB，进行压缩
                    if (fileInfo.Length > 1024 * 1024)
                        return true;
                }
                
                // 检查是否需要格式转换
                var extension = Path.GetExtension(imagePath).ToLowerInvariant();
                if (extension is ".bmp" or ".tiff" or ".tif")
                    return true;
                
                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 计算新的图片尺寸，保持宽高比（默认缩小50%）
        /// </summary>
        private System.Drawing.Size CalculateNewSize(System.Drawing.Size originalSize, double maxWidth)
        {
            // 默认缩小50%
            var scaledWidth = originalSize.Width * 0.5;
            var scaledHeight = originalSize.Height * 0.5;
            
            // 检查缩小后是否仍超过最大宽度限制
            if (scaledWidth <= maxWidth)
            {
                return new System.Drawing.Size((int)scaledWidth, (int)scaledHeight);
            }

            // 如果缩小50%后仍超过限制，再根据最大宽度调整
            var ratio = maxWidth / scaledWidth;
            var finalWidth = (int)(scaledWidth * ratio);
            var finalHeight = (int)(scaledHeight * ratio);

            return new System.Drawing.Size(finalWidth, finalHeight);
        }

        /// <summary>
        /// 获取优化后的文件扩展名
        /// </summary>
        private string GetOptimizedExtension(string originalPath)
        {
            var extension = Path.GetExtension(originalPath).ToLowerInvariant();
            
            // 将某些格式转换为更适合的格式
            return extension switch
            {
                ".bmp" or ".tiff" or ".tif" => ".jpg", // 转换为JPEG以减小文件大小
                ".gif" => ".png", // 保持透明度
                _ => extension
            };
        }

        /// <summary>
        /// 获取图片格式
        /// </summary>
        private ImageFormat GetImageFormat(string extension)
        {
            return extension.ToLowerInvariant() switch
            {
                ".jpg" or ".jpeg" => ImageFormat.Jpeg,
                ".png" => ImageFormat.Png,
                ".gif" => ImageFormat.Gif,
                ".bmp" => ImageFormat.Bmp,
                ".tiff" or ".tif" => ImageFormat.Tiff,
                _ => ImageFormat.Jpeg
            };
        }

        /// <summary>
        /// 获取编码器参数（用于JPEG压缩）
        /// </summary>
        private EncoderParameters? GetEncoderParameters(ImageQuality quality)
        {
            var qualityValue = (int)quality;
            var qualityParam = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, qualityValue);
            return new EncoderParameters(1) { Param = { [0] = qualityParam } };
        }

        /// <summary>
        /// 获取图片编码器
        /// </summary>
        private ImageCodecInfo GetImageCodec(ImageFormat format)
        {
            var codecs = ImageCodecInfo.GetImageEncoders();
            foreach (var codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return codecs[0]; // 默认返回第一个编码器
        }

        /// <summary>
        /// 获取图片信息（尺寸、大小等）
        /// </summary>
        /// <param name="imagePath">图片路径</param>
        /// <returns>图片信息</returns>
        public ImageInfo GetImageInfo(string imagePath)
        {
            try
            {
                var fileInfo = new FileInfo(imagePath);
                using var image = System.Drawing.Image.FromFile(imagePath);
                return new ImageInfo
                {
                    Width = image.Width,
                    Height = image.Height,
                    FileSizeBytes = fileInfo.Length,
                    Format = image.RawFormat.ToString()
                };
            }
            catch (Exception ex)
            {
                return new ImageInfo 
                { 
                    Width = 0, 
                    Height = 0, 
                    FileSizeBytes = 0, 
                    Format = "Unknown",
                    Error = ex.Message 
                };
            }
        }

        /// <summary>
        /// 清理过期缓存
        /// </summary>
        private Task CleanupExpiredCacheAsync()
        {
            return Task.Run(() =>
            {
                try
                {
                    if (!Directory.Exists(_cacheDirectory))
                        return;

                    var cutoffTime = DateTime.Now.AddDays(-7); // 清理7天前的缓存
                    var files = Directory.GetFiles(_cacheDirectory);
                    
                    foreach (var file in files)
                    {
                        try
                        {
                            var fileInfo = new FileInfo(file);
                            if (fileInfo.LastAccessTime < cutoffTime)
                            {
                                File.Delete(file);
                            }
                        }
                        catch
                        {
                            // 忽略清理失败的文件
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"清理缓存失败: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 清理临时文件
        /// </summary>
        public void CleanupTempFiles()
        {
            try
            {
                var tempDir = Path.Combine(Path.GetTempPath(), "MarkdownToPdf");
                if (Directory.Exists(tempDir))
                {
                    Directory.Delete(tempDir, true);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"清理临时文件失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取缓存统计信息
        /// </summary>
        public CacheStatistics GetCacheStatistics()
        {
            var cacheCount = _imageCache.Count;
            var processingCount = _processingTasks.Count;
            
            long cacheSize = 0;
            if (Directory.Exists(_cacheDirectory))
            {
                var files = Directory.GetFiles(_cacheDirectory);
                cacheSize = files.Sum(file => new FileInfo(file).Length);
            }

            return new CacheStatistics
            {
                CachedItemCount = cacheCount,
                ProcessingItemCount = processingCount,
                CacheSizeBytes = cacheSize,
                CacheDirectory = _cacheDirectory
            };
        }

        /// <summary>
        /// 触发进度变化事件
        /// </summary>
        protected virtual void OnProgressChanged(ImageProcessingProgressEventArgs e)
        {
            ProgressChanged?.Invoke(this, e);
        }

        #region IDisposable 实现

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // 等待所有处理任务完成
                    Task.WaitAll(_processingTasks.Values.ToArray(), TimeSpan.FromSeconds(5));
                    
                    // 清理缓存
                    _imageCache.Clear();
                    _processingTasks.Clear();
                }
                _disposed = true;
            }
        }

        ~ImageProcessor()
        {
            Dispose(false);
        }

        #endregion
    }

    /// <summary>
    /// 图片信息
    /// </summary>
    public class ImageInfo
    {
        public int Width { get; set; }
        public int Height { get; set; }
        public long FileSizeBytes { get; set; }
        public string Format { get; set; } = string.Empty;
        public string? Error { get; set; }

        public string FileSizeFormatted => FormatFileSize(FileSizeBytes);
        public string DimensionsFormatted => $"{Width} × {Height}";

        private static string FormatFileSize(long bytes)
        {
            if (bytes < 1024) return $"{bytes} B";
            if (bytes < 1024 * 1024) return $"{bytes / 1024.0:F1} KB";
            if (bytes < 1024 * 1024 * 1024) return $"{bytes / (1024.0 * 1024):F1} MB";
            return $"{bytes / (1024.0 * 1024 * 1024):F1} GB";
        }
    }

    /// <summary>
    /// 图片处理进度事件参数
    /// </summary>
    public class ImageProcessingProgressEventArgs : EventArgs
    {
        public int Progress { get; }
        public int TotalImages { get; }
        public string Message { get; }

        public ImageProcessingProgressEventArgs(int progress, int totalImages, string message)
        {
            Progress = Math.Max(0, Math.Min(100, progress));
            TotalImages = totalImages;
            Message = message;
        }
    }

    /// <summary>
    /// 缓存统计信息
    /// </summary>
    public class CacheStatistics
    {
        public int CachedItemCount { get; set; }
        public int ProcessingItemCount { get; set; }
        public long CacheSizeBytes { get; set; }
        public string CacheDirectory { get; set; } = string.Empty;

        public string CacheSizeFormatted => FormatFileSize(CacheSizeBytes);

        private static string FormatFileSize(long bytes)
        {
            if (bytes < 1024) return $"{bytes} B";
            if (bytes < 1024 * 1024) return $"{bytes / 1024.0:F1} KB";
            if (bytes < 1024 * 1024 * 1024) return $"{bytes / (1024.0 * 1024):F1} MB";
            return $"{bytes / (1024.0 * 1024 * 1024):F1} GB";
        }
    }

    /// <summary>
    /// 扩展方法类
    /// </summary>
    public static class EnumerableExtensions
    {
        /// <summary>
        /// 将集合分割成指定大小的批次
        /// </summary>
        public static IEnumerable<T[]> Chunk<T>(this IEnumerable<T> source, int size)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (size <= 0) throw new ArgumentOutOfRangeException(nameof(size));

            return ChunkIterator(source, size);
        }

        private static IEnumerable<T[]> ChunkIterator<T>(IEnumerable<T> source, int size)
        {
            using var enumerator = source.GetEnumerator();
            while (enumerator.MoveNext())
            {
                var chunk = new T[size];
                chunk[0] = enumerator.Current;
                int i = 1;

                for (; i < size && enumerator.MoveNext(); i++)
                {
                    chunk[i] = enumerator.Current;
                }

                if (i == size)
                {
                    yield return chunk;
                }
                else
                {
                    var lastChunk = new T[i];
                    Array.Copy(chunk, lastChunk, i);
                    yield return lastChunk;
                }
            }
        }
    }
} 
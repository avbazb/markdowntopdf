using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace MarkdownToPdf.Models
{
    /// <summary>
    /// PDF转换设置模型
    /// 包含页面设置、字体配置、图片处理等选项
    /// </summary>
    public class ConversionSettings : INotifyPropertyChanged
    {
        #region 页面设置
        
        private PageSize _pageSize = PageSize.A4;
        /// <summary>页面大小</summary>
        public PageSize PageSize
        {
            get => _pageSize;
            set => SetProperty(ref _pageSize, value);
        }

        private double _marginTop = 20;
        /// <summary>上边距 (mm)</summary>
        public double MarginTop
        {
            get => _marginTop;
            set => SetProperty(ref _marginTop, Math.Max(0, value));
        }

        private double _marginBottom = 20;
        /// <summary>下边距 (mm)</summary>
        public double MarginBottom
        {
            get => _marginBottom;
            set => SetProperty(ref _marginBottom, Math.Max(0, value));
        }

        private double _marginLeft = 15;
        /// <summary>左边距 (mm)</summary>
        public double MarginLeft
        {
            get => _marginLeft;
            set => SetProperty(ref _marginLeft, Math.Max(0, value));
        }

        private double _marginRight = 15;
        /// <summary>右边距 (mm)</summary>
        public double MarginRight
        {
            get => _marginRight;
            set => SetProperty(ref _marginRight, Math.Max(0, value));
        }

        #endregion

        #region 字体设置

        private string _fontFamily = "Microsoft YaHei";
        /// <summary>字体族，默认使用微软雅黑支持中文</summary>
        public string FontFamily
        {
            get => _fontFamily;
            set => SetProperty(ref _fontFamily, value ?? "Microsoft YaHei");
        }

        private double _fontSize = 12;
        /// <summary>字体大小</summary>
        public double FontSize
        {
            get => _fontSize;
            set => SetProperty(ref _fontSize, Math.Max(8, Math.Min(72, value)));
        }

        private double _lineHeight = 1.5;
        /// <summary>行高倍数</summary>
        public double LineHeight
        {
            get => _lineHeight;
            set => SetProperty(ref _lineHeight, Math.Max(1.0, Math.Min(3.0, value)));
        }

        #endregion

        #region 图片设置

        private ImageQuality _imageQuality = ImageQuality.High;
        /// <summary>图片质量</summary>
        public ImageQuality ImageQuality
        {
            get => _imageQuality;
            set => SetProperty(ref _imageQuality, value);
        }

        private bool _compressImages = true;
        /// <summary>是否压缩图片以优化文件大小</summary>
        public bool CompressImages
        {
            get => _compressImages;
            set => SetProperty(ref _compressImages, value);
        }

        private double _maxImageWidth = 500;
        /// <summary>图片最大宽度 (px)</summary>
        public double MaxImageWidth
        {
            get => _maxImageWidth;
            set => SetProperty(ref _maxImageWidth, Math.Max(100, value));
        }

        #endregion

        #region 性能优化

        private bool _enableLargeFileOptimization = true;
        /// <summary>启用大文件优化</summary>
        public bool EnableLargeFileOptimization
        {
            get => _enableLargeFileOptimization;
            set => SetProperty(ref _enableLargeFileOptimization, value);
        }

        private int _chunkSize = 1000;
        /// <summary>分块处理大小 (行数)</summary>
        public int ChunkSize
        {
            get => _chunkSize;
            set => SetProperty(ref _chunkSize, Math.Max(100, Math.Min(10000, value)));
        }

        /// <summary>
        /// 高性能模式设置
        /// </summary>
        public PerformanceSettings Performance { get; set; } = new PerformanceSettings();

        #endregion

        #region INotifyPropertyChanged 实现

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        #endregion
    }

    /// <summary>页面大小枚举</summary>
    public enum PageSize
    {
        A4,
        A3,
        A5,
        Letter,
        Legal,
        Custom
    }

    /// <summary>
    /// 图片质量枚举
    /// </summary>
    public enum ImageQuality
    {
        Low = 60,
        Medium = 80,
        High = 95
    }

    /// <summary>
    /// 高性能模式设置
    /// </summary>
    public class PerformanceSettings : INotifyPropertyChanged
    {
        private int _maxParallelImageProcessing = Environment.ProcessorCount * 2;
        private int _maxParallelPdfProcessing = Environment.ProcessorCount;
        private int _imageBatchSize = 20;
        private int _memoryOptimizationLevel = 2;
        private bool _enableAggressiveCaching = true;
        private bool _enableMemoryOptimization = true;

        /// <summary>
        /// 图片处理最大并行度
        /// </summary>
        public int MaxParallelImageProcessing
        {
            get => _maxParallelImageProcessing;
            set => SetProperty(ref _maxParallelImageProcessing, Math.Max(1, Math.Min(Environment.ProcessorCount * 8, value)));
        }

        /// <summary>
        /// PDF处理最大并行度
        /// </summary>
        public int MaxParallelPdfProcessing
        {
            get => _maxParallelPdfProcessing;
            set => SetProperty(ref _maxParallelPdfProcessing, Math.Max(1, Math.Min(Environment.ProcessorCount * 4, value)));
        }

        /// <summary>
        /// 图片批处理大小
        /// </summary>
        public int ImageBatchSize
        {
            get => _imageBatchSize;
            set => SetProperty(ref _imageBatchSize, Math.Max(5, Math.Min(100, value)));
        }

        /// <summary>
        /// 内存优化级别 (1-5)
        /// </summary>
        public int MemoryOptimizationLevel
        {
            get => _memoryOptimizationLevel;
            set => SetProperty(ref _memoryOptimizationLevel, Math.Max(1, Math.Min(5, value)));
        }

        /// <summary>
        /// 启用激进缓存
        /// </summary>
        public bool EnableAggressiveCaching
        {
            get => _enableAggressiveCaching;
            set => SetProperty(ref _enableAggressiveCaching, value);
        }

        /// <summary>
        /// 启用内存优化
        /// </summary>
        public bool EnableMemoryOptimization
        {
            get => _enableMemoryOptimization;
            set => SetProperty(ref _enableMemoryOptimization, value);
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (!EqualityComparer<T>.Default.Equals(field, value))
            {
                field = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
} 
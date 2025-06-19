using System;
using System.Windows;
using System.Windows.Controls;

namespace MarkdownToPdf.Controls
{
    /// <summary>
    /// 进度对话框控件
    /// </summary>
    public partial class ProgressDialog : UserControl
    {
        public event EventHandler? CancelRequested;

        public ProgressDialog()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 更新进度
        /// </summary>
        public void UpdateProgress(int progress, string message)
        {
            ProgressBar.Value = progress;
            MessageTextBlock.Text = message;
        }

        /// <summary>
        /// 取消按钮点击事件
        /// </summary>
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            CancelRequested?.Invoke(this, EventArgs.Empty);
        }
    }
} 
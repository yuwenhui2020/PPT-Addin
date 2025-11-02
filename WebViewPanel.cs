using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using static System.ActivationContext;
//这是右边添加页面的侧栏
namespace PowerPointAddIn
{
    public partial class WebViewPanel : UserControl
    {
        public event Action<string, int, int> AddWebViewRequested;
        public WebViewPanel()
        {
            InitializeComponent();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            string url = txtUrl.Text;
            int width = (int)numWidth.Value;
            int height = (int)numHeight.Value;
            // 如果勾选了相对路径且是file协议
            if (chkRelativePath.Checked && url.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
            {
                // 提取文件名
                string fileName = Path.GetFileName(new Uri(url).LocalPath);
                // 获取PPT所在目录
                string pptPath = Globals.ThisAddIn.Application.ActivePresentation.Path;
                // 构建相对路径的完整路径
                string fullPath = Path.Combine(pptPath, fileName);
                // 转换为file URI
                url = new Uri(fullPath).AbsoluteUri;
            }
            AddWebViewRequested?.Invoke(url, width, height);

        }
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "All Files (*.*)|*.*";
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // 使用绝对路径
                    txtUrl.Text = new Uri(openFileDialog.FileName).AbsoluteUri;
                }
            }
        }
        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }


    }
}
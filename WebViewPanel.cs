using QRCoder;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using static System.ActivationContext;

namespace PowerPointAddIn
{
    public partial class WebViewPanel : UserControl
    {
        public event Action<string, int, int> AddWebViewRequested;
        private bool _consoleEnabled = false;
        private string _consoleUrl = "";
        private static int _serverPort = 8888;//web控制台页面端口
        private static int _screenSharePort = 8887;//屏幕共享
        private static int _textInputPort = 8889;//控制输入
        public WebViewPanel()
        {
            InitializeComponent();
            UpdateConsoleStatus();
        }
        public void UpdateConsoleStatus(bool enabled = false, string url = "")
        {
            _consoleEnabled = enabled;
            _consoleUrl = url;

            lblConsoleStatus.Text = _consoleEnabled ?
                $"控制台已启用：您可直接点击\n {_consoleUrl}" :
                "控制台未启用";

            btnToggleConsole.Text = _consoleEnabled ?
                "关闭控制台" :
                "开启控制台";

            UpdateQRCode();
        }
        private void UpdateQRCode()
        {
            if (_consoleEnabled && !string.IsNullOrEmpty(_consoleUrl))
            {
                // 生成二维码
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(_consoleUrl, QRCodeGenerator.ECCLevel.Q);
                QRCode qrCode = new QRCode(qrCodeData);
                Bitmap qrCodeImage = qrCode.GetGraphic(5);

                qrCodePictureBox.Image = qrCodeImage;
                qrCodePictureBox.Visible = true;
            }
            else
            {
                qrCodePictureBox.Image = null;
                qrCodePictureBox.Visible = false;
            }
        }
        private void btnToggleConsole_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.ToggleWebConsole();
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

        private void lblConsoleStatus_Click(object sender, EventArgs e)
        {
            if (_consoleEnabled)
            {
            System.Diagnostics.Process.Start( _consoleUrl);
            }
        }

        private void qrCodePictureBox_Click(object sender, EventArgs e)
        {
            if (_consoleEnabled)
            {
                System.Diagnostics.Process.Start(_consoleUrl);
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            
        }

        private void label6_Click(object sender, EventArgs e)
        {
            // 配置防火墙
            FirewallHelper.AddFirewallRule(_serverPort);
            FirewallHelper.AddFirewallRule(_screenSharePort);
            FirewallHelper.AddFirewallRule(_textInputPort);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.ToggleQRCodePanel();
        }

        private void label8_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.ToggleSymbolPanel();
        }

        private void label9_Click(object sender, EventArgs e)
        {
            ContactForm contactForm = new ContactForm();
            contactForm.ShowDialog();
        }
    }
}
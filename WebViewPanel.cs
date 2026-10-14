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
        public WebViewPanel()
        {
            InitializeComponent();
            UpdateConsoleStatus();
        }
        public void UpdateConsoleStatus(bool enabled = false, string url = "")
        {
            _consoleEnabled = enabled;
            _consoleUrl = url;

            string codeLine = "";
            try
            {
                var code = WebControl.AccessCode;
                if (_consoleEnabled && !string.IsNullOrEmpty(code))
                    codeLine = $"\n\n本次访问码：{code}\n（地址已自动带上，一般无需手工输入）";
            }
            catch { }

            lblConsoleStatus.Text = _consoleEnabled ?
                $"控制台已启用：您可直接点击\n {_consoleUrl}{codeLine}" :
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
                try
                {
                    QRCodeGenerator qrGenerator = new QRCodeGenerator();
                    QRCodeData qrCodeData = qrGenerator.CreateQrCode(_consoleUrl, QRCodeGenerator.ECCLevel.Q);
                    QRCode qrCode = new QRCode(qrCodeData);
                    Bitmap qrCodeImage = qrCode.GetGraphic(5);

                    // 修复资源泄漏：重复开关控制台会不断生成新 Bitmap，
                    // 旧的 Image 若不先 Dispose，会累积 GDI 句柄最终导致界面异常。
                    var oldImage = qrCodePictureBox.Image;
                    qrCodePictureBox.Image = qrCodeImage;
                    if (oldImage != null && !ReferenceEquals(oldImage, qrCodeImage))
                    {
                        oldImage.Dispose();
                    }
                    qrCodePictureBox.Visible = true;
                }
                catch (Exception ex)
                {
                    qrCodePictureBox.Visible = false;
                    MessageBox.Show($"生成二维码失败: {ex.Message}");
                }
            }
            else
            {
                var oldImage = qrCodePictureBox.Image;
                qrCodePictureBox.Image = null;
                if (oldImage != null) oldImage.Dispose();
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
                // 关键修复：不再“当场”把相对路径解析成绝对路径存储（那等于存了绝对路径，
                // 移动 PPT 后会失效）。改为存储“相对于 PPT 所在目录”的相对路径串，
                // 由 ThisAddIn 在放映/显示时按“当前 PPT 目录”实时解析为绝对 file URI。
                string relativeUrl = Globals.ThisAddIn.MakeRelativeWebViewUrl(url);
                if (relativeUrl != null)
                {
                    url = relativeUrl;
                }
                else
                {
                    // 无法转相对：两种可能，都必须分别处理，绝不能落到同一句
                    // Path.Combine(pptPath, fileName) —— PPT 未保存时 pptPath 为空串，
                    // Path.Combine 会抛 ArgumentException 直接导致添加失败。
                    string pptPath = Globals.ThisAddIn.Application.ActivePresentation.Path;
                    string fileName = Path.GetFileName(new Uri(url).LocalPath);

                    if (string.IsNullOrEmpty(pptPath))
                    {
                        // PPT 尚未保存：没有目录可作为相对基准，
                        // 保留用户原本选择的“绝对路径”（而非拼接出一个无效路径）。
                        MessageBox.Show("相对路径功能需要先把 PPT 保存到磁盘。\n\n" +
                            "本次已保留您选择的绝对路径；保存 PPT 后请重新添加网页以启用相对路径。",
                            "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        // url 保持原样（绝对路径）
                    }
                    else
                    {
                        // 网页不在 PPT 目录内：按需求只能接受同目录文件，故退回绝对路径
                        MessageBox.Show("所选网页不在 PPT 所在目录内，无法使用相对路径。\n已改用绝对路径，移动后可能失效。",
                            "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        url = new Uri(Path.Combine(pptPath, fileName)).AbsoluteUri;
                    }
                }
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
            // 配置 URL 保留项 + 防火墙：统一走 WebControl 的“最多一次提权”流程，
            // 不再按端口逐个提权（原实现会弹 3 次 UAC）。
            if (WebControl.EnsureNetworkAccess())
            {
                MessageBox.Show("已配置网络访问（URL 保留项 / 防火墙）。", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
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
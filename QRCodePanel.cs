using System;
using System.Drawing;
using System.Windows.Forms;
using QRCoder;

namespace PowerPointAddIn
{
    public partial class QRCodePanel : UserControl
    {
        public event Action<Bitmap> AddQRCodeRequested;
        private Bitmap _currentQRCode;

        public QRCodePanel()
        {
            InitializeComponent();
            cmbType.SelectedIndex = 0; // 默认选择"文本"
            UpdateUIForMode();
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            GenerateQRCode();
        }

        private void btnAddToSlide_Click(object sender, EventArgs e)
        {
            if (_currentQRCode != null)
            {
                AddQRCodeRequested?.Invoke(_currentQRCode);
            }
            else
            {
                MessageBox.Show("请先生成二维码");
            }
        }

        private void GenerateQRCode()
        {
            string content = txtContent.Text;
            if (string.IsNullOrWhiteSpace(content))
            {
                MessageBox.Show("请输入内容");
                return;
            }

            try
            {
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                string formattedContent = content;

                if (cmbType.SelectedItem.ToString() == "WiFi")
                {
                    string password = txtWiFiPassword.Text;
                    if (string.IsNullOrWhiteSpace(password))
                    {
                        MessageBox.Show("请输入WiFi密码");
                        return;
                    }
                    formattedContent = $"WIFI:T:WPA;S:{content};P:{password};;";
                }

                QRCodeData qrCodeData = qrGenerator.CreateQrCode(formattedContent, QRCodeGenerator.ECCLevel.Q);
                QRCode qrCode = new QRCode(qrCodeData);
                _currentQRCode = qrCode.GetGraphic(5);

                picQRCode.Image = _currentQRCode;
                picQRCode.SizeMode = PictureBoxSizeMode.Zoom;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成二维码失败: {ex.Message}");
            }
        }

        private void cmbType_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateUIForMode();
        }

        private void UpdateUIForMode()
        {
            bool isWiFiMode = cmbType.SelectedItem.ToString() == "WiFi";

            // 更新标签文本
            lblContent.Text = isWiFiMode ? "WiFi名称：" : "输入内容：";

            // 显示/隐藏WiFi密码相关控件
            txtWiFiPassword.Visible = isWiFiMode;
            lblWiFiPassword.Visible = isWiFiMode;

            // 重置输入框
            if (!isWiFiMode)
            {
                txtWiFiPassword.Text = "";
            }
        }

        private void lblWiFiPassword_Click(object sender, EventArgs e)
        {

        }
    }
}
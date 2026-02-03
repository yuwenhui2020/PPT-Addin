namespace PowerPointAddIn
{
    partial class QRCodePanel
    {
        private System.ComponentModel.IContainer components = null;

        private System.Windows.Forms.TextBox txtContent;
        private System.Windows.Forms.ComboBox cmbType;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.PictureBox picQRCode;
        private System.Windows.Forms.Button btnAddToSlide;
        private System.Windows.Forms.Label lblContent;
        private System.Windows.Forms.TextBox txtWiFiPassword;
        private System.Windows.Forms.Label lblWiFiPassword;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.txtContent = new System.Windows.Forms.TextBox();
            this.cmbType = new System.Windows.Forms.ComboBox();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.picQRCode = new System.Windows.Forms.PictureBox();
            this.btnAddToSlide = new System.Windows.Forms.Button();
            this.lblContent = new System.Windows.Forms.Label();
            this.txtWiFiPassword = new System.Windows.Forms.TextBox();
            this.lblWiFiPassword = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.picQRCode)).BeginInit();
            this.SuspendLayout();
            // 
            // txtContent
            // 
            this.txtContent.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtContent.Location = new System.Drawing.Point(10, 148);
            this.txtContent.Name = "txtContent";
            this.txtContent.Size = new System.Drawing.Size(234, 21);
            this.txtContent.TabIndex = 0;
            // 
            // cmbType
            // 
            this.cmbType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbType.FormattingEnabled = true;
            this.cmbType.Items.AddRange(new object[] {"文本或网页","WiFi"});
            this.cmbType.Location = new System.Drawing.Point(10, 97);
            this.cmbType.Name = "cmbType";
            this.cmbType.Size = new System.Drawing.Size(99, 20);
            this.cmbType.TabIndex = 1;
            this.cmbType.SelectedIndexChanged += new System.EventHandler(this.cmbType_SelectedIndexChanged);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnGenerate.Location = new System.Drawing.Point(10, 233);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(234, 30);
            this.btnGenerate.TabIndex = 2;
            this.btnGenerate.Text = "生成二维码";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // picQRCode
            // 
            this.picQRCode.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.picQRCode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.picQRCode.Location = new System.Drawing.Point(10, 338);
            this.picQRCode.Name = "picQRCode";
            this.picQRCode.Size = new System.Drawing.Size(234, 270);
            this.picQRCode.TabIndex = 3;
            this.picQRCode.TabStop = false;
            // 
            // btnAddToSlide
            // 
            this.btnAddToSlide.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAddToSlide.Location = new System.Drawing.Point(10, 288);
            this.btnAddToSlide.Name = "btnAddToSlide";
            this.btnAddToSlide.Size = new System.Drawing.Size(234, 30);
            this.btnAddToSlide.TabIndex = 4;
            this.btnAddToSlide.Text = "添加到PPT";
            this.btnAddToSlide.UseVisualStyleBackColor = true;
            this.btnAddToSlide.Click += new System.EventHandler(this.btnAddToSlide_Click);
            // 
            // lblContent
            // 
            this.lblContent.AutoSize = true;
            this.lblContent.Location = new System.Drawing.Point(8, 126);
            this.lblContent.Name = "lblContent";
            this.lblContent.Size = new System.Drawing.Size(65, 12);
            this.lblContent.TabIndex = 5;
            this.lblContent.Text = "输入内容：";
            // 
            // txtWiFiPassword
            // 
            this.txtWiFiPassword.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtWiFiPassword.Location = new System.Drawing.Point(10, 195);
            this.txtWiFiPassword.Name = "txtWiFiPassword";
            this.txtWiFiPassword.Size = new System.Drawing.Size(234, 21);
            this.txtWiFiPassword.TabIndex = 6;
            this.txtWiFiPassword.Visible = false;
            // 
            // lblWiFiPassword
            // 
            this.lblWiFiPassword.AutoSize = true;
            this.lblWiFiPassword.Location = new System.Drawing.Point(8, 176);
            this.lblWiFiPassword.Name = "lblWiFiPassword";
            this.lblWiFiPassword.Size = new System.Drawing.Size(65, 12);
            this.lblWiFiPassword.TabIndex = 7;
            this.lblWiFiPassword.Text = "WiFi密码：";
            this.lblWiFiPassword.Visible = false;
            this.lblWiFiPassword.Click += new System.EventHandler(this.lblWiFiPassword_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(11, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(135, 20);
            this.label1.TabIndex = 8;
            this.label1.Text = "二维码生成器";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 75);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(101, 12);
            this.label2.TabIndex = 9;
            this.label2.Text = "请选择二维码类型";
            // 
            // QRCodePanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblWiFiPassword);
            this.Controls.Add(this.txtWiFiPassword);
            this.Controls.Add(this.lblContent);
            this.Controls.Add(this.btnAddToSlide);
            this.Controls.Add(this.picQRCode);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.cmbType);
            this.Controls.Add(this.txtContent);
            this.Name = "QRCodePanel";
            this.Size = new System.Drawing.Size(254, 625);
            ((System.ComponentModel.ISupportInitialize)(this.picQRCode)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}
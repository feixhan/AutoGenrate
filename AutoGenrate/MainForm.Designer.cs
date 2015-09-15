namespace AutoGenrate
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.label1 = new System.Windows.Forms.Label();
            this.sourceFileTexBox = new System.Windows.Forms.TextBox();
            this.funcListBox = new System.Windows.Forms.ListBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.aboutRightMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.panel1 = new System.Windows.Forms.Panel();
            this.genBtn = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.openBtn = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.contextMenuStrip1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "源文件：";
            // 
            // sourceFileTexBox
            // 
            this.sourceFileTexBox.Location = new System.Drawing.Point(81, 20);
            this.sourceFileTexBox.Name = "sourceFileTexBox";
            this.sourceFileTexBox.Size = new System.Drawing.Size(339, 21);
            this.sourceFileTexBox.TabIndex = 1;
            // 
            // funcListBox
            // 
            this.funcListBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.funcListBox.FormattingEnabled = true;
            this.funcListBox.ItemHeight = 12;
            this.funcListBox.Location = new System.Drawing.Point(12, 67);
            this.funcListBox.Name = "funcListBox";
            this.funcListBox.Size = new System.Drawing.Size(534, 268);
            this.funcListBox.TabIndex = 4;
            this.funcListBox.SizeChanged += new System.EventHandler(this.funcListBox_SizeChanged);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutRightMenu});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(101, 26);
            // 
            // aboutRightMenu
            // 
            this.aboutRightMenu.Name = "aboutRightMenu";
            this.aboutRightMenu.Size = new System.Drawing.Size(100, 22);
            this.aboutRightMenu.Text = "关于";
            this.aboutRightMenu.Click += new System.EventHandler(this.aboutRightMenu_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.genBtn);
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Controls.Add(this.funcListBox);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(558, 372);
            this.panel1.TabIndex = 6;
            // 
            // genBtn
            // 
            this.genBtn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.genBtn.Location = new System.Drawing.Point(155, 341);
            this.genBtn.Name = "genBtn";
            this.genBtn.Size = new System.Drawing.Size(102, 23);
            this.genBtn.TabIndex = 5;
            this.genBtn.Text = "生成测试Code";
            this.genBtn.UseVisualStyleBackColor = true;
            this.genBtn.Click += new System.EventHandler(this.genBtn_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.openBtn);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.sourceFileTexBox);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(534, 49);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            // 
            // openBtn
            // 
            this.openBtn.Location = new System.Drawing.Point(443, 18);
            this.openBtn.Name = "openBtn";
            this.openBtn.Size = new System.Drawing.Size(75, 23);
            this.openBtn.TabIndex = 2;
            this.openBtn.Text = "打开";
            this.openBtn.UseVisualStyleBackColor = true;
            this.openBtn.Click += new System.EventHandler(this.openBtn_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(309, 341);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(110, 23);
            this.button1.TabIndex = 7;
            this.button1.Text = "生成函数列表xls";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(558, 372);
            this.ContextMenuStrip = this.contextMenuStrip1;
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.Text = "自动生成UnitTest";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.contextMenuStrip1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox sourceFileTexBox;
        private System.Windows.Forms.ListBox funcListBox;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem aboutRightMenu;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button openBtn;
        private System.Windows.Forms.Button genBtn;
        private System.Windows.Forms.Button button1;
    }
}


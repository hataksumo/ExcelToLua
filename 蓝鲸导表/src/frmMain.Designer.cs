namespace ExcelToLua
{
    partial class frmMain
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
            this.btnSele = new System.Windows.Forms.Button();
            this.btnOptWords = new System.Windows.Forms.Button();
            this.btnOptDesign = new System.Windows.Forms.Button();
            this.btnComoileLua = new System.Windows.Forms.Button();
            this.lblLoading = new System.Windows.Forms.Label();
            this.lblLoadDesc = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnSele
            // 
            this.btnSele.Location = new System.Drawing.Point(1, 1);
            this.btnSele.Name = "btnSele";
            this.btnSele.Size = new System.Drawing.Size(92, 42);
            this.btnSele.TabIndex = 1;
            this.btnSele.Text = "选择文件";
            this.btnSele.UseVisualStyleBackColor = true;
            this.btnSele.Click += new System.EventHandler(this.btnSele_Click);
            this.btnSele.KeyUp += new System.Windows.Forms.KeyEventHandler(this.btnSele_KeyUp);
            // 
            // btnOptWords
            // 
            this.btnOptWords.Location = new System.Drawing.Point(214, 1);
            this.btnOptWords.Name = "btnOptWords";
            this.btnOptWords.Size = new System.Drawing.Size(96, 42);
            this.btnOptWords.TabIndex = 7;
            this.btnOptWords.Text = "导出文本表";
            this.btnOptWords.UseVisualStyleBackColor = true;
            this.btnOptWords.Click += new System.EventHandler(this.btnCalELO_Click);
            // 
            // btnOptDesign
            // 
            this.btnOptDesign.Location = new System.Drawing.Point(214, 60);
            this.btnOptDesign.Name = "btnOptDesign";
            this.btnOptDesign.Size = new System.Drawing.Size(96, 35);
            this.btnOptDesign.TabIndex = 9;
            this.btnOptDesign.Text = "导出设计表";
            this.btnOptDesign.UseVisualStyleBackColor = true;
            this.btnOptDesign.Click += new System.EventHandler(this.btnOptDesign_Click);
            // 
            // btnComoileLua
            // 
            this.btnComoileLua.Location = new System.Drawing.Point(1, 60);
            this.btnComoileLua.Name = "btnComoileLua";
            this.btnComoileLua.Size = new System.Drawing.Size(104, 39);
            this.btnComoileLua.TabIndex = 10;
            this.btnComoileLua.Text = "编译LUA";
            this.btnComoileLua.UseVisualStyleBackColor = true;
            this.btnComoileLua.Click += new System.EventHandler(this.btnComoileLua_Click);
            // 
            // lblLoading
            // 
            this.lblLoading.AutoSize = true;
            this.lblLoading.Font = new System.Drawing.Font("微软雅黑", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblLoading.Location = new System.Drawing.Point(37, 4);
            this.lblLoading.Name = "lblLoading";
            this.lblLoading.Size = new System.Drawing.Size(262, 38);
            this.lblLoading.TabIndex = 11;
            this.lblLoading.Text = "正在加载配置表......";
            this.lblLoading.Visible = false;
            // 
            // lblLoadDesc
            // 
            this.lblLoadDesc.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblLoadDesc.Location = new System.Drawing.Point(12, 59);
            this.lblLoadDesc.Name = "lblLoadDesc";
            this.lblLoadDesc.Size = new System.Drawing.Size(308, 39);
            this.lblLoadDesc.TabIndex = 12;
            this.lblLoadDesc.Text = "加载开始...";
            this.lblLoadDesc.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblLoadDesc.Visible = false;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(332, 107);
            this.Controls.Add(this.lblLoadDesc);
            this.Controls.Add(this.lblLoading);
            this.Controls.Add(this.btnComoileLua);
            this.Controls.Add(this.btnOptDesign);
            this.Controls.Add(this.btnOptWords);
            this.Controls.Add(this.btnSele);
            this.Name = "frmMain";
            this.Text = "数值策划工具";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSele;
        private System.Windows.Forms.Button btnOptWords;
        private System.Windows.Forms.Button btnOptDesign;
        private System.Windows.Forms.Button btnComoileLua;
        private System.Windows.Forms.Label lblLoading;
        private System.Windows.Forms.Label lblLoadDesc;
    }
}


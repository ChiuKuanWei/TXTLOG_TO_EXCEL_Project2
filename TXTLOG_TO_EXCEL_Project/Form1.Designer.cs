namespace TXTLOG_TO_EXCEL_Project
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.dlg_MDB = new System.Windows.Forms.OpenFileDialog();
            this.txt_TXTPATH = new System.Windows.Forms.TextBox();
            this.lbl_TXT = new System.Windows.Forms.Label();
            this.dlg_TXT = new System.Windows.Forms.OpenFileDialog();
            this.btn_Report = new System.Windows.Forms.Button();
            this.check_Ripple = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // dlg_MDB
            // 
            this.dlg_MDB.FileName = "openFileDialog1";
            // 
            // txt_TXTPATH
            // 
            this.txt_TXTPATH.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_TXTPATH.Location = new System.Drawing.Point(102, 9);
            this.txt_TXTPATH.Multiline = true;
            this.txt_TXTPATH.Name = "txt_TXTPATH";
            this.txt_TXTPATH.ReadOnly = true;
            this.txt_TXTPATH.Size = new System.Drawing.Size(586, 72);
            this.txt_TXTPATH.TabIndex = 3;
            this.txt_TXTPATH.TabStop = false;
            // 
            // lbl_TXT
            // 
            this.lbl_TXT.AutoSize = true;
            this.lbl_TXT.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbl_TXT.Location = new System.Drawing.Point(12, 9);
            this.lbl_TXT.Name = "lbl_TXT";
            this.lbl_TXT.Size = new System.Drawing.Size(84, 16);
            this.lbl_TXT.TabIndex = 2;
            this.lbl_TXT.Text = "TXT檔選擇";
            this.lbl_TXT.Click += new System.EventHandler(this.lbl_TXT_Click);
            // 
            // dlg_TXT
            // 
            this.dlg_TXT.FileName = "openFileDialog1";
            // 
            // btn_Report
            // 
            this.btn_Report.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Report.Location = new System.Drawing.Point(711, 55);
            this.btn_Report.Name = "btn_Report";
            this.btn_Report.Size = new System.Drawing.Size(72, 26);
            this.btn_Report.TabIndex = 4;
            this.btn_Report.Text = "轉檔";
            this.btn_Report.UseVisualStyleBackColor = true;
            this.btn_Report.Click += new System.EventHandler(this.btn_Report_Click);
            // 
            // check_Ripple
            // 
            this.check_Ripple.AutoSize = true;
            this.check_Ripple.Location = new System.Drawing.Point(15, 40);
            this.check_Ripple.Name = "check_Ripple";
            this.check_Ripple.Size = new System.Drawing.Size(55, 16);
            this.check_Ripple.TabIndex = 5;
            this.check_Ripple.Text = "Ripple";
            this.check_Ripple.UseVisualStyleBackColor = true;
            this.check_Ripple.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 97);
            this.Controls.Add(this.check_Ripple);
            this.Controls.Add(this.btn_Report);
            this.Controls.Add(this.txt_TXTPATH);
            this.Controls.Add(this.lbl_TXT);
            this.Name = "Form1";
            this.Text = "For_R4K0DA01Q";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.OpenFileDialog dlg_MDB;
        private System.Windows.Forms.TextBox txt_TXTPATH;
        private System.Windows.Forms.Label lbl_TXT;
        private System.Windows.Forms.OpenFileDialog dlg_TXT;
        private System.Windows.Forms.Button btn_Report;
        private System.Windows.Forms.CheckBox check_Ripple;
    }
}


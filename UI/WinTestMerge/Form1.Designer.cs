namespace WinTestMerge
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
 
        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }
 
        #region Windows Form Designer generated code
 
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnDisconnect = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.lstFiles = new System.Windows.Forms.ListView();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.lstTarget = new System.Windows.Forms.ListView();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.lstDocxs = new System.Windows.Forms.ListView();
            this.lstLog = new System.Windows.Forms.ListView();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.btnTarget = new System.Windows.Forms.Button();
            this.btnOpen = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.btnDocx = new System.Windows.Forms.Button();
            this.tabD = new System.Windows.Forms.TabPage();
            this.lstXml = new System.Windows.Forms.ListView();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabD.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnDisconnect
            // 
            this.btnDisconnect.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.btnDisconnect.ForeColor = System.Drawing.Color.Maroon;
            this.btnDisconnect.Location = new System.Drawing.Point(105, 364);
            this.btnDisconnect.Name = "btnDisconnect";
            this.btnDisconnect.Size = new System.Drawing.Size(107, 45);
            this.btnDisconnect.TabIndex = 3;
            this.btnDisconnect.Text = "disconnect";
            this.btnDisconnect.UseVisualStyleBackColor = true;
            this.btnDisconnect.Click += new System.EventHandler(this.btnDisconnect_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabD);
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Location = new System.Drawing.Point(12, 39);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(549, 319);
            this.tabControl1.TabIndex = 4;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.lstFiles);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(541, 293);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Merge  Folder";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // lstFiles
            // 
            this.lstFiles.Location = new System.Drawing.Point(3, 6);
            this.lstFiles.MultiSelect = false;
            this.lstFiles.Name = "lstFiles";
            this.lstFiles.Size = new System.Drawing.Size(532, 284);
            this.lstFiles.TabIndex = 3;
            this.lstFiles.UseCompatibleStateImageBehavior = false;
            this.lstFiles.View = System.Windows.Forms.View.List;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.lstTarget);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(541, 293);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Disconnect Folder";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // lstTarget
            // 
            this.lstTarget.Location = new System.Drawing.Point(4, 6);
            this.lstTarget.MultiSelect = false;
            this.lstTarget.Name = "lstTarget";
            this.lstTarget.Size = new System.Drawing.Size(532, 281);
            this.lstTarget.TabIndex = 4;
            this.lstTarget.UseCompatibleStateImageBehavior = false;
            this.lstTarget.View = System.Windows.Forms.View.List;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.lstDocxs);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(541, 293);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Docx Folder";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // lstDocxs
            // 
            this.lstDocxs.Location = new System.Drawing.Point(4, 6);
            this.lstDocxs.MultiSelect = false;
            this.lstDocxs.Name = "lstDocxs";
            this.lstDocxs.Size = new System.Drawing.Size(532, 281);
            this.lstDocxs.TabIndex = 5;
            this.lstDocxs.UseCompatibleStateImageBehavior = false;
            this.lstDocxs.View = System.Windows.Forms.View.List;
            // 
            // lstLog
            // 
            this.lstLog.Location = new System.Drawing.Point(7, 453);
            this.lstLog.MultiSelect = false;
            this.lstLog.Name = "lstLog";
            this.lstLog.Size = new System.Drawing.Size(539, 114);
            this.lstLog.TabIndex = 6;
            this.lstLog.UseCompatibleStateImageBehavior = false;
            this.lstLog.View = System.Windows.Forms.View.List;
            // 
            // btnRefresh
            // 
            this.btnRefresh.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnRefresh.Location = new System.Drawing.Point(436, 10);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(79, 23);
            this.btnRefresh.TabIndex = 10;
            this.btnRefresh.Text = "refresh ";
            this.btnRefresh.UseVisualStyleBackColor = false;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click_1);
            // 
            // btnTarget
            // 
            this.btnTarget.Location = new System.Drawing.Point(152, 10);
            this.btnTarget.Name = "btnTarget";
            this.btnTarget.Size = new System.Drawing.Size(131, 23);
            this.btnTarget.TabIndex = 9;
            this.btnTarget.Text = "open target disconnect";
            this.btnTarget.UseVisualStyleBackColor = true;
            this.btnTarget.Click += new System.EventHandler(this.btnTarget_Click_1);
            // 
            // btnOpen
            // 
            this.btnOpen.Location = new System.Drawing.Point(7, 10);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(134, 23);
            this.btnOpen.TabIndex = 8;
            this.btnOpen.Text = "open source merge";
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click_1);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(218, 364);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(100, 45);
            this.button1.TabIndex = 11;
            this.button1.Text = "docm to docx";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnDocx
            // 
            this.btnDocx.Location = new System.Drawing.Point(289, 10);
            this.btnDocx.Name = "btnDocx";
            this.btnDocx.Size = new System.Drawing.Size(141, 23);
            this.btnDocx.TabIndex = 12;
            this.btnDocx.Text = "open docx";
            this.btnDocx.UseVisualStyleBackColor = true;
            this.btnDocx.Click += new System.EventHandler(this.btnDocx_Click);
            // 
            // tabD
            // 
            this.tabD.Controls.Add(this.lstXml);
            this.tabD.Location = new System.Drawing.Point(4, 22);
            this.tabD.Name = "tabD";
            this.tabD.Size = new System.Drawing.Size(541, 293);
            this.tabD.TabIndex = 3;
            this.tabD.Text = "Xml Change Merge";
            this.tabD.UseVisualStyleBackColor = true;
            // 
            // lstXml
            // 
            this.lstXml.Location = new System.Drawing.Point(3, 3);
            this.lstXml.MultiSelect = false;
            this.lstXml.Name = "lstXml";
            this.lstXml.Size = new System.Drawing.Size(527, 279);
            this.lstXml.TabIndex = 5;
            this.lstXml.UseCompatibleStateImageBehavior = false;
            this.lstXml.View = System.Windows.Forms.View.List;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(16, 364);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(83, 45);
            this.button2.TabIndex = 13;
            this.button2.Text = "Merge";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(324, 364);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(138, 45);
            this.button3.TabIndex = 14;
            this.button3.Text = "Disconnect And Convert To Docx";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(573, 600);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.btnDocx);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.btnTarget);
            this.Controls.Add(this.btnOpen);
            this.Controls.Add(this.lstLog);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.btnDisconnect);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            this.tabD.ResumeLayout(false);
            this.ResumeLayout(false);

        }
 
        #endregion
 
        private System.Windows.Forms.Button btnDisconnect;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.ListView lstFiles;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.ListView lstTarget;
        private System.Windows.Forms.ListView lstLog;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.Button btnTarget;
        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.ListView lstDocxs;
        private System.Windows.Forms.Button btnDocx;
        private System.Windows.Forms.TabPage tabD;
        private System.Windows.Forms.ListView lstXml;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
    }
}
 
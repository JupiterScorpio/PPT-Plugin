
namespace PPTAddin
{
    partial class ExportSetting
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
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.lib1_opt = new System.Windows.Forms.RadioButton();
            this.lib2_opt = new System.Windows.Forms.RadioButton();
            this.btn_ok = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(29, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Object Name:";
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(123, 33);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(176, 21);
            this.textBox1.TabIndex = 1;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // lib1_opt
            // 
            this.lib1_opt.AutoSize = true;
            this.lib1_opt.Location = new System.Drawing.Point(32, 92);
            this.lib1_opt.Name = "lib1_opt";
            this.lib1_opt.Size = new System.Drawing.Size(62, 17);
            this.lib1_opt.TabIndex = 2;
            this.lib1_opt.TabStop = true;
            this.lib1_opt.Text = "Library1";
            this.lib1_opt.UseVisualStyleBackColor = true;
            this.lib1_opt.CheckedChanged += new System.EventHandler(this.lib1_opt_CheckedChanged);
            // 
            // lib2_opt
            // 
            this.lib2_opt.AutoSize = true;
            this.lib2_opt.Location = new System.Drawing.Point(134, 92);
            this.lib2_opt.Name = "lib2_opt";
            this.lib2_opt.Size = new System.Drawing.Size(62, 17);
            this.lib2_opt.TabIndex = 3;
            this.lib2_opt.TabStop = true;
            this.lib2_opt.Text = "Library2";
            this.lib2_opt.UseVisualStyleBackColor = true;
            this.lib2_opt.CheckedChanged += new System.EventHandler(this.lib2_opt_CheckedChanged);
            // 
            // btn_ok
            // 
            this.btn_ok.Enabled = false;
            this.btn_ok.Location = new System.Drawing.Point(277, 85);
            this.btn_ok.Name = "btn_ok";
            this.btn_ok.Size = new System.Drawing.Size(63, 39);
            this.btn_ok.TabIndex = 4;
            this.btn_ok.Text = "OK";
            this.btn_ok.UseVisualStyleBackColor = true;
            this.btn_ok.Click += new System.EventHandler(this.btn_ok_Click);
            // 
            // ExportSetting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(359, 145);
            this.Controls.Add(this.btn_ok);
            this.Controls.Add(this.lib2_opt);
            this.Controls.Add(this.lib1_opt);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ExportSetting";
            this.Text = "ExportSetting";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.ExportSetting_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.RadioButton lib1_opt;
        private System.Windows.Forms.RadioButton lib2_opt;
        private System.Windows.Forms.Button btn_ok;
    }
}
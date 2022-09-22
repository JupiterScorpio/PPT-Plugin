
namespace PPTAddin
{
    partial class ShortkeySel
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
            this.ctrl_opt = new System.Windows.Forms.RadioButton();
            this.shift_opt = new System.Windows.Forms.RadioButton();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btn_cancel = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // ctrl_opt
            // 
            this.ctrl_opt.AutoSize = true;
            this.ctrl_opt.Location = new System.Drawing.Point(21, 16);
            this.ctrl_opt.Name = "ctrl_opt";
            this.ctrl_opt.Size = new System.Drawing.Size(73, 17);
            this.ctrl_opt.TabIndex = 0;
            this.ctrl_opt.TabStop = true;
            this.ctrl_opt.Text = "CapsLock";
            this.ctrl_opt.UseVisualStyleBackColor = true;
            this.ctrl_opt.CheckedChanged += new System.EventHandler(this.ctrl_opt_CheckedChanged);
            // 
            // shift_opt
            // 
            this.shift_opt.AutoSize = true;
            this.shift_opt.Location = new System.Drawing.Point(21, 58);
            this.shift_opt.Name = "shift_opt";
            this.shift_opt.Size = new System.Drawing.Size(46, 17);
            this.shift_opt.TabIndex = 1;
            this.shift_opt.TabStop = true;
            this.shift_opt.Text = "Shift";
            this.shift_opt.UseVisualStyleBackColor = true;
            this.shift_opt.CheckedChanged += new System.EventHandler(this.shift_opt_CheckedChanged_1);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(147, 26);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(59, 33);
            this.button1.TabIndex = 2;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.shift_opt);
            this.groupBox1.Controls.Add(this.ctrl_opt);
            this.groupBox1.Location = new System.Drawing.Point(20, 26);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(100, 98);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "ShortKey List";
            // 
            // btn_cancel
            // 
            this.btn_cancel.Location = new System.Drawing.Point(147, 84);
            this.btn_cancel.Name = "btn_cancel";
            this.btn_cancel.Size = new System.Drawing.Size(59, 33);
            this.btn_cancel.TabIndex = 4;
            this.btn_cancel.Text = "Cancel";
            this.btn_cancel.UseVisualStyleBackColor = true;
            this.btn_cancel.Click += new System.EventHandler(this.btn_cancel_Click);
            // 
            // ShortkeySel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(231, 138);
            this.Controls.Add(this.btn_cancel);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ShortkeySel";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Shortkey Selection";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.ShortkeySel_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RadioButton ctrl_opt;
        private System.Windows.Forms.RadioButton shift_opt;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btn_cancel;
    }
}
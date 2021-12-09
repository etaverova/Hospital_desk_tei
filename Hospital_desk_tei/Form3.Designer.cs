namespace Hospital_desk_tei
{
    partial class Form3
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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.domainUpDown1 = new System.Windows.Forms.DomainUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(476, 82);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(324, 31);
            this.textBox1.TabIndex = 9;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(79, 82);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(157, 25);
            this.label2.TabIndex = 8;
            this.label2.Text = "Номер палаты";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(479, 371);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(242, 64);
            this.button2.TabIndex = 7;
            this.button2.Text = "Отмена";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(178, 371);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(242, 64);
            this.button1.TabIndex = 6;
            this.button1.Text = " Сохранить";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // domainUpDown1
            // 
            this.domainUpDown1.Items.Add("1");
            this.domainUpDown1.Items.Add("2");
            this.domainUpDown1.Items.Add("3");
            this.domainUpDown1.Items.Add("4");
            this.domainUpDown1.Items.Add("5");
            this.domainUpDown1.Items.Add("6");
            this.domainUpDown1.Location = new System.Drawing.Point(479, 168);
            this.domainUpDown1.Name = "domainUpDown1";
            this.domainUpDown1.ReadOnly = true;
            this.domainUpDown1.Size = new System.Drawing.Size(82, 31);
            this.domainUpDown1.TabIndex = 10;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(79, 184);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(285, 25);
            this.label1.TabIndex = 11;
            this.label1.Text = "Вместимость палаты (чел.)";
            // 
            // Form3
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(892, 542);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.domainUpDown1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "Form3";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Палаты";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DomainUpDown domainUpDown1;
        private System.Windows.Forms.Label label1;
    }
}
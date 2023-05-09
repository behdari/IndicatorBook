namespace IndicatorBook.Sent
{
    partial class SentTemplateForm
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
            this.components = new System.ComponentModel.Container();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.textBoxAddress = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxFileName = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.radioButtonA4 = new System.Windows.Forms.RadioButton();
            this.radioButtonA5 = new System.Windows.Forms.RadioButton();
            this.label13 = new System.Windows.Forms.Label();
            this.buttonSubmit = new System.Windows.Forms.Button();
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.Controls.Add(this.textBoxAddress, 2, 2);
            this.tableLayoutPanel1.Controls.Add(this.label6, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.textBoxFileName, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.label3, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.label8, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 2, 1);
            this.tableLayoutPanel1.Controls.Add(this.label13, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.buttonSubmit, 2, 3);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 4;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11.11111F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11.11111F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11.11111F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11.11111F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(516, 218);
            this.tableLayoutPanel1.TabIndex = 7;
            // 
            // textBoxAddress
            // 
            this.textBoxAddress.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.textBoxAddress.Font = new System.Drawing.Font("B Nazanin", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.textBoxAddress.Location = new System.Drawing.Point(90, 119);
            this.textBoxAddress.Name = "textBoxAddress";
            this.textBoxAddress.ReadOnly = true;
            this.textBoxAddress.Size = new System.Drawing.Size(407, 31);
            this.textBoxAddress.TabIndex = 2;
            // 
            // label6
            // 
            this.label6.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label6.AutoSize = true;
            this.label6.BackColor = System.Drawing.Color.Transparent;
            this.label6.Location = new System.Drawing.Point(58, 19);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(10, 15);
            this.label6.TabIndex = 6;
            this.label6.Text = ":";
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("B Nazanin", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label1.Location = new System.Drawing.Point(3, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(49, 24);
            this.label1.TabIndex = 0;
            this.label1.Text = "نام فایل";
            // 
            // textBoxFileName
            // 
            this.textBoxFileName.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.textBoxFileName.Font = new System.Drawing.Font("B Nazanin", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.textBoxFileName.Location = new System.Drawing.Point(90, 11);
            this.textBoxFileName.Name = "textBoxFileName";
            this.textBoxFileName.Size = new System.Drawing.Size(407, 31);
            this.textBoxFileName.TabIndex = 0;
            this.textBoxFileName.TextChanged += new System.EventHandler(this.textBoxFileName_TextChanged);
            this.textBoxFileName.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBoxFileName_KeyDown);
            // 
            // label3
            // 
            this.label3.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("B Nazanin", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label3.Location = new System.Drawing.Point(8, 123);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(39, 24);
            this.label3.TabIndex = 2;
            this.label3.Text = "آدرس";
            // 
            // label8
            // 
            this.label8.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label8.AutoSize = true;
            this.label8.BackColor = System.Drawing.Color.Transparent;
            this.label8.Location = new System.Drawing.Point(58, 127);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(10, 15);
            this.label8.TabIndex = 8;
            this.label8.Text = ":";
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.BackColor = System.Drawing.Color.Transparent;
            this.tableLayoutPanel2.ColumnCount = 2;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 222F));
            this.tableLayoutPanel2.Controls.Add(this.radioButtonA4, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.radioButtonA5, 1, 0);
            this.tableLayoutPanel2.Location = new System.Drawing.Point(74, 57);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(440, 48);
            this.tableLayoutPanel2.TabIndex = 36;
            // 
            // radioButtonA4
            // 
            this.radioButtonA4.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.radioButtonA4.AutoSize = true;
            this.radioButtonA4.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButtonA4.Location = new System.Drawing.Point(86, 12);
            this.radioButtonA4.Name = "radioButtonA4";
            this.radioButtonA4.Size = new System.Drawing.Size(46, 23);
            this.radioButtonA4.TabIndex = 28;
            this.radioButtonA4.TabStop = true;
            this.radioButtonA4.Text = "A4";
            this.radioButtonA4.UseVisualStyleBackColor = true;
            // 
            // radioButtonA5
            // 
            this.radioButtonA5.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.radioButtonA5.AutoSize = true;
            this.radioButtonA5.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButtonA5.Location = new System.Drawing.Point(306, 12);
            this.radioButtonA5.Name = "radioButtonA5";
            this.radioButtonA5.Size = new System.Drawing.Size(46, 23);
            this.radioButtonA5.TabIndex = 27;
            this.radioButtonA5.TabStop = true;
            this.radioButtonA5.Text = "A5";
            this.radioButtonA5.UseVisualStyleBackColor = true;
            // 
            // label13
            // 
            this.label13.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label13.AutoSize = true;
            this.label13.BackColor = System.Drawing.Color.Transparent;
            this.label13.Font = new System.Drawing.Font("B Nazanin", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label13.Location = new System.Drawing.Point(11, 69);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(32, 24);
            this.label13.TabIndex = 20;
            this.label13.Text = "سایز";
            // 
            // buttonSubmit
            // 
            this.buttonSubmit.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.buttonSubmit.Font = new System.Drawing.Font("B Nazanin", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.buttonSubmit.Location = new System.Drawing.Point(250, 171);
            this.buttonSubmit.Name = "buttonSubmit";
            this.buttonSubmit.Size = new System.Drawing.Size(87, 37);
            this.buttonSubmit.TabIndex = 35;
            this.buttonSubmit.Text = "تایید";
            this.buttonSubmit.UseVisualStyleBackColor = true;
            this.buttonSubmit.Click += new System.EventHandler(this.buttonSubmit_Click);
            // 
            // errorProvider1
            // 
            this.errorProvider1.ContainerControl = this;
            // 
            // SentTemplateForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(516, 218);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "SentTemplateForm";
            this.Text = "قالب نامه ارسالی";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private TableLayoutPanel tableLayoutPanel1;
        public TextBox textBoxAddress;
        private Label label6;
        private Label label1;
        public TextBox textBoxFileName;
        private Label label3;
        private Label label8;
        private Button buttonSubmit;
        private TableLayoutPanel tableLayoutPanel2;
        public RadioButton radioButtonA4;
        public RadioButton radioButtonA5;
        private Label label13;
        private ErrorProvider errorProvider1;
    }
}
namespace MenuCreator
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.dabRadio = new System.Windows.Forms.RadioButton();
            this.cartRadio = new System.Windows.Forms.RadioButton();
            this.edibleRadio = new System.Windows.Forms.RadioButton();
            this.jointRadio = new System.Windows.Forms.RadioButton();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.txtConsole = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.noExtract = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.fDelay = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.yPos = new System.Windows.Forms.TextBox();
            this.xPos = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.button9 = new System.Windows.Forms.Button();
            this.radio_custom = new System.Windows.Forms.RadioButton();
            this.radio_1080 = new System.Windows.Forms.RadioButton();
            this.radio_4k = new System.Windows.Forms.RadioButton();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.cSizeW = new System.Windows.Forms.TextBox();
            this.cSizeH = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.excelRange1 = new System.Windows.Forms.TextBox();
            this.excelRange2 = new System.Windows.Forms.TextBox();
            this.button10 = new System.Windows.Forms.Button();
            this.button11 = new System.Windows.Forms.Button();
            this.button12 = new System.Windows.Forms.Button();
            this.button13 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButton1);
            this.groupBox1.Controls.Add(this.dabRadio);
            this.groupBox1.Controls.Add(this.cartRadio);
            this.groupBox1.Controls.Add(this.edibleRadio);
            this.groupBox1.Controls.Add(this.jointRadio);
            this.groupBox1.Location = new System.Drawing.Point(10, 11);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(107, 129);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Menu";
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Checked = true;
            this.radioButton1.Location = new System.Drawing.Point(4, 17);
            this.radioButton1.Margin = new System.Windows.Forms.Padding(2);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(56, 17);
            this.radioButton1.TabIndex = 4;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "Flower";
            this.radioButton1.UseVisualStyleBackColor = true;
            this.radioButton1.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // dabRadio
            // 
            this.dabRadio.AutoSize = true;
            this.dabRadio.Location = new System.Drawing.Point(4, 104);
            this.dabRadio.Margin = new System.Windows.Forms.Padding(2);
            this.dabRadio.Name = "dabRadio";
            this.dabRadio.Size = new System.Drawing.Size(88, 17);
            this.dabRadio.TabIndex = 3;
            this.dabRadio.Text = "Concentrates";
            this.dabRadio.UseVisualStyleBackColor = true;
            this.dabRadio.CheckedChanged += new System.EventHandler(this.dabRadio_CheckedChanged);
            // 
            // cartRadio
            // 
            this.cartRadio.AutoSize = true;
            this.cartRadio.Location = new System.Drawing.Point(4, 82);
            this.cartRadio.Margin = new System.Windows.Forms.Padding(2);
            this.cartRadio.Name = "cartRadio";
            this.cartRadio.Size = new System.Drawing.Size(72, 17);
            this.cartRadio.TabIndex = 2;
            this.cartRadio.Text = "Cartridges";
            this.cartRadio.UseVisualStyleBackColor = true;
            this.cartRadio.CheckedChanged += new System.EventHandler(this.cartRadio_CheckedChanged);
            // 
            // edibleRadio
            // 
            this.edibleRadio.AutoSize = true;
            this.edibleRadio.Location = new System.Drawing.Point(4, 60);
            this.edibleRadio.Margin = new System.Windows.Forms.Padding(2);
            this.edibleRadio.Name = "edibleRadio";
            this.edibleRadio.Size = new System.Drawing.Size(59, 17);
            this.edibleRadio.TabIndex = 1;
            this.edibleRadio.Text = "Edibles";
            this.edibleRadio.UseVisualStyleBackColor = true;
            this.edibleRadio.CheckedChanged += new System.EventHandler(this.edibleRadio_CheckedChanged);
            // 
            // jointRadio
            // 
            this.jointRadio.AutoSize = true;
            this.jointRadio.Location = new System.Drawing.Point(4, 38);
            this.jointRadio.Margin = new System.Windows.Forms.Padding(2);
            this.jointRadio.Name = "jointRadio";
            this.jointRadio.Size = new System.Drawing.Size(52, 17);
            this.jointRadio.TabIndex = 0;
            this.jointRadio.Text = "Joints";
            this.jointRadio.UseVisualStyleBackColor = true;
            this.jointRadio.CheckedChanged += new System.EventHandler(this.jointRadio_CheckedChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(10, 171);
            this.button1.Margin = new System.Windows.Forms.Padding(2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(108, 40);
            this.button1.TabIndex = 2;
            this.button1.Text = "Edit";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Enabled = false;
            this.button2.Location = new System.Drawing.Point(10, 231);
            this.button2.Margin = new System.Windows.Forms.Padding(2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(108, 40);
            this.button2.TabIndex = 3;
            this.button2.Text = "Create Image";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(10, 276);
            this.button3.Margin = new System.Windows.Forms.Padding(2);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(108, 40);
            this.button3.TabIndex = 4;
            this.button3.Text = "Preview";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(10, 320);
            this.button4.Margin = new System.Windows.Forms.Padding(2);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(108, 40);
            this.button4.TabIndex = 5;
            this.button4.Text = "Open WebPage";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.White;
            this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBox1.Location = new System.Drawing.Point(121, 11);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(960, 540);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 6;
            this.pictureBox1.TabStop = false;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(10, 517);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(106, 35);
            this.button5.TabIndex = 7;
            this.button5.Text = "Upload and run video";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(10, 558);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(106, 35);
            this.button6.TabIndex = 8;
            this.button6.Text = "Reboot TV";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(10, 476);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(106, 35);
            this.button7.TabIndex = 9;
            this.button7.Text = "End Video";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(10, 435);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(106, 35);
            this.button8.TabIndex = 10;
            this.button8.Text = "Create Video";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // txtConsole
            // 
            this.txtConsole.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.txtConsole.ForeColor = System.Drawing.SystemColors.ControlLight;
            this.txtConsole.Location = new System.Drawing.Point(1086, 11);
            this.txtConsole.Multiline = true;
            this.txtConsole.Name = "txtConsole";
            this.txtConsole.Size = new System.Drawing.Size(325, 653);
            this.txtConsole.TabIndex = 11;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.noExtract);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.fDelay);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.yPos);
            this.groupBox2.Controls.Add(this.xPos);
            this.groupBox2.Location = new System.Drawing.Point(122, 558);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(240, 106);
            this.groupBox2.TabIndex = 12;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Video Customization";
            // 
            // noExtract
            // 
            this.noExtract.AutoSize = true;
            this.noExtract.Location = new System.Drawing.Point(122, 71);
            this.noExtract.Name = "noExtract";
            this.noExtract.Size = new System.Drawing.Size(101, 17);
            this.noExtract.TabIndex = 7;
            this.noExtract.Text = "Dont Extract Gif";
            this.noExtract.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(119, 22);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(88, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Frame Delay (ms)";
            // 
            // fDelay
            // 
            this.fDelay.Location = new System.Drawing.Point(122, 42);
            this.fDelay.MaxLength = 4;
            this.fDelay.Name = "fDelay";
            this.fDelay.Size = new System.Drawing.Size(85, 20);
            this.fDelay.TabIndex = 5;
            this.fDelay.Text = "8";
            this.fDelay.TextChanged += new System.EventHandler(this.fDelay_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(28, 22);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(46, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Insert At";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 71);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(14, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Y";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 45);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(14, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "X";
            // 
            // yPos
            // 
            this.yPos.Location = new System.Drawing.Point(31, 68);
            this.yPos.MaxLength = 5;
            this.yPos.Name = "yPos";
            this.yPos.Size = new System.Drawing.Size(62, 20);
            this.yPos.TabIndex = 1;
            this.yPos.Text = "400";
            this.yPos.TextChanged += new System.EventHandler(this.yPos_TextChanged);
            // 
            // xPos
            // 
            this.xPos.Location = new System.Drawing.Point(31, 42);
            this.xPos.MaxLength = 5;
            this.xPos.Name = "xPos";
            this.xPos.Size = new System.Drawing.Size(62, 20);
            this.xPos.TabIndex = 0;
            this.xPos.Text = "175";
            this.xPos.TextChanged += new System.EventHandler(this.xPos_TextChanged);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.button9);
            this.groupBox3.Controls.Add(this.radio_custom);
            this.groupBox3.Controls.Add(this.radio_1080);
            this.groupBox3.Controls.Add(this.radio_4k);
            this.groupBox3.Controls.Add(this.label8);
            this.groupBox3.Controls.Add(this.label7);
            this.groupBox3.Controls.Add(this.cSizeW);
            this.groupBox3.Controls.Add(this.cSizeH);
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.excelRange1);
            this.groupBox3.Controls.Add(this.excelRange2);
            this.groupBox3.Enabled = false;
            this.groupBox3.Location = new System.Drawing.Point(368, 558);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(309, 106);
            this.groupBox3.TabIndex = 13;
            this.groupBox3.TabStop = false;
            // 
            // button9
            // 
            this.button9.Location = new System.Drawing.Point(6, 64);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(97, 23);
            this.button9.TabIndex = 11;
            this.button9.Text = "Save";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // radio_custom
            // 
            this.radio_custom.AutoSize = true;
            this.radio_custom.Location = new System.Drawing.Point(236, 64);
            this.radio_custom.Name = "radio_custom";
            this.radio_custom.Size = new System.Drawing.Size(60, 17);
            this.radio_custom.TabIndex = 10;
            this.radio_custom.Text = "Custom";
            this.radio_custom.UseVisualStyleBackColor = true;
            // 
            // radio_1080
            // 
            this.radio_1080.AutoSize = true;
            this.radio_1080.Location = new System.Drawing.Point(175, 64);
            this.radio_1080.Name = "radio_1080";
            this.radio_1080.Size = new System.Drawing.Size(55, 17);
            this.radio_1080.TabIndex = 9;
            this.radio_1080.Text = "1080p";
            this.radio_1080.UseVisualStyleBackColor = true;
            // 
            // radio_4k
            // 
            this.radio_4k.AutoSize = true;
            this.radio_4k.Checked = true;
            this.radio_4k.Location = new System.Drawing.Point(131, 64);
            this.radio_4k.Name = "radio_4k";
            this.radio_4k.Size = new System.Drawing.Size(38, 17);
            this.radio_4k.TabIndex = 8;
            this.radio_4k.TabStop = true;
            this.radio_4k.Text = "4K";
            this.radio_4k.UseVisualStyleBackColor = true;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(205, 40);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(15, 17);
            this.label8.TabIndex = 7;
            this.label8.Text = "x";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(132, 22);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(59, 13);
            this.label7.TabIndex = 6;
            this.label7.Text = "Image Size";
            // 
            // cSizeW
            // 
            this.cSizeW.Location = new System.Drawing.Point(131, 38);
            this.cSizeW.MaxLength = 4;
            this.cSizeW.Name = "cSizeW";
            this.cSizeW.Size = new System.Drawing.Size(68, 20);
            this.cSizeW.TabIndex = 5;
            this.cSizeW.Text = "1920";
            // 
            // cSizeH
            // 
            this.cSizeH.Location = new System.Drawing.Point(228, 39);
            this.cSizeH.MaxLength = 4;
            this.cSizeH.Name = "cSizeH";
            this.cSizeH.Size = new System.Drawing.Size(68, 20);
            this.cSizeH.TabIndex = 4;
            this.cSizeH.Text = "1080";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(6, 22);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(68, 13);
            this.label6.TabIndex = 3;
            this.label6.Text = "Excel Range";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(48, 38);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(13, 17);
            this.label5.TabIndex = 2;
            this.label5.Text = ":";
            // 
            // excelRange1
            // 
            this.excelRange1.Location = new System.Drawing.Point(6, 38);
            this.excelRange1.MaxLength = 4;
            this.excelRange1.Name = "excelRange1";
            this.excelRange1.Size = new System.Drawing.Size(36, 20);
            this.excelRange1.TabIndex = 1;
            this.excelRange1.Text = "A1";
            // 
            // excelRange2
            // 
            this.excelRange2.Location = new System.Drawing.Point(67, 38);
            this.excelRange2.MaxLength = 4;
            this.excelRange2.Name = "excelRange2";
            this.excelRange2.Size = new System.Drawing.Size(36, 20);
            this.excelRange2.TabIndex = 0;
            this.excelRange2.Text = "P69";
            // 
            // button10
            // 
            this.button10.Location = new System.Drawing.Point(11, 364);
            this.button10.Margin = new System.Windows.Forms.Padding(2);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(108, 40);
            this.button10.TabIndex = 14;
            this.button10.Text = "Print Menu As List";
            this.button10.UseVisualStyleBackColor = true;
            this.button10.Click += new System.EventHandler(this.button10_Click);
            // 
            // button11
            // 
            this.button11.Enabled = false;
            this.button11.Location = new System.Drawing.Point(682, 566);
            this.button11.Margin = new System.Windows.Forms.Padding(2);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(108, 40);
            this.button11.TabIndex = 15;
            this.button11.Text = "Exceless Editor";
            this.button11.UseVisualStyleBackColor = true;
            this.button11.Click += new System.EventHandler(this.button11_Click_1);
            // 
            // button12
            // 
            this.button12.Enabled = false;
            this.button12.Location = new System.Drawing.Point(682, 622);
            this.button12.Margin = new System.Windows.Forms.Padding(2);
            this.button12.Name = "button12";
            this.button12.Size = new System.Drawing.Size(108, 40);
            this.button12.TabIndex = 16;
            this.button12.Text = "Create File For Exceless Edits";
            this.button12.UseVisualStyleBackColor = true;
            this.button12.Click += new System.EventHandler(this.button12_Click);
            // 
            // button13
            // 
            this.button13.Location = new System.Drawing.Point(10, 615);
            this.button13.Margin = new System.Windows.Forms.Padding(2);
            this.button13.Name = "button13";
            this.button13.Size = new System.Drawing.Size(108, 40);
            this.button13.TabIndex = 17;
            this.button13.Text = "I Need Help";
            this.button13.UseVisualStyleBackColor = true;
            this.button13.Click += new System.EventHandler(this.button13_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1418, 676);
            this.Controls.Add(this.button13);
            this.Controls.Add(this.button12);
            this.Controls.Add(this.button11);
            this.Controls.Add(this.button10);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.txtConsole);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Menu Creator";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton dabRadio;
        private System.Windows.Forms.RadioButton cartRadio;
        private System.Windows.Forms.RadioButton edibleRadio;
        private System.Windows.Forms.RadioButton jointRadio;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.TextBox txtConsole;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox yPos;
        private System.Windows.Forms.TextBox xPos;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox fDelay;
        private System.Windows.Forms.CheckBox noExtract;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.RadioButton radio_custom;
        private System.Windows.Forms.RadioButton radio_1080;
        private System.Windows.Forms.RadioButton radio_4k;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox cSizeW;
        private System.Windows.Forms.TextBox cSizeH;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox excelRange1;
        private System.Windows.Forms.TextBox excelRange2;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button button10;
        private System.Windows.Forms.Button button11;
        private System.Windows.Forms.Button button12;
        private System.Windows.Forms.Button button13;
    }
}


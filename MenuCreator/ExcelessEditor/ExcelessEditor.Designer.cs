namespace MenuCreator.ExcelessEditor
{
    partial class ExcelessEditor
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
            this.companyGrid = new System.Windows.Forms.DataGridView();
            this.productGrid = new System.Windows.Forms.DataGridView();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.companyGrid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.productGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // companyGrid
            // 
            this.companyGrid.AllowUserToAddRows = false;
            this.companyGrid.AllowUserToDeleteRows = false;
            this.companyGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.companyGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.companyGrid.Location = new System.Drawing.Point(13, 32);
            this.companyGrid.MultiSelect = false;
            this.companyGrid.Name = "companyGrid";
            this.companyGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.companyGrid.Size = new System.Drawing.Size(309, 646);
            this.companyGrid.TabIndex = 0;
            this.companyGrid.SelectionChanged += new System.EventHandler(this.companyGrid_SelectionChanged);
            // 
            // productGrid
            // 
            this.productGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.productGrid.Location = new System.Drawing.Point(328, 32);
            this.productGrid.Name = "productGrid";
            this.productGrid.Size = new System.Drawing.Size(513, 646);
            this.productGrid.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(847, 446);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(179, 39);
            this.button1.TabIndex = 2;
            this.button1.Text = "Delete Product";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(847, 638);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(179, 39);
            this.button2.TabIndex = 3;
            this.button2.Text = "Delete Company";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(847, 489);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(179, 39);
            this.button3.TabIndex = 4;
            this.button3.Text = "Auto Re-Letter Product";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(847, 593);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(179, 39);
            this.button4.TabIndex = 5;
            this.button4.Text = "Change Company Number";
            this.button4.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Companys";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(325, 13);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(44, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Product";
            // 
            // button5
            // 
            this.button5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button5.Location = new System.Drawing.Point(847, 32);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(179, 39);
            this.button5.TabIndex = 8;
            this.button5.Text = "Save To Excel && Create Image";
            this.button5.UseVisualStyleBackColor = true;
            // 
            // button6
            // 
            this.button6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button6.Location = new System.Drawing.Point(847, 77);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(179, 39);
            this.button6.TabIndex = 9;
            this.button6.Text = "Save Products";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // ExcelessEditor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1110, 690);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.productGrid);
            this.Controls.Add(this.companyGrid);
            this.Name = "ExcelessEditor";
            this.Text = "Exceless Editor:: MENU";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Shown += new System.EventHandler(this.ExcelessEditor_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.companyGrid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.productGrid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView companyGrid;
        private System.Windows.Forms.DataGridView productGrid;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button6;
    }
}
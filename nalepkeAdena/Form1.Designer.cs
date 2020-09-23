namespace nalepkeAdena
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.potrditev = new System.Windows.Forms.Button();
            this.fileSystemWatcher1 = new System.IO.FileSystemWatcher();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.createBarcode = new System.Windows.Forms.Button();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.barcodeText = new System.Windows.Forms.TextBox();
            this.btnBedList = new System.Windows.Forms.Button();
            this.btnCreateLabelBed = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // potrditev
            // 
            this.potrditev.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.potrditev.Font = new System.Drawing.Font("Arial Narrow", 16F);
            this.potrditev.Location = new System.Drawing.Point(16, 105);
            this.potrditev.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.potrditev.Name = "potrditev";
            this.potrditev.Size = new System.Drawing.Size(187, 62);
            this.potrditev.TabIndex = 1;
            this.potrditev.Text = "Kreiraj nalepke";
            this.potrditev.UseVisualStyleBackColor = true;
            this.potrditev.Click += new System.EventHandler(this.potrditev_Click);
            // 
            // fileSystemWatcher1
            // 
            this.fileSystemWatcher1.EnableRaisingEvents = true;
            this.fileSystemWatcher1.SynchronizingObject = this;
            this.fileSystemWatcher1.Changed += new System.IO.FileSystemEventHandler(this.fileSystemWatcher1_Changed);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "Izberi datoteko:";
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // button1
            // 
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.button1.Font = new System.Drawing.Font("Arial Narrow", 16F);
            this.button1.Location = new System.Drawing.Point(16, 15);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(187, 62);
            this.button1.TabIndex = 0;
            this.button1.Text = "Izberi listo naročil";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(16, 218);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(187, 85);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 3;
            this.pictureBox1.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label1.Location = new System.Drawing.Point(211, 15);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(253, 312);
            this.label1.TabIndex = 2;
            this.label1.Text = resources.GetString("label1.Text");
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // createBarcode
            // 
            this.createBarcode.Location = new System.Drawing.Point(376, 444);
            this.createBarcode.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.createBarcode.Name = "createBarcode";
            this.createBarcode.Size = new System.Drawing.Size(277, 26);
            this.createBarcode.TabIndex = 4;
            this.createBarcode.Text = "Kreiraj neko coo";
            this.createBarcode.UseVisualStyleBackColor = true;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Location = new System.Drawing.Point(40, 343);
            this.pictureBox2.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(285, 127);
            this.pictureBox2.TabIndex = 5;
            this.pictureBox2.TabStop = false;
            // 
            // barcodeText
            // 
            this.barcodeText.Location = new System.Drawing.Point(415, 372);
            this.barcodeText.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.barcodeText.Name = "barcodeText";
            this.barcodeText.Size = new System.Drawing.Size(132, 22);
            this.barcodeText.TabIndex = 6;
            // 
            // btnBedList
            // 
            this.btnBedList.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.btnBedList.Location = new System.Drawing.Point(608, 15);
            this.btnBedList.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnBedList.Name = "btnBedList";
            this.btnBedList.Size = new System.Drawing.Size(207, 91);
            this.btnBedList.TabIndex = 8;
            this.btnBedList.Text = "Naloži listo za postelje";
            this.btnBedList.UseVisualStyleBackColor = true;
            this.btnBedList.Click += new System.EventHandler(this.btnBedList_Click);
            // 
            // btnCreateLabelBed
            // 
            this.btnCreateLabelBed.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.btnCreateLabelBed.Location = new System.Drawing.Point(608, 218);
            this.btnCreateLabelBed.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnCreateLabelBed.Name = "btnCreateLabelBed";
            this.btnCreateLabelBed.Size = new System.Drawing.Size(195, 106);
            this.btnCreateLabelBed.TabIndex = 10;
            this.btnCreateLabelBed.Text = "Kreiraj etike za postelje";
            this.btnCreateLabelBed.UseVisualStyleBackColor = true;
            this.btnCreateLabelBed.Click += new System.EventHandler(this.btnCreateLabelBed_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.button2.Location = new System.Drawing.Point(953, 26);
            this.button2.Margin = new System.Windows.Forms.Padding(4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(207, 91);
            this.button2.TabIndex = 11;
            this.button2.Text = "Naloži listo za okvirje postelje";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button5_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.button3.Location = new System.Drawing.Point(953, 226);
            this.button3.Margin = new System.Windows.Forms.Padding(4);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(207, 91);
            this.button3.TabIndex = 12;
            this.button3.Text = "Kreiraj etikete za okvirje";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1277, 577);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.btnCreateLabelBed);
            this.Controls.Add(this.btnBedList);
            this.Controls.Add(this.barcodeText);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.createBarcode);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.potrditev);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "Form1";
            this.Text = " ";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button potrditev;
        private System.IO.FileSystemWatcher fileSystemWatcher1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox barcodeText;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Button createBarcode;
        private System.Windows.Forms.Button btnBedList;
        private System.Windows.Forms.Button btnCreateLabelBed;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
    }
}


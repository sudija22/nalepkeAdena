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
            this.btnBedList = new System.Windows.Forms.Button();
            this.btnCreateLabelBed = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.specialSticker = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.checkBoxSmallSticker = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // potrditev
            // 
            this.potrditev.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.potrditev.Font = new System.Drawing.Font("Arial Narrow", 16F);
            this.potrditev.Location = new System.Drawing.Point(10, 146);
            this.potrditev.Name = "potrditev";
            this.potrditev.Size = new System.Drawing.Size(151, 58);
            this.potrditev.TabIndex = 1;
            this.potrditev.Text = "Kreiraj nalepke za okvirje";
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
            this.button1.Location = new System.Drawing.Point(10, 72);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(151, 68);
            this.button1.TabIndex = 0;
            this.button1.Text = "Izberi listo naročil okvirja";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(101, 370);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(284, 147);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 3;
            this.pictureBox1.TabStop = false;
            // 
            // btnBedList
            // 
            this.btnBedList.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.btnBedList.Location = new System.Drawing.Point(200, 72);
            this.btnBedList.Name = "btnBedList";
            this.btnBedList.Size = new System.Drawing.Size(147, 68);
            this.btnBedList.TabIndex = 8;
            this.btnBedList.Text = "Naloži listo za postelje";
            this.btnBedList.UseVisualStyleBackColor = true;
            this.btnBedList.Click += new System.EventHandler(this.btnBedList_Click);
            // 
            // btnCreateLabelBed
            // 
            this.btnCreateLabelBed.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.btnCreateLabelBed.Location = new System.Drawing.Point(200, 146);
            this.btnCreateLabelBed.Name = "btnCreateLabelBed";
            this.btnCreateLabelBed.Size = new System.Drawing.Size(147, 58);
            this.btnCreateLabelBed.TabIndex = 10;
            this.btnCreateLabelBed.Text = "Kreiraj etike za postelje";
            this.btnCreateLabelBed.UseVisualStyleBackColor = true;
            this.btnCreateLabelBed.Click += new System.EventHandler(this.btnCreateLabelBed_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.button2.Location = new System.Drawing.Point(366, 72);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(166, 68);
            this.button2.TabIndex = 11;
            this.button2.Text = "Naloži generalno listo";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button5_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.button3.Location = new System.Drawing.Point(366, 146);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(166, 58);
            this.button3.TabIndex = 12;
            this.button3.Text = "Kreiraj končno listo";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.potrditev_Click5);
            // 
            // comboBox1
            // 
            this.comboBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8"});
            this.comboBox1.Location = new System.Drawing.Point(325, 286);
            this.comboBox1.Margin = new System.Windows.Forms.Padding(2);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(92, 28);
            this.comboBox1.TabIndex = 13;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(44, 35);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(79, 20);
            this.label2.TabIndex = 15;
            this.label2.Text = "OKVIRJI";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(229, 35);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(98, 20);
            this.label3.TabIndex = 16;
            this.label3.Text = "POSTELJE";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(392, 35);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(119, 20);
            this.label4.TabIndex = 17;
            this.label4.Text = "GENERALNO";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(6, 289);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(301, 20);
            this.label5.TabIndex = 19;
            this.label5.Text = "Zaporedna številka kamijona v tednu";
            // 
            // specialSticker
            // 
            this.specialSticker.AutoSize = true;
            this.specialSticker.Location = new System.Drawing.Point(227, 221);
            this.specialSticker.Name = "specialSticker";
            this.specialSticker.Size = new System.Drawing.Size(137, 17);
            this.specialSticker.TabIndex = 20;
            this.specialSticker.Text = "SPECIALNE NALEPKE";
            this.specialSticker.UseVisualStyleBackColor = true;
            this.specialSticker.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(544, 35);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(202, 20);
            this.label1.TabIndex = 21;
            this.label1.Text = "MALE NALEPKE OKVIR";
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.button4.Location = new System.Drawing.Point(548, 72);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(166, 68);
            this.button4.TabIndex = 22;
            this.button4.Text = "Naloži listo male nalepke";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F);
            this.button5.Location = new System.Drawing.Point(548, 146);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(166, 58);
            this.button5.TabIndex = 23;
            this.button5.Text = "Kreiraj male nalepke";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click_1);
            // 
            // checkBoxSmallSticker
            // 
            this.checkBoxSmallSticker.AutoSize = true;
            this.checkBoxSmallSticker.Location = new System.Drawing.Point(575, 221);
            this.checkBoxSmallSticker.Name = "checkBoxSmallSticker";
            this.checkBoxSmallSticker.Size = new System.Drawing.Size(104, 17);
            this.checkBoxSmallSticker.TabIndex = 24;
            this.checkBoxSmallSticker.Text = "Končne nalepke";
            this.checkBoxSmallSticker.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(786, 629);
            this.Controls.Add(this.checkBoxSmallSticker);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.specialSticker);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.btnCreateLabelBed);
            this.Controls.Add(this.btnBedList);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.potrditev);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = " ";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button potrditev;
        private System.IO.FileSystemWatcher fileSystemWatcher1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button btnBedList;
        private System.Windows.Forms.Button btnCreateLabelBed;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.CheckBox specialSticker;
        private System.Windows.Forms.CheckBox checkBoxSmallSticker;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label1;
    }
}


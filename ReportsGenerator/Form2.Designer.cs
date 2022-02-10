using System.ComponentModel;

namespace ReportsGenerator
{
    partial class Form2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2));
            this.Project = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.Order = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.Block = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.Drawing = new System.Windows.Forms.TextBox();
            this.BrowseButton = new System.Windows.Forms.Button();
            this.GoButton = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.label5 = new System.Windows.Forms.Label();
            this.WorkFolder = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // Project
            // 
            this.Project.Location = new System.Drawing.Point(130, 9);
            this.Project.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Project.Name = "Project";
            this.Project.Size = new System.Drawing.Size(137, 23);
            this.Project.TabIndex = 0;
            this.Project.Text = "10510";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 15);
            this.label1.TabIndex = 1;
            this.label1.Text = "Проект";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 43);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(37, 15);
            this.label2.TabIndex = 3;
            this.label2.Text = "Заказ";
            // 
            // Order
            // 
            this.Order.Location = new System.Drawing.Point(130, 39);
            this.Order.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Order.Name = "Order";
            this.Order.Size = new System.Drawing.Size(137, 23);
            this.Order.TabIndex = 2;
            this.Order.Text = "056001";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 73);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(47, 15);
            this.label3.TabIndex = 5;
            this.label3.Text = "Секция";
            // 
            // Block
            // 
            this.Block.Location = new System.Drawing.Point(130, 69);
            this.Block.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Block.Name = "Block";
            this.Block.Size = new System.Drawing.Size(137, 23);
            this.Block.TabIndex = 4;
            this.Block.Text = "06001";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(13, 103);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(48, 15);
            this.label4.TabIndex = 7;
            this.label4.Text = "Чертёж";
            // 
            // Drawing
            // 
            this.Drawing.Location = new System.Drawing.Point(130, 99);
            this.Drawing.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Drawing.Name = "Drawing";
            this.Drawing.Size = new System.Drawing.Size(137, 23);
            this.Drawing.TabIndex = 6;
            this.Drawing.Text = "10510.362112.06001";
            // 
            // BrowseButton
            // 
            this.BrowseButton.Location = new System.Drawing.Point(14, 174);
            this.BrowseButton.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.BrowseButton.Name = "BrowseButton";
            this.BrowseButton.Size = new System.Drawing.Size(124, 27);
            this.BrowseButton.TabIndex = 8;
            this.BrowseButton.Text = "Выбрать папку...";
            this.BrowseButton.UseVisualStyleBackColor = true;
            this.BrowseButton.Click += new System.EventHandler(this.BrowseButton_Click);
            // 
            // GoButton
            // 
            this.GoButton.Location = new System.Drawing.Point(180, 174);
            this.GoButton.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.GoButton.Name = "GoButton";
            this.GoButton.Size = new System.Drawing.Size(88, 27);
            this.GoButton.TabIndex = 9;
            this.GoButton.Text = "Го";
            this.GoButton.UseVisualStyleBackColor = true;
            this.GoButton.Click += new System.EventHandler(this.GoButton_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(13, 133);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(88, 15);
            this.label5.TabIndex = 11;
            this.label5.Text = "Рабочая папка";
            // 
            // WorkFolder
            // 
            this.WorkFolder.Location = new System.Drawing.Point(130, 129);
            this.WorkFolder.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.WorkFolder.Name = "WorkFolder";
            this.WorkFolder.Size = new System.Drawing.Size(137, 23);
            this.WorkFolder.TabIndex = 10;
            this.WorkFolder.Text = "E:\\1";
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(281, 211);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.WorkFolder);
            this.Controls.Add(this.GoButton);
            this.Controls.Add(this.BrowseButton);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.Drawing);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.Block);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.Order);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Project);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form2";
            this.Text = "Генератор ведомостей";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox Project;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox Order;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox Block;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox Drawing;
        private System.Windows.Forms.Button BrowseButton;
        private System.Windows.Forms.Button GoButton;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox WorkFolder;
    }
}
using System.Windows.Forms;

namespace ReportsGenerator
{
	public sealed partial class Form1 : Form
	{
		protected override void Dispose(bool disposing)
		{
			try
			{
				if (disposing && this.components != null)
				{
					this.components.Dispose();
				}
			}
			finally
			{
				base.Dispose(disposing);
			}
		}
		
		private void InitializeComponent()
		{
			System.ComponentModel.ComponentResourceManager componentResourceManager = new System.ComponentModel.ComponentResourceManager(typeof(ReportsGenerator.Form1));
			this.MainBackgroundWorker = new System.ComponentModel.BackgroundWorker();
			this.LblProjectName = new System.Windows.Forms.Label();
			this.LblBlock = new System.Windows.Forms.Label();
			this.LblDrawName = new System.Windows.Forms.Label();
			this.TxbProject = new System.Windows.Forms.TextBox();
			this.TxbBlock = new System.Windows.Forms.TextBox();
			this.TxbDraw = new System.Windows.Forms.TextBox();
			this.TxbFolder = new System.Windows.Forms.TextBox();
			this.LblFolder = new System.Windows.Forms.Label();
			this.BtnFolder = new System.Windows.Forms.Button();
			this.BtnStart = new System.Windows.Forms.Button();
			this.FldrWorkDir = new System.Windows.Forms.FolderBrowserDialog();
			this.RichLogBox = new System.Windows.Forms.RichTextBox();
			this.MainMenuStrip = new System.Windows.Forms.MenuStrip();
			this.МенюToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.SettingsStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.ФайлПлотностейМатериаловToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            //this.HelpStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			//this.AboutSettingsStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.ToolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			this.ВыходToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.OpenDensityFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.MainMenuStrip.SuspendLayout();
			this.SuspendLayout();
			this.MainBackgroundWorker.WorkerReportsProgress = true;
			this.MainBackgroundWorker.WorkerSupportsCancellation = true;
			this.LblProjectName.AutoSize = true;
			System.Windows.Forms.Control lblProjectName = this.LblProjectName;
			System.Drawing.Point location = new System.Drawing.Point(48, 33);
			lblProjectName.Location = location;
			this.LblProjectName.Name = "LblProjectName";
			System.Windows.Forms.Control lblProjectName2 = this.LblProjectName;
			System.Drawing.Size size = new System.Drawing.Size(47, 13);
			lblProjectName2.Size = size;
			this.LblProjectName.TabIndex = 0;
			this.LblProjectName.Text = "Проект:";
			this.LblBlock.AutoSize = true;
			System.Windows.Forms.Control lblBlock = this.LblBlock;
			location = new System.Drawing.Point(12, 55);
			lblBlock.Location = location;
			this.LblBlock.Name = "LblBlock";
			System.Windows.Forms.Control lblBlock2 = this.LblBlock;
			size = new System.Drawing.Size(83, 13);
			lblBlock2.Size = size;
			this.LblBlock.TabIndex = 1;
			this.LblBlock.Text = "Номер секции:";
			this.LblDrawName.AutoSize = true;
			System.Windows.Forms.Control lblDrawName = this.LblDrawName;
			location = new System.Drawing.Point(6, 77);
			lblDrawName.Location = location;
			this.LblDrawName.Name = "LblDrawName";
			System.Windows.Forms.Control lblDrawName2 = this.LblDrawName;
			size = new System.Drawing.Size(89, 13);
			lblDrawName2.Size = size;
			this.LblDrawName.TabIndex = 2;
			this.LblDrawName.Text = "Номер чертежа:";
			this.TxbProject.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			System.Windows.Forms.Control txbProject = this.TxbProject;
			location = new System.Drawing.Point(101, 31);
			txbProject.Location = location;
			this.TxbProject.Name = "TxbProject";
			System.Windows.Forms.Control txbProject2 = this.TxbProject;
			size = new System.Drawing.Size(163, 20);
			txbProject2.Size = size;
			this.TxbProject.TabIndex = 4;
			this.TxbBlock.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			System.Windows.Forms.Control txbBlock = this.TxbBlock;
			location = new System.Drawing.Point(101, 53);
			txbBlock.Location = location;
			this.TxbBlock.Name = "TxbBlock";
			System.Windows.Forms.Control txbBlock2 = this.TxbBlock;
			size = new System.Drawing.Size(163, 20);
			txbBlock2.Size = size;
			this.TxbBlock.TabIndex = 5;
			this.TxbDraw.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			System.Windows.Forms.Control txbDraw = this.TxbDraw;
			location = new System.Drawing.Point(101, 75);
			txbDraw.Location = location;
			this.TxbDraw.Name = "TxbDraw";
			System.Windows.Forms.Control txbDraw2 = this.TxbDraw;
			size = new System.Drawing.Size(163, 20);
			txbDraw2.Size = size;
			this.TxbDraw.TabIndex = 6;
			this.TxbFolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			System.Windows.Forms.Control txbFolder = this.TxbFolder;
			location = new System.Drawing.Point(101, 97);
			txbFolder.Location = location;
			this.TxbFolder.Name = "TxbFolder";
			System.Windows.Forms.Control txbFolder2 = this.TxbFolder;
			size = new System.Drawing.Size(163, 20);
			txbFolder2.Size = size;
			this.TxbFolder.TabIndex = 8;
			this.LblFolder.AutoSize = true;
			System.Windows.Forms.Control lblFolder = this.LblFolder;
			location = new System.Drawing.Point(10, 99);
			lblFolder.Location = location;
			this.LblFolder.Name = "LblFolder";
			System.Windows.Forms.Control lblFolder2 = this.LblFolder;
			size = new System.Drawing.Size(85, 13);
			lblFolder2.Size = size;
			this.LblFolder.TabIndex = 9;
			this.LblFolder.Text = "Рабочая папка:";
			this.BtnFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			System.Windows.Forms.Control btnFolder = this.BtnFolder;
			location = new System.Drawing.Point(9, 120);
			btnFolder.Location = location;
			this.BtnFolder.Name = "BtnFolder";
			System.Windows.Forms.Control btnFolder2 = this.BtnFolder;
			size = new System.Drawing.Size(255, 23);
			btnFolder2.Size = size;
			this.BtnFolder.TabIndex = 10;
			this.BtnFolder.Text = "Выбрать папку...";
			this.BtnFolder.UseVisualStyleBackColor = true;
			this.BtnStart.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			System.Windows.Forms.Control btnStart = this.BtnStart;
			location = new System.Drawing.Point(9, 145);
			btnStart.Location = location;
			this.BtnStart.Name = "BtnStart";
			System.Windows.Forms.Control btnStart2 = this.BtnStart;
			size = new System.Drawing.Size(255, 23);
			btnStart2.Size = size;
			this.BtnStart.TabIndex = 11;
			this.BtnStart.Text = "Создать ведомости";
			this.BtnStart.UseVisualStyleBackColor = true;
			this.FldrWorkDir.RootFolder = System.Environment.SpecialFolder.DesktopDirectory;
			this.RichLogBox.BackColor = System.Drawing.SystemColors.Window;
			this.RichLogBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			System.Windows.Forms.Control richLogBox = this.RichLogBox;
			location = new System.Drawing.Point(270, 31);
			richLogBox.Location = location;
			this.RichLogBox.Name = "RichLogBox";
			System.Windows.Forms.Control richLogBox2 = this.RichLogBox;
			size = new System.Drawing.Size(259, 137);
			richLogBox2.Size = size;
			this.RichLogBox.TabIndex = 12;
			this.RichLogBox.Text = "";
			this.MainMenuStrip.BackColor = System.Drawing.SystemColors.Menu;
			this.MainMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[]
			{
				this.МенюToolStripMenuItem
			});
			System.Windows.Forms.Control mainMenuStrip = this.MainMenuStrip;
			location = new System.Drawing.Point(0, 0);
			mainMenuStrip.Location = location;
			this.MainMenuStrip.Name = "MainMenuStrip";
			System.Windows.Forms.Control mainMenuStrip2 = this.MainMenuStrip;
			size = new System.Drawing.Size(536, 24);
			mainMenuStrip2.Size = size;
			this.MainMenuStrip.TabIndex = 13;
			this.MainMenuStrip.Text = "MenuStrip1";
			this.МенюToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[]
			{
				this.SettingsStripMenuItem,
				//this.HelpStripMenuItem,
				//this.AboutSettingsStripMenuItem,
				this.ToolStripSeparator1,
				this.ВыходToolStripMenuItem
			});
			this.МенюToolStripMenuItem.Name = "МенюToolStripMenuItem";
			System.Windows.Forms.ToolStripItem менюToolStripMenuItem = this.МенюToolStripMenuItem;
			size = new System.Drawing.Size(53, 20);
			менюToolStripMenuItem.Size = size;
			this.МенюToolStripMenuItem.Text = "Меню";
			this.SettingsStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[]
			{
				this.ФайлПлотностейМатериаловToolStripMenuItem
			});
			this.SettingsStripMenuItem.Name = "SettingsStripMenuItem";
			System.Windows.Forms.ToolStripItem settingsStripMenuItem = this.SettingsStripMenuItem;
			size = new System.Drawing.Size(152, 22);
			settingsStripMenuItem.Size = size;
			this.SettingsStripMenuItem.Text = "Настройки";
			this.ФайлПлотностейМатериаловToolStripMenuItem.Name = "ФайлПлотностейМатериаловToolStripMenuItem";
			System.Windows.Forms.ToolStripItem файлПлотностейМатериаловToolStripMenuItem = this.ФайлПлотностейМатериаловToolStripMenuItem;
			size = new System.Drawing.Size(239, 22);
			файлПлотностейМатериаловToolStripMenuItem.Size = size;
			this.ФайлПлотностейМатериаловToolStripMenuItem.Text = "Файл плотностей материалов";
			//this.HelpStripMenuItem.Name = "HelpStripMenuItem";
			//System.Windows.Forms.ToolStripItem helpStripMenuItem = this.HelpStripMenuItem;
			size = new System.Drawing.Size(152, 22);
			//helpStripMenuItem.Size = size;
			//this.HelpStripMenuItem.Text = "Помощь";
			//this.AboutSettingsStripMenuItem.Name = "AboutSettingsStripMenuItem";
			//System.Windows.Forms.ToolStripItem aboutSettingsStripMenuItem = this.AboutSettingsStripMenuItem;
			size = new System.Drawing.Size(152, 22);
			//aboutSettingsStripMenuItem.Size = size;
			//this.AboutSettingsStripMenuItem.Text = "О программе";
			this.ToolStripSeparator1.Name = "ToolStripSeparator1";
			System.Windows.Forms.ToolStripItem toolStripSeparator = this.ToolStripSeparator1;
			size = new System.Drawing.Size(149, 6);
			toolStripSeparator.Size = size;
			this.ВыходToolStripMenuItem.Name = "ВыходToolStripMenuItem";
			System.Windows.Forms.ToolStripItem выходToolStripMenuItem = this.ВыходToolStripMenuItem;
			size = new System.Drawing.Size(152, 22);
			выходToolStripMenuItem.Size = size;
			this.ВыходToolStripMenuItem.Text = "Выход";
			this.OpenDensityFileDialog.Filter = "Файл плотностей материалов (*.def)|*.def";
			System.Drawing.SizeF autoScaleDimensions = new System.Drawing.SizeF(6f, 13f);
			this.AutoScaleDimensions = autoScaleDimensions;
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.SystemColors.Window;
			this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			size = new System.Drawing.Size(536, 175);
			this.ClientSize = size;
			this.Controls.Add(this.RichLogBox);
			this.Controls.Add(this.BtnStart);
			this.Controls.Add(this.BtnFolder);
			this.Controls.Add(this.LblFolder);
			this.Controls.Add(this.TxbFolder);
			this.Controls.Add(this.TxbDraw);
			this.Controls.Add(this.TxbBlock);
			this.Controls.Add(this.TxbProject);
			this.Controls.Add(this.LblDrawName);
			this.Controls.Add(this.LblBlock);
			this.Controls.Add(this.LblProjectName);
			this.Controls.Add(this.MainMenuStrip);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.Name = "Form1";
			this.Text = "Генератор ведомостей";
			this.MainMenuStrip.ResumeLayout(false);
			this.MainMenuStrip.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();
		}

		private System.ComponentModel.IContainer components;
	}
}

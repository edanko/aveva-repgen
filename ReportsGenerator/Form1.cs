using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using ReportsGenerator.My;

namespace ReportsGenerator;

public sealed partial class Form1
{
    //[AccessedThroughProperty("LblProjectName")]

    //[AccessedThroughProperty("LblBlock")]

    //[AccessedThroughProperty("LblDrawName")]

    //[AccessedThroughProperty("TxbProject")]

    //[AccessedThroughProperty("TxbBlock")]

    //[AccessedThroughProperty("TxbDraw")]

    //[AccessedThroughProperty("TxbFolder")]

    //[AccessedThroughProperty("LblFolder")]

    //[AccessedThroughProperty("BtnFolder")]
    private Button _btnFolder;

    //[AccessedThroughProperty("BtnStart")]
    private Button _btnStart;


    //[AccessedThroughProperty("MainBackgroundWorker")]
    private BackgroundWorker _mainBackgroundWorker;

    //[AccessedThroughProperty("FldrWorkDir")]

    //[AccessedThroughProperty("RichLogBox")]

    //[AccessedThroughProperty("MainMenuStrip")]

    //[AccessedThroughProperty("МенюToolStripMenuItem")]

    //[AccessedThroughProperty("SettingsStripMenuItem")]

    //[AccessedThroughProperty("HelpStripMenuItem")]
    //private ToolStripMenuItem _HelpStripMenuItem;

    //[AccessedThroughProperty("AboutSettingsStripMenuItem")]
    //private ToolStripMenuItem _AboutSettingsStripMenuItem;

    //[AccessedThroughProperty("ВыходToolStripMenuItem")]
    private ToolStripMenuItem _выходToolStripMenuItem;

    //[AccessedThroughProperty("ФайлПлотностейМатериаловToolStripMenuItem")]
    private ToolStripMenuItem _файлПлотностейМатериаловToolStripMenuItem;

    public Form1()
    {
        Load += Form1_Load;
        FormClosed += Form1_FormClosed;
        InitializeComponent();
    }

    internal BackgroundWorker MainBackgroundWorker
    {
        get => _mainBackgroundWorker;

        set
        {
            var value2 = new ProgressChangedEventHandler(MainBackgroundWorker_ProgressChanged);
            var value3 = new RunWorkerCompletedEventHandler(MainBackgroundWorker_RunWorkerCompleted);
            var value4 = new DoWorkEventHandler(MainBackgroundWorker_DoWork);
            if (_mainBackgroundWorker != null)
            {
                _mainBackgroundWorker.ProgressChanged -= value2;
                _mainBackgroundWorker.RunWorkerCompleted -= value3;
                _mainBackgroundWorker.DoWork -= value4;
            }

            _mainBackgroundWorker = value;
            if (_mainBackgroundWorker != null)
            {
                _mainBackgroundWorker.ProgressChanged += value2;
                _mainBackgroundWorker.RunWorkerCompleted += value3;
                _mainBackgroundWorker.DoWork += value4;
            }
        }
    }

    internal Label LblProjectName { get; set; }

    internal Label LblBlock { get; set; }

    internal Label LblDrawName { get; set; }

    internal TextBox TxbProject { get; set; }

    internal TextBox TxbBlock { get; set; }

    internal TextBox TxbDraw { get; set; }

    internal TextBox TxbFolder { get; set; }

    internal Label LblFolder { get; set; }

    internal Button BtnFolder
    {
        get => _btnFolder;

        set
        {
            var value2 = new EventHandler(BtnFolder_Click);
            if (_btnFolder != null) _btnFolder.Click -= value2;
            _btnFolder = value;
            if (_btnFolder != null) _btnFolder.Click += value2;
        }
    }

    internal Button BtnStart
    {
        get => _btnStart;

        set
        {
            var value2 = new EventHandler(BtnStart_Click);
            if (_btnStart != null) _btnStart.Click -= value2;
            _btnStart = value;
            if (_btnStart != null) _btnStart.Click += value2;
        }
    }

    internal FolderBrowserDialog FldrWorkDir { get; set; }

    internal RichTextBox RichLogBox { get; set; }

    internal new MenuStrip MainMenuStrip { get; set; }

    internal ToolStripMenuItem МенюToolStripMenuItem { get; set; }

    internal ToolStripMenuItem SettingsStripMenuItem { get; set; }


    internal ToolStripMenuItem ВыходToolStripMenuItem
    {
        get => _выходToolStripMenuItem;

        set
        {
            var value2 = new EventHandler(ВыходToolStripMenuItem_Click);
            if (_выходToolStripMenuItem != null) _выходToolStripMenuItem.Click -= value2;
            _выходToolStripMenuItem = value;
            if (_выходToolStripMenuItem != null) _выходToolStripMenuItem.Click += value2;
        }
    }

    internal ToolStripMenuItem ФайлПлотностейМатериаловToolStripMenuItem
    {
        get => _файлПлотностейМатериаловToolStripMenuItem;

        set
        {
            var value2 = new EventHandler(ФайлПлотностейМатериаловToolStripMenuItem_Click);
            if (_файлПлотностейМатериаловToolStripMenuItem != null)
                _файлПлотностейМатериаловToolStripMenuItem.Click -= value2;
            _файлПлотностейМатериаловToolStripMenuItem = value;
            if (_файлПлотностейМатериаловToolStripMenuItem != null)
                _файлПлотностейМатериаловToolStripMenuItem.Click += value2;
        }
    }

    internal OpenFileDialog OpenDensityFileDialog { get; set; }

    internal ToolStripSeparator ToolStripSeparator1 { get; set; }

    private void BtnStart_Click(object sender, EventArgs e)
    {
        if (Operators.CompareString(MySettingsProperty.Settings.QualityList, "", false) == 0)
        {
            MakeRed("Не выбран файл плотностей материалов (sbh_quality_list.def)\r\n");
            return;
        }

        if (Directory.Exists(TxbFolder.Text))
        {
            UpdateSettings();
            ChangeFormSize("min");
            MainBackgroundWorker.RunWorkerAsync();
            return;
        }

        MakeRed("Выбрана несуществующая папка\r\n");
    }

    private void Form1_Load(object sender, EventArgs e)
    {
        TxbProject.Text = MySettingsProperty.Settings.Project;
        TxbBlock.Text = MySettingsProperty.Settings.Block;
        TxbDraw.Text = MySettingsProperty.Settings.Draw;
        TxbFolder.Text = MySettingsProperty.Settings.WorkDir;
        RichLogBox.AppendText("--\r\n");
    }


    public void MakeRed(string s)
    {
        if (!Information.IsNothing(s))
        {
            RichLogBox.AppendText(s);
            RichLogBox.Select(checked(RichLogBox.TextLength - s.Length), s.Length);
            RichLogBox.SelectionColor = Color.Red;
        }
    }

    private void UpdateSettings()
    {
        MySettingsProperty.Settings.Project = TxbProject.Text;
        MySettingsProperty.Settings.Block = TxbBlock.Text;
        MySettingsProperty.Settings.Draw = TxbDraw.Text;
        MySettingsProperty.Settings.WorkDir = TxbFolder.Text;
        MySettingsProperty.Settings.Save();
    }

    private void Form1_FormClosed(object sender, FormClosedEventArgs e)
    {
        UpdateSettings();
    }

    private void BtnFolder_Click(object sender, EventArgs e)
    {
        FldrWorkDir.ShowDialog();
        var selectedPath = FldrWorkDir.SelectedPath;
        if (!Information.IsNothing(selectedPath) && Operators.CompareString(selectedPath, "", false) != 0)
            TxbFolder.Text = selectedPath;
    }

    private void MainBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
    {
        var backgroundWorker = (BackgroundWorker) sender;
        if (backgroundWorker.CancellationPending)
            e.Cancel = true;
        else
            DataProcessor.GenerateAll(backgroundWorker);
    }

    private void MainBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
        if (e.Cancelled)
        {
            RichLogBox.AppendText($"Принудительная отмена в {Conversions.ToString(DateTime.Now)}\r\n");
            Enabled = true;
        }
        else if (e.Error != null)
        {
            var s = $"Произошла ошибка: {e.Error.Message}";
            MakeRed(s);
            Enabled = true;
            ChangeFormSize("max");
        }
        else
        {
            Enabled = true;
            ChangeFormSize("max");
            RichLogBox.SaveFile($"{MySettingsProperty.Settings.WorkDir}\\logfile.rtf");
        }
    }

    private void MainBackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
    {
        if (Operators.ConditionalCompareObjectNotEqual(e.UserState, null, false))
        {
            if (Strings.InStr(e.UserState.ToString(), "!") != 0)
                MakeRed(e.UserState.ToString());
            else if (Strings.InStr(e.UserState.ToString(), "@") != 0)
                Text = $"{MySettingsProperty.Settings.FormTitle} {e.UserState}";
            else
                RichLogBox.AppendText(e.UserState.ToString());
        }
    }

    private void ФайлПлотностейМатериаловToolStripMenuItem_Click(object sender, EventArgs e)
    {
        OpenDensityFileDialog.ShowDialog();
        var fileName = OpenDensityFileDialog.FileName;
        if (!Information.IsNothing(fileName) && Operators.CompareString(fileName, "", false) != 0)
            MySettingsProperty.Settings.QualityList = fileName;
    }

    private void ChangeFormSize(string mode)
    {
        if (Operators.CompareString(mode, "max", false) == 0)
        {
            Width = 552;
            МенюToolStripMenuItem.Enabled = true;
            RichLogBox.SetBounds(270, 31, 259, 137);
        }
        else if (Operators.CompareString(mode, "min", false) == 0)
        {
            Width = 291;
            RichLogBox.SetBounds(9, 31, 259, 137);
            МенюToolStripMenuItem.Enabled = false;
        }
    }

    private void ВыходToolStripMenuItem_Click(object sender, EventArgs e)
    {
        Close();
    }

    //[AccessedThroughProperty("OpenDensityFileDialog")]

    //[AccessedThroughProperty("ToolStripSeparator1")]
}
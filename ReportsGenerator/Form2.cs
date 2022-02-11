using ReportsGenerator.Properties;

namespace ReportsGenerator
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void GoButton_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(WorkFolder.Text))
            {
                MessageBox.Show("Папка не существует.");
                return;
            }
            DataProcessor.GenerateAll();
        }

        private void BrowseWorkDirButton_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            var path = folderBrowserDialog1.SelectedPath;

            if (string.IsNullOrEmpty(path))
            {
                WorkFolder.Text = path;
            }
        }

        private void BrowseQualityListButton_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowHelp = true;
            openFileDialog1.FileName = "sbh_quality_list.def";
            openFileDialog1.Filter = "Def files (*.def)|*.def";
            openFileDialog1.ShowDialog();
            var path = openFileDialog1.FileName;

            if (string.IsNullOrEmpty(path))
            {
                QualityList.Text = path;
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            if (Settings.Default.HasSetDefaults)
            {
                Location = Settings.Default.Location;
                Size = Settings.Default.Size;
            }

            Project.DataBindings.Add("Text", Settings.Default, "Project", true,
                DataSourceUpdateMode.OnPropertyChanged);
            Order.DataBindings.Add("Text", Settings.Default, "Order", true,
                DataSourceUpdateMode.OnPropertyChanged);
            Block.DataBindings.Add("Text", Settings.Default, "Block", true,
                DataSourceUpdateMode.OnPropertyChanged);
            Drawing.DataBindings.Add("Text", Settings.Default, "Drawing", true,
                DataSourceUpdateMode.OnPropertyChanged);
            WorkFolder.DataBindings.Add("Text", Settings.Default, "WorkFolder", true,
                DataSourceUpdateMode.OnPropertyChanged);
            QualityList.DataBindings.Add("Text", Settings.Default, "QualityList", true,
                DataSourceUpdateMode.OnPropertyChanged);
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (WindowState == FormWindowState.Normal)
            {
                Settings.Default.Location = Location;
                Settings.Default.Size = Size;
            }
            else
            {
                Settings.Default.Location = RestoreBounds.Location;
                Settings.Default.Size = RestoreBounds.Size;
            }

            Settings.Default.HasSetDefaults = true;

            Settings.Default.Save();
        }
    }
}

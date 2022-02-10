using System.ComponentModel;

namespace ReportsGenerator
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void GoButton_Click(object sender, System.EventArgs e)
        {
            if (!Directory.Exists(WorkFolder.Text))
            {
                MessageBox.Show("Папка не существует.");
                return;
            }
            BackgroundWorker backgroundWorker = (BackgroundWorker)sender;
            DataProcessor.GenerateAll(backgroundWorker);
        }

        private void BrowseWorkDirButton_Click(object sender, System.EventArgs e)
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
            openFileDialog1.ShowDialog();
            var path = openFileDialog1.FileName;

            if (string.IsNullOrEmpty(path))
            {
                QualityList.Text = path;
            }
        }
    }
}

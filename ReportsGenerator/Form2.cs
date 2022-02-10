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

        private void BrowseButton_Click(object sender, System.EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            var path = folderBrowserDialog1.SelectedPath;

            if (string.IsNullOrEmpty(path))
            {
                WorkFolder.Text = path;
            }
        }
    }
}

using System;
using System.Windows;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
//using System.Windows.Forms;

namespace ExcelApplication1
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            backgroundWorker = ((BackgroundWorker)this.FindResource("backgroundWorker"));
            this.textBlock2.Text = "";
            this.textBlock3.Text = "";
            button1.IsEnabled = true;
            cmdCancel.IsEnabled = false;
        }

        private BackgroundWorker backgroundWorker;
        private string selectedPath = string.Empty;
        private List<ExcelFile> excelFileList = new List<ExcelFile>();
        private Settings objSettings = new Settings();

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            
            System.Windows.Forms.FolderBrowserDialog fb = new System.Windows.Forms.FolderBrowserDialog();
            if (fb.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //选择的文件夹路径
                selectedPath = fb.SelectedPath + "/";
                this.FindFile(selectedPath);
                this.textBlock2.Text = "你选择的目录为:" + fb.SelectedPath;
                this.textBlock3.Text = "文件数量为:" + excelFileList.Count;
                listBox1.ItemsSource = excelFileList;
            }
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            var selectFileNum = excelFileList.Count(c => c.IsSelected == true);
            if (selectFileNum <= 0)
            {
                MessageBox.Show("该目录下没有指定文件,请至少指定一个文件");
                return;
            }
            var objSettingsWindow = new SettingsWindow();
            objSettingsWindow.setSettings(this.objSettings);
            if (objSettingsWindow.ShowDialog().Value)
            {
                this.objSettings = objSettingsWindow.getSettings();
            }
            else
            {
                MessageBox.Show("请重新设置卡券信息");
                return;
            }

            // Disable the button
            button1.IsEnabled = false;
            cmdCancel.IsEnabled = true;
            this.progressBar1.Maximum = excelFileList.Count;
            // Start the process on another thread.
            backgroundWorker.RunWorkerAsync();            
        }

        //采用递归的方式遍历，文件夹和子文件中的所有文件。
        private void FindFile(string dirPath) //参数dirPath为指定的目录
        { 
            //在指定目录及子目录下查找文件
            DirectoryInfo Dir=new DirectoryInfo(dirPath);
            try
            {
                foreach(DirectoryInfo d in Dir.GetDirectories())//查找子目录 
                {
                    FindFile(d.FullName);
                }
                foreach(FileInfo f in Dir.GetFiles("*.xls")) //查找文件
                {
                    excelFileList.Add(new ExcelFile { Name = f.Name, Path = f.FullName, IsSelected = true, Status="未处理" });
                }
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }


        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            bool isOk = Worker.merge(selectedPath, excelFileList, this.objSettings, backgroundWorker);

            if (backgroundWorker.CancellationPending)
            {
                e.Cancel = true;
                return;
            }

            // Return the result.
            e.Result = isOk;            
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show("已取消.");
            }
            else if (e.Error != null)
            {
                // An error was thrown by the DoWork event handler.
                MessageBox.Show(e.Error.Message, "处理失败");
            }
            else
            {
                MessageBox.Show("处理结束");
            }
            button1.IsEnabled = true;
            cmdCancel.IsEnabled = false;
            progressBar1.Value = 0;
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        private void cmdCancel_Click(object sender, RoutedEventArgs e)
        {
            backgroundWorker.CancelAsync();
        }
    }
}

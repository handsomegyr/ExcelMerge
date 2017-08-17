using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ExcelApplication1
{
    /// <summary>
    /// Settings.xaml 的交互逻辑
    /// </summary>
    public partial class SettingsWindow : Window
    {
        protected Settings objSettings = new Settings();
        protected bool isCanceled = false;
        public SettingsWindow()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            this.isCanceled = false;
            this.Close();

        }

        public Settings getSettings()
        {
            return this.objSettings;
        }
        public void setSettings(Settings objSettings)
        {
             this.objSettings =objSettings ;
        }
        private void button2_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog fb = new System.Windows.Forms.FolderBrowserDialog();
            if (fb.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //选择的文件夹路径
                txtTargetPath.Text = fb.SelectedPath + "/";
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (this.isCanceled)
            {
                this.DialogResult = false;
                return;
            }
            this.DialogResult = false;
            e.Cancel = true;
            if (string.IsNullOrEmpty(this.txtActivityId.Text.Trim()))
            {
                MessageBox.Show("请输入活动ID");
                return;
            }
            if (string.IsNullOrEmpty(this.txtPrizeId.Text.Trim()))
            {
                MessageBox.Show("请输入奖品ID");
                return;
            }
            if (string.IsNullOrEmpty(this.txtPrjcode.Text.Trim()))
            {
                MessageBox.Show("请输入活动编号");
                return;
            }
            if (string.IsNullOrEmpty(this.txtTargetPath.Text.Trim()))
            {
                MessageBox.Show("请指定保存文件路径");
                return;
            }
            objSettings.ActivityId = this.txtActivityId.Text.Trim();
            objSettings.PrizeId = this.txtPrizeId.Text.Trim();
            objSettings.IsUsed = this.chkIsUsed.IsChecked ?? false;
            objSettings.Prjcode = this.txtPrjcode.Text.Trim();
            objSettings.TargetPath = this.txtTargetPath.Text.Trim();

            e.Cancel = false;
            this.DialogResult = true;
            
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.isCanceled = false;
            this.txtActivityId.Text = objSettings.ActivityId;
            this.txtPrizeId.Text = objSettings.PrizeId;
            this.chkIsUsed.IsChecked = objSettings.IsUsed ? true : false;
            this.txtPrjcode.Text = objSettings.Prjcode;
            this.txtTargetPath.Text = objSettings.TargetPath;
        }

        private void button3_Click(object sender, RoutedEventArgs e)
        {
            this.isCanceled = true;
            this.Close();
        }
    }
}

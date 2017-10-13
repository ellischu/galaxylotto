using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace GalaxyLotto
{
    /// <summary>
    /// Monitor.xaml 的互動邏輯
    /// </summary>
    public partial class Monitor : Window
    {
        public Monitor()
        {
            InitializeComponent();
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            this.Topmost = true;
        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            this.Topmost = false;
        }
        private void ProcedureMonitor_Closed(object sender, EventArgs e)
        {
            foreach (Window wintest in System.Windows.Application.Current.Windows)
            {
                if (wintest.Name == "NotifyIcon")
                {
                    WinNotifyIcon WinNotifyIcon00 = (WinNotifyIcon)wintest;
                    WinNotifyIcon00.notifyIcon00.ContextMenu.MenuItems["notifyIconProgressBar"].Checked = false;
                }
            }
        }
        private void ProcedureMonitor_Loaded(object sender, RoutedEventArgs e)
        {
            foreach (Window wintest in System.Windows.Application.Current.Windows)
            {
                if (wintest.Name == "NotifyIcon")
                {
                    WinNotifyIcon WinNotifyIcon00 = (WinNotifyIcon)wintest;
                    WinNotifyIcon00.notifyIcon00.ContextMenu.MenuItems["notifyIconProgressBar"].Checked = true;
                }
            }
        }
        private void ProcedureMonitor_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }
        private void Image_MouseEnter(object sender, MouseEventArgs e)
        {
            Image image = (Image)sender;
            image.Source = new BitmapImage(new Uri("pack://siteoforigin:,,,/Resources/btnclose01.png"));
        }
        private void Image_MouseLeave(object sender, MouseEventArgs e)
        {
            Image image = (Image)sender;
            image.Source = new BitmapImage(new Uri("pack://siteoforigin:,,,/Resources/btnclose00.png"));
        }
        private void Image_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.Close();
            }

        }

    }
}
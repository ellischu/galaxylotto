using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace GalaxyLotto
{
    /// <summary>
    /// NoticefyIcon.xaml 的互動邏輯
    /// </summary>
    public partial class WinNotifyIcon : Window
    {
        public NotifyIcon notifyIcon00;
        public WinNotifyIcon()
        {
            InitializeComponent();
            InitializeNotifyIcon();
        }

        private void InitializeNotifyIcon()
        {
            // Configure and show a notification icon in the system tray
            this.notifyIcon00 = new NotifyIcon()
            {
                BalloonTipText = "Hello, GalaxyLotto!",
                Text = "Hello, GalaxyLotto!",
                Icon = (System.Drawing.Icon)Properties.Resources.planet
            };
            this.notifyIcon00.DoubleClick += new EventHandler(NotifyIcon00_DoubleClick);
            //TODO:初始化 notifyIcon菜單
            System.Windows.Forms.ContextMenu notifyIconMenu = new System.Windows.Forms.ContextMenu();
            System.Windows.Forms.MenuItem notifyIconMenuItem = new System.Windows.Forms.MenuItem();

            notifyIconMenuItem = new System.Windows.Forms.MenuItem()
            {
                Name = "notifyIconMain",
                Index = 0,
                Checked = false,
                Text = "開啟主程式(&Main)"
            };
            notifyIconMenuItem.Click += new EventHandler(ChkshowMain_Click);
            notifyIconMenu.MenuItems.Add(notifyIconMenuItem);

            notifyIconMenuItem = new System.Windows.Forms.MenuItem()
            {
                Name = "notifyIconProgressBar",
                Index = 0,
                Checked = false,
                Text = "顯示進度(P&rocess)"
            };
            notifyIconMenuItem.Click += new EventHandler(ChkshowProcBar_Click);
            notifyIconMenu.MenuItems.Add(notifyIconMenuItem);

            notifyIconMenuItem = new System.Windows.Forms.MenuItem()
            {
                Name = "notifyIconSearch",
                Index = 0,
                Checked = false,
                Text = "搜尋(&Search)"
            };
            notifyIconMenuItem.Click += new EventHandler(NotifyIconMISearch_Click);
            notifyIconMenu.MenuItems.Add(notifyIconMenuItem);

            notifyIconMenuItem = new System.Windows.Forms.MenuItem()
            {
                Name = "notifyIconCreatePurple",
                Index = 0,
                Checked = false,
                Text = "紫微(&Purple)"
            };
            notifyIconMenuItem.Click += new EventHandler(NotifyIconMIPurple_Click);
            notifyIconMenu.MenuItems.Add(notifyIconMenuItem);

            notifyIconMenuItem = new System.Windows.Forms.MenuItem()
            {
                Index = 0,
                Text = "結束(E&xit)"
            };
            notifyIconMenuItem.Click += new EventHandler(NotifyIconMIClose_Click);
            notifyIconMenu.MenuItems.Add(notifyIconMenuItem);

            notifyIconMenu.Name = "notifyIconCMenu";
            this.notifyIcon00.ContextMenu = notifyIconMenu;
            this.notifyIcon00.Visible = true;
            //this.WindowState = WindowState.Minimized;
            //this.ShowInTaskbar = false;
            this.notifyIcon00.ShowBalloonTip(100);
        }
        private void NotifyIcon00_DoubleClick(object sender, EventArgs e)
        {
            foreach (Window wintest in System.Windows.Application.Current.Windows)
            {
                if (wintest.Name == "GalaxyMainXAML")
                {
                    switch (wintest.Visibility)
                    {
                        case Visibility.Hidden:
                            wintest.Show();
                            notifyIcon00.ContextMenu.MenuItems["notifyIconMain"].Checked = true;
                            break;
                        case Visibility.Visible:
                            wintest.Hide();
                            notifyIcon00.ContextMenu.MenuItems["notifyIconMain"].Checked = false;
                            break;
                    }
                }
            }

        }
        private void NotifyIconMIPurple_Click(object sender, EventArgs e)

        {
            System.Windows.Forms.MenuItem mi = (System.Windows.Forms.MenuItem)sender;
            if (mi.Checked == false)
            {
                foreach (Window wintest in System.Windows.Application.Current.Windows)
                {
                    if (wintest.Name == "GalaxyMainXAML")
                    {
                        wintest.Show();
                        wintest.Width = double.Parse(Properties.Resources.strPurple_Widht);
                        wintest.Height = double.Parse(Properties.Resources.strPurple_Height);                        
                        Grid grid = (Grid)wintest.Content;
                        Frame frame = (Frame)grid.Children[1];
                        frame.Source = new Uri("/GalaxyLotto;component/Page/PagePurple.xaml", UriKind.Relative);
                        //wintest..miSearchFreq_Click();
                    }
                }
                mi.Checked = true;
            }
            else
            {
                foreach (Window wintest in System.Windows.Application.Current.Windows)
                {
                    if (wintest.Name == "GalaxyMainXAML")
                    {
                        wintest.Show();
                        wintest.Width = double.Parse(Properties.Resources.strPage01_Widht); ;
                        wintest.Height = double.Parse(Properties.Resources.strPage01_Height);
                        Grid grid = (Grid)wintest.Content;
                        Frame frame = (Frame)grid.Children[1];
                        frame.Source = new Uri("/GalaxyLotto;component/Page/Page01.xaml", UriKind.Relative);
                    }
                }
                mi.Checked = false;
            }
        }
        private void NotifyIconMISearch_Click(object sender, EventArgs e)

        {
            System.Windows.Forms.MenuItem mi = (System.Windows.Forms.MenuItem)sender;
            if (mi.Checked == false)
            {
                foreach (Window wintest in System.Windows.Application.Current.Windows)
                {
                    if (wintest.Name == "GalaxyMainXAML")
                    {
                        wintest.Show();
                        wintest.Width = double.Parse(Properties.Resources.strSearch_Widht);
                        wintest.Height = double.Parse(Properties.Resources.strSearch_Height);
                        Grid grid = (Grid)wintest.Content;
                        Frame frame = (Frame)grid.Children[1];
                        frame.Source = new Uri("/GalaxyLotto;component/Page/PageSearch.xaml", UriKind.Relative);
                        //wintest..miSearchFreq_Click();
                    }
                }
                mi.Checked = true;
            }
            else
            {
                foreach (Window wintest in System.Windows.Application.Current.Windows)
                {
                    if (wintest.Name == "GalaxyMainXAML")
                    {
                        wintest.Width = double.Parse(Properties.Resources.strPage01_Widht); 
                        wintest.Height = double.Parse(Properties.Resources.strPage01_Height);
                        wintest.Show();
                        Grid grid = (Grid)wintest.Content;
                        Frame frame = (Frame)grid.Children[1];
                        frame.Source = new Uri("/GalaxyLotto;component/Page/Page01.xaml", UriKind.Relative);
                    }
                }
                mi.Checked = false;
            }
        }
        private void NotifyIconMIClose_Click(object sender, EventArgs e)

        {
            this.notifyIcon00.Visible = false;
            System.Windows.Application.Current.Shutdown();
        }
        private void ChkshowMain_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.MenuItem mi = (System.Windows.Forms.MenuItem)sender;
            foreach (Window wintest in System.Windows.Application.Current.Windows)
            {
                if (wintest.Name == "GalaxyMainXAML")
                {
                    switch (wintest.Visibility)
                    {
                        case Visibility.Hidden:
                            wintest.Show();
                            mi.Checked = true;
                            break;
                        case Visibility.Visible:
                            wintest.Hide();
                            mi.Checked = false;
                            break;
                    }
                }
            }
        }
        private void ChkshowProcBar_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.MenuItem mi = (System.Windows.Forms.MenuItem)sender;
            if (mi.Checked == false)
            {
                Monitor winMonitor = new Monitor();
                winMonitor.Show();
                mi.Checked = true;
            }
            else
            {
                foreach (Window wintest in System.Windows.Application.Current.Windows)
                {
                    if (wintest.Name == "ProcedureMonitor")
                    {
                        wintest.Hide();
                    }
                }
                mi.Checked = false;
            }
        }
        private void NotifyIcon_Loaded(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }
    }
}

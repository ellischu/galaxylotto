using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Xml;
using System.Data.SqlClient;
using System.Data;
using System.Windows.Controls.DataVisualization.Charting;
using System.Windows.Threading;
using System.Windows.Controls.Primitives;
using System.ComponentModel;
using GalaxyLotto.ClassLibrary;
using System.Text;
using static GalaxyLotto.ClassLibrary.CGLStructure;

namespace GalaxyLotto
{
    public partial class GalaxyMain : Window
    {
        #region 公用參數
        private CGLDataSet gDataSetGlotto = new CGLDataSet(TableType.LottoBig);
        //private stuGLSearch stuRibbonSearchOption = new stuGLSearch();
        //private stuWindowShow stuRibbonWinShow = new stuWindowShow();
        //private bool IsSearchAllGo = false;
        private readonly DispatcherTimer dispTimer = new DispatcherTimer();
        private delegate void UpdateProgressBarDelegate(DependencyProperty dp, Object value);
        private delegate Dictionary<string, object> GetSearchResultDelegate(StuGLSearch stuSearch);
        private delegate void SearchAllDelegate(StuGLSearch stuSearch);
        private delegate void SearchFrequencyDelegate(StuGLSearch stuSearch);
        private delegate DataTable delegateUpdateData(DataTable dtDataTable);
        WinNotifyIcon Winnotify01;
        #endregion 公用參數
        public GalaxyMain()
        {
            InitializeComponent();
            Winnotify01 = new WinNotifyIcon();
            Winnotify01.Show();
            GoMainPage();
        }

        #region 首頁
        private void GoMainPage_Click(object sender, RoutedEventArgs e)
        {
            GoMainPage();
        }
        private void GoMainPage()
        {
            if (frmMain.Source.OriginalString == "GalaxyLotto;component/Page/Page01.xaml")
            {
                frmMain.Refresh();
            }
            else
            {
                GalaxyMainXAML.Width = double.Parse(Properties.Resources.strPage01_Widht);
                GalaxyMainXAML.Height = double.Parse(Properties.Resources.strPage01_Height);
                Winnotify01.notifyIcon00.ContextMenu.MenuItems[2].Checked = false;
                Winnotify01.notifyIcon00.ContextMenu.MenuItems[3].Checked = false;
                frmMain.Source = new Uri("/GalaxyLotto;component/Page/Page01.xaml", UriKind.Relative);
            }
        }
        #endregion 首頁

        #region 更新
        /// <summary>
        /// 大樂透更新
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MiUpdateBig_Click(object sender, RoutedEventArgs e) // 大樂透更新
        {
            UpdateDataNow(TableType.LottoBig);
            //UpdateSum(TableType.LottoBig);
        }
        /// <summary>
        /// 威力彩更新
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MiUpdateWeli_Click(object sender, RoutedEventArgs e) // 威力彩更新 
        {
            UpdateDataNow(TableType.LottoWeli);
            //UpdateSum(TableType.LottoWeli);
        }
        /// <summary>
        /// 539 更新
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MiUpdate539_Click(object sender, RoutedEventArgs e) // 539 更新
        {
            UpdateDataNow(TableType.Lotto539);
            //UpdateSum(TableType.Lotto539);
        }
        private void MiUpdateDafu_Click(object sender, RoutedEventArgs e) // 大福彩更新
        {
            UpdateDataNow(TableType.LottoDafu);
            //UpdateSum(TableType.LottoDafu);
        }
        private void MiUpdateSix_Click(object sender, RoutedEventArgs e) // 六合彩更新
        {
            UpdateDataNow(TableType.LottoSix);
            //UpdateSum(TableType.LottoSix);
        }
        private void MiUpdateWC_Click(object sender, RoutedEventArgs e) // 中西曆更新
        {
            UpdateDataNow(TableType.DateWC);
        }
        private void MiUpdatePurple_Click(object sender, RoutedEventArgs e) //紫微更新
        {
            taskBarItemInfo1.ProgressState = System.Windows.Shell.TaskbarItemProgressState.Normal;
            taskBarItemInfo1.ProgressValue = 0;
            DataTable dtPurpleDate = new CGLFunc().GetPurpleDate();
            Window UpdatePurple = new Window()
            {
                WindowState = WindowState.Minimized,
                Visibility = Visibility.Hidden,
                WindowStyle = WindowStyle.None
            };
            delegateUpdateData delUpdateData = new delegateUpdateData(new CGLFunc().GetPurple);
            var dtResult = UpdatePurple.Dispatcher.Invoke(delUpdateData, DispatcherPriority.Background, dtPurpleDate);
            DataTable dtPurpleUpdate = (DataTable)dtResult;
            new CGLFunc().UpdateData(TableType.DataPurple, dtPurpleUpdate);
            ShowDataUpdate(TableType.DataPurple, dtPurpleUpdate);
            //pbStatusProcessBar.Visibility = Visibility.Hidden;
            // tiDisplay01.Focus();
            taskBarItemInfo1.ProgressState = System.Windows.Shell.TaskbarItemProgressState.None;
        }
        private void MiUpdateOptions_Click(object sender, RoutedEventArgs e) //參數資料更新
        {
            UpdateDataNow(TableType.DataOption);
        }
        /// <summary>
        /// 更新資料庫
        /// </summary>
        /// <param name="lottotype"></param>
        private void UpdateDataNow(TableType lottotype) //更新資料庫
        {
            taskBarItemInfo1.ProgressState = System.Windows.Shell.TaskbarItemProgressState.Normal;
            taskBarItemInfo1.ProgressValue = 0;
            DataTable tblCsvFile = new CGLFunc().ReadCsvFile(lottotype);
            new CGLFunc().UpdateData(lottotype, tblCsvFile);
            //ShowLastNumbers();
            GoMainPage();
            ShowDataUpdate(lottotype, tblCsvFile);
            //pbStatusProcessBar.Visibility = Visibility.Hidden;
            taskBarItemInfo1.ProgressState = System.Windows.Shell.TaskbarItemProgressState.None;
        }
        private void ShowDataUpdate(TableType lottotype, DataTable tblCsvFile) //顯示資料更新結果
        {
            CGLDataSet DataSet00 = new CGLDataSet(lottotype);
            #region Show in Web(html)
            string strCurrentDirectory = System.IO.Directory.GetCurrentDirectory();
            string strHtmlDirectory = System.IO.Path.Combine(strCurrentDirectory, "html");
            string strFileName = string.Format("update_{0}.html", DataSet00.ID);
            //StuGLSearch stuSearch00 = new StuGLSearch();
            new CGLFunc().FileExist(strHtmlDirectory, strFileName);
            //if (System.IO.File.Exists(System.IO.Path.Combine(strHtmlDirectory, strFileName)))
            //{
            //    System.IO.File.Delete(System.IO.Path.Combine(strHtmlDirectory, strFileName));
            //}
            //System.IO.File.Create(System.IO.Path.Combine(strHtmlDirectory, strFileName));
            StringBuilder stringbuilder = new StringBuilder();
            stringbuilder.AppendLine("<!DOCTYPE html>");
            stringbuilder.AppendLine("<html lang = 'zh-tw' xmlns = 'http://www.w3.org/1999/xhtml' >");
            stringbuilder.AppendLine("<head>");
            stringbuilder.AppendLine(string.Format("<title>{0}資料更新</title>", DataSet00.LottoDescription));
            stringbuilder.AppendLine("<meta charset = 'utf-8' />");
            stringbuilder.AppendLine("<link rel = 'stylesheet' href = '../css/glmain.css' />");
            stringbuilder.AppendLine("</head>");
            stringbuilder.AppendLine("<body>");
            stringbuilder.AppendLine("<table>");
            stringbuilder.AppendLine(string.Format("<caption>{0} 資料更新</caption>", DataSet00.LottoDescription));
            stringbuilder.AppendLine("<thead>");
            stringbuilder.AppendLine("<tr>");
            Dictionary<string, string> CFieldIDToName = new CGLFunc().CFieldNameID(1);
            Dictionary<string, string> CIDToName = new CGLFunc().CNameID(1);
            foreach (DataColumn dcColumn in tblCsvFile.Columns)
            {
                stringbuilder.AppendLine(string.Format("<th id='{0}' class='{0}' >", dcColumn.ColumnName));
                stringbuilder.AppendLine(CFieldIDToName[dcColumn.ColumnName]);
                stringbuilder.AppendLine("</th>");
            }

            stringbuilder.AppendLine("</tr>");
            stringbuilder.AppendLine("</thead>");
            stringbuilder.AppendLine("<tbody>");
            foreach (DataRow drRow in tblCsvFile.Rows)
            {
                stringbuilder.AppendLine("<tr>");
                for (int intcolumn = 0; intcolumn < tblCsvFile.Columns.Count; intcolumn++)
                {
                    stringbuilder.AppendLine("<td>");
                    string strdata = drRow[intcolumn].ToString();
                    if (CIDToName.ContainsKey(strdata))
                    {
                        strdata = CIDToName[strdata];
                    }
                    stringbuilder.AppendLine(strdata);
                    stringbuilder.AppendLine("</td>");
                }
                stringbuilder.AppendLine("</tr>");
            }
            stringbuilder.AppendLine("</tbody>");
            stringbuilder.AppendLine("</table>");
            stringbuilder.AppendLine("</body>");
            stringbuilder.AppendLine("</html>");

            System.IO.File.WriteAllText(System.IO.Path.Combine(strHtmlDirectory, strFileName), stringbuilder.ToString(), Encoding.UTF8);
            Process.Start(System.IO.Path.Combine(strHtmlDirectory, strFileName));
            #endregion
        }

        #endregion 更新

        #region 搜尋
        private void MiSearchFreq_Click(object sender, RoutedEventArgs e)
        {
            if (frmMain.Source.OriginalString == "GalaxyLotto;component/Page/PageSearch.xaml")
            {
                frmMain.Refresh();
            }
            else
            {
                GalaxyMainXAML.Width = double.Parse(Properties.Resources.strSearch_Widht);
                GalaxyMainXAML.Height = double.Parse(Properties.Resources.strSearch_Height);
                Winnotify01.notifyIcon00.ContextMenu.MenuItems[2].Checked = true;
                Winnotify01.notifyIcon00.ContextMenu.MenuItems[3].Checked = false;
                frmMain.Source = new Uri("/GalaxyLotto;component/Page/PageSearch.xaml", UriKind.Relative);
            }
        }
        #endregion 搜尋

        #region 樂透相關

        private void MiTaiwanLottoLink_Click(object sender, RoutedEventArgs e) //台灣彩券網站
        {
            Process.Start("http://www.taiwanlottery.com.tw");
        }
        private void MinfdLink_Click(object sender, RoutedEventArgs e) //電腦彩券樂透網
        {
            Process.Start("http://www.nfd.com.tw/lottery/index.htm");

        }
        private void MiarcLink_Click(object sender, RoutedEventArgs e) //樂透研究院
        {
            Process.Start("http://lotto.arclink.com.tw/");

        }
        private void MiLotto080eLink_Click(object sender, RoutedEventArgs e) //080e 贏發贏易
        {
            Process.Start("http://080e.com/");


        }
        private void MiLotto168Link_Click(object sender, RoutedEventArgs e) //樂透彩168
        {
            Process.Start("http://www.lotto168.com/");

        }
        private void MiauzonetLink_Click(object sender, RoutedEventArgs e) //奧索樂透網
        {
            Process.Start("http://lotto.auzonet.com/");

        }
        private void MibestshopLink_Click(object sender, RoutedEventArgs e) //樂透彩資訊網
        {
            Process.Start("http://lotto.bestshop.com.tw/");

        }
        private void Mihkjc_Click(object sender, RoutedEventArgs e) //六合彩官網
        {
            Process.Start("http://bet.hkjc.com/marksix/default.aspx");
        }
        private void Mipilio_Click(object sender, RoutedEventArgs e) //樂透彩幸運發財網
        {
            Process.Start("http://www.pilio.idv.tw/");
        }
        private void Miolotw_Click(object sender, RoutedEventArgs e) //棋王資訊科技網
        {
            Process.Start("http://www.olotw.com/");

        }

        #endregion 樂透相關      

        #region 命理相關
        private void Milead_Click(object sender, RoutedEventArgs e) //百合命相
        {
            Process.Start("http://www.lead.idv.tw/wgs/joytable.asp");

        }
        private void Micfarmcale_Click(object sender, RoutedEventArgs e) //網路農民曆
        {
            Process.Start("http://www.chu.edu.tw/~anita/");

        }
        private void Micfarmcale01_Click(object sender, RoutedEventArgs e) //農民曆
        {
            Process.Start("https://www.profate.com.tw/nongli");
        }

        #endregion 命理相關

        #region 其他連結
        private void MiGeomagnetic_Click(object sender, RoutedEventArgs e) //日本地磁
        {
            Process.Start("http://wdc.kugi.kyoto-u.ac.jp/dstdir/index.html");

        }

        #endregion 其他連結

        #region 功能
        private void MiMonitor_Click(object sender, RoutedEventArgs e) //進度視窗
        {
            bool IsWinSearchExist = false;
            foreach (Window wintest in System.Windows.Application.Current.Windows)
            {
                if (wintest.Name == "ProcedureMonitor")
                {
                    IsWinSearchExist = true;
                    wintest.Show();
                    wintest.Focus();
                }
            }
            if (!IsWinSearchExist)
            {
                Monitor winMonitor = new Monitor();
                winMonitor.Show();
            }

        }
        private void ChangeIcon(object sender, RoutedEventArgs e) //程式圖示
        {
            if (Winnotify01 != null)
            {
                RadioButton rbIcon = (RadioButton)sender;
                Uri iconUri;
                switch (rbIcon.Name)
                {
                    case "planet":
                        iconUri = new Uri("pack://siteoforigin:,,,/Resources/planet.ico", UriKind.RelativeOrAbsolute);
                        this.Icon = BitmapFrame.Create(iconUri);
                        Winnotify01.Icon = this.Icon;
                        Winnotify01.notifyIcon00.Icon = Properties.Resources.planet;
                        break;
                    case "planet_venus":
                        iconUri = new Uri("pack://siteoforigin:,,,/Resources/planet_venus.ico", UriKind.RelativeOrAbsolute);
                        this.Icon = BitmapFrame.Create(iconUri);
                        Winnotify01.Icon = this.Icon;
                        Winnotify01.notifyIcon00.Icon = Properties.Resources.planet_venus;
                        break;
                    case "planet_asteroid":
                        iconUri = new Uri("pack://siteoforigin:,,,/Resources/planet_asteroid.ico", UriKind.RelativeOrAbsolute);
                        this.Icon = BitmapFrame.Create(iconUri);
                        Winnotify01.Icon = this.Icon;
                        Winnotify01.notifyIcon00.Icon = Properties.Resources.planet_asteroid;
                        break;
                    case "planet_blackhole":
                        iconUri = new Uri("pack://siteoforigin:,,,/Resources/planet_blackhole.ico", UriKind.RelativeOrAbsolute);
                        this.Icon = BitmapFrame.Create(iconUri);
                        Winnotify01.Icon = this.Icon;
                        Winnotify01.notifyIcon00.Icon = Properties.Resources.planet_blackhole;
                        break;
                }
            }
        }

        #endregion 功能

        #region 紫薇
        private void ShowPurple(object sender, RoutedEventArgs e) //開啟紫薇
        {
            if (frmMain.Source.OriginalString == "GalaxyLotto;component/Page/PagePurple.xaml")
            {
                frmMain.Refresh();
            }
            else
            {
                GalaxyMainXAML.Width = double.Parse(Properties.Resources.strPurple_Widht);
                GalaxyMainXAML.Height = double.Parse(Properties.Resources.strPurple_Height);
                Winnotify01.notifyIcon00.ContextMenu.MenuItems[2].Checked = false;
                Winnotify01.notifyIcon00.ContextMenu.MenuItems[3].Checked = true;
                frmMain.Source = new Uri("/GalaxyLotto;component/Page/PagePurple.xaml", UriKind.Relative);
            }
        }

        #endregion 紫薇

        #region 離開
        private void CloseApp() //關閉主程式
        {
            Winnotify01.notifyIcon00.Visible = false;
            System.Windows.Application.Current.Shutdown();
        }
        private void GalaxyMainXAML_Closed(object sender, EventArgs e)
        {
            CloseApp();
        }
        private void RbClose_Click(object sender, RoutedEventArgs e)
        {
            CloseApp();
        }
        #endregion 離開

        #region 狀態列
        private void PbStatusProcessBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            //taskBarItemInfo1.ProgressValue = e.NewValue / pbStatusProcessBar.Maximum;
        }
        #endregion 狀態列

    }
}


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
using System.Windows.Navigation;
using System.Windows.Shapes;
using GalaxyLotto.ClassLibrary;
using System.Data.SqlClient;
using System.Data;
using static GalaxyLotto.ClassLibrary.CGLStructure;

namespace GalaxyLotto
{
    /// <summary>
    /// Page01.xaml 的互動邏輯
    /// </summary>
    public partial class Page01 : Page
    {
        public Page01()
        {
            InitializeComponent();
            ShowLastNumbers();
        }

        public void ShowLastNumbers() //顯示末期資料
        {
            CGLDataSet GLdataSet00 = new CGLDataSet(TableType.LottoBig);
            #region Set LottoType dictionary
            Dictionary<int, TableType> lottoType = new Dictionary<int, TableType>();
            Dictionary<int, RoutedEventHandler> dicRoutedEventHandler = new Dictionary<int, RoutedEventHandler>();
            lottoType.Add(0, TableType.LottoBig);
            dicRoutedEventHandler.Add(0, new RoutedEventHandler(BtnShowAllBig_Click));
            lottoType.Add(1, TableType.Lotto539);
            dicRoutedEventHandler.Add(1, new RoutedEventHandler(BtnShowAll539_Click));
            lottoType.Add(2, TableType.LottoDafu);
            dicRoutedEventHandler.Add(2, new RoutedEventHandler(BtnShowAllDafu_Click));
            lottoType.Add(3, TableType.LottoWeli);
            dicRoutedEventHandler.Add(3, new RoutedEventHandler(BtnShowAllWeli_Click));
            lottoType.Add(4, TableType.LottoSix);
            dicRoutedEventHandler.Add(4, new RoutedEventHandler(BtnShowAllSix_Click));
            #endregion

            //string[] strLastNumbers = new string[6];

            for (int i = 1; i <= 5; i++)
            {
                #region Set Label
                GLdataSet00 = new CGLDataSet(lottoType[i - 1]);
                StackPanel wp001 = new StackPanel()
                {
                    Height = 32,
                    VerticalAlignment = VerticalAlignment.Top,
                    Orientation = System.Windows.Controls.Orientation.Horizontal
                };
                System.Windows.Controls.Button btnShowAll = new System.Windows.Controls.Button()
                {
                    Content = "全顯",
                    Foreground = Brushes.Blue,
                    BorderThickness = new Thickness(0.0),
                    Margin = new Thickness(1, 1, 1, 1)
                };
                btnShowAll.Click += dicRoutedEventHandler[i - 1];
                wp001.Children.Add(btnShowAll);

                System.Windows.Controls.Label lbl001 = new System.Windows.Controls.Label()
                {
                    HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch,
                    HorizontalContentAlignment = System.Windows.HorizontalAlignment.Right,
                    Width = 75,
                    Content = GLdataSet00.LottoDescription + ":"
                };
                wp001.Children.Add(lbl001);

                Label lbl002 = new Label()
                {
                    HorizontalContentAlignment = System.Windows.HorizontalAlignment.Stretch,
                    Width = 90
                };
                #endregion
                try
                {
                    //if (oleConnect.State != ConnectionState.Connecting) oleConnect.Open();
                    StuGLData stuData00 = new StuGLData()
                    {
                        LottoType = lottoType[i - 1],
                        DataBaseType = DatabaseType.Data,

                        StrSELECT = " TOP 1 * ",
                        StrFROM = "tblData",
                        StrWHERE = string.Format("[{0}] > 0  ", "lngL1"),
                        StrORDER = string.Format("[{0}] DESC ; ", "lngDateSN")
                    };
                    DataTable endData = stuData00.GetSourceData(stuData00);
                    #region Set Last Nums
                    foreach (DataRow datarow00 in endData.Rows)
                    {
                        lbl002.Content = string.Format("{0}", datarow00["lngDateSN"]);
                        wp001.Children.Add(lbl002);
                        for (int Number00 = 1; Number00 <= GLdataSet00.CountNumber; Number00++)
                        {
                            Image pic001 = new Image()
                            {
                                Source = new BitmapImage(
                                new Uri("pack://siteoforigin:,,,/Resources/N" +
                                string.Format("{0:D2}", datarow00["lngL" + Number00.ToString()]) + ".png",
                                UriKind.Absolute))
                            };
                            wp001.Children.Add(pic001);
                        }
                        if (GLdataSet00.SCountNumber > 0)
                        {
                            Image picspe = new Image()
                            {
                                Source = new BitmapImage(
                                new Uri("pack://siteoforigin:,,,/Resources/spe.gif", UriKind.Absolute)
                                )
                            };
                            wp001.Children.Add(picspe);
                            Image pic001 = new Image()
                            {
                                Source = new BitmapImage(
                                new Uri("pack://siteoforigin:,,,/Resources/N" +
                                string.Format("{0:D2}", datarow00["lngS1"]) + ".png",
                                UriKind.Absolute))
                            };
                            wp001.Children.Add(pic001);
                        }
                    }
                    #endregion
                }
                catch (Exception e)
                {
                    if (e.Source != null)
                        System.Windows.MessageBox.Show(string.Format("ShowLastNumbers:{0}", e.Message), "Error", MessageBoxButton.OK);
                    //throw;
                }
                spMain.Children.Add(wp001);
            }
        }
        private void BtnShowAllBig_Click(object sender, RoutedEventArgs e) //顯示樂透彩
        {
            new CGLFunc().ShowAllData(TableType.LottoBig);
        }
        private void BtnShowAll539_Click(object sender, RoutedEventArgs e) //顯示539
        {
            new CGLFunc().ShowAllData(TableType.Lotto539);
        }
        private void BtnShowAllWeli_Click(object sender, RoutedEventArgs e) //顯示威力彩
        {
            new CGLFunc().ShowAllData(TableType.LottoWeli);
        }
        private void BtnShowAllSix_Click(object sender, RoutedEventArgs e) //顯示六合彩
        {
            new CGLFunc().ShowAllData(TableType.LottoSix);
        }
        private void BtnShowAllDafu_Click(object sender, RoutedEventArgs e) //顯示大福彩
        {
            new CGLFunc().ShowAllData(TableType.LottoDafu);
        }
    }
}

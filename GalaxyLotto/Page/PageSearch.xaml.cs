using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.DataVisualization.Charting;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using System.Xml;
using GalaxyLotto.ClassLibrary;
using System.Printing;
using System.Windows.Markup;
using System.Diagnostics;
using System.Windows.Controls.Primitives;
using static GalaxyLotto.ClassLibrary.CGLSearch;
using static GalaxyLotto.ClassLibrary.CGLFreq;
using static GalaxyLotto.ClassLibrary.CGLMethod;
using static GalaxyLotto.ClassLibrary.CGLData;

namespace GalaxyLotto
{
    /// <summary>
    /// PageSearch.xaml 的互動邏輯
    /// </summary>
    public partial class PageSearch : Page
    {
        #region 公用參數
        private CGLDataSet gDataSetGlotto = new CGLDataSet(TableType.LottoBig);
        private StuGLSearch stuRibbonSearchOption;

        //Dictionary<string, string> dicCurrentData;
        //List<int> lstCurrentData;
        public DataSet globalDataSet = new DataSet();
        private bool IsSearchAllGo = false;
        private readonly DispatcherTimer dispTimer = new DispatcherTimer();

        private delegate void UpdateProgressBarDelegate(DependencyProperty dp, Object value);
        private delegate Dictionary<string, object> GetSearchFreqDelegate(StuGLSearch stuSearch);
        private delegate void SearchAllDelegate(StuGLSearch stuSearch);
        private delegate void SearchDelegate(StuGLSearch stuSearch);

        BackgroundWorker bwSearchFreq, bwSearchFreqAll01, bwHit, bwBackgroundWorker00;
        #endregion 公用參數
        public PageSearch()
        {
            InitializeComponent();
            SearchOption_init();
        }

        #region initialization part
        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            foreach (Window wintest in System.Windows.Application.Current.Windows)
            {
                if (wintest.Name == "NotifyIcon")
                {
                    WinNotifyIcon WinNotifyIcon00 = (WinNotifyIcon)wintest;
                    WinNotifyIcon00.notifyIcon00.ContextMenu.MenuItems["notifyIconSearch"].Checked = true;
                }
                if (wintest.Name == "ProcedureMonitor")
                {
                    chkshowMonitor.IsChecked = true;
                }
            }
        }
        private StuGLSearch SetCompareString(StuGLSearch stuSearch00)
        {
            StuGLSearch stuReturn = stuSearch00;
            List<string> lstCompare = new List<string>();

            //stuReturn.StrCompares = "";

            #region FieldMode
            if (stuReturn.BoolFieldMode)
            {
                if ((string)cmbCompare01.SelectedValue != "none")
                {
                    lstCompare.Add(cmbCompare01.SelectedValue.ToString());
                }
                if ((string)cmbCompare02.SelectedValue != "none")
                {
                    lstCompare.Add(cmbCompare02.SelectedValue.ToString());
                }
                if ((string)cmbCompare03.SelectedValue != "none")
                {
                    lstCompare.Add(cmbCompare03.SelectedValue.ToString());
                }
                if ((string)cmbCompare04.SelectedValue != "none")
                {
                    lstCompare.Add(cmbCompare04.SelectedValue.ToString());
                }
                if ((string)cmbCompare05.SelectedValue != "none")
                {
                    lstCompare.Add(cmbCompare05.SelectedValue.ToString());
                }
            }
            #endregion FieldMode

            if (lstCompare.Count > 0)
            {
                lstCompare.Sort();
                stuReturn.StrCompares = string.Join("#", lstCompare.Distinct());
                //stuReturn = new CGLSearch().SetComparesDetail(stuReturn);
            }
            else
            {
                stuReturn.StrCompares = "gen";
                stuReturn.StrComparesDetail = "none";
                stuReturn.BoolFieldMode = false;
            }

            return stuReturn;
        }
        private void SearchOption_init()
        {
            #region Reading optionsearch.xml
            DataSet daOptions = new DataSet("Options");
            daOptions = SetOptionsFromXML();
            SetCombobox(ref cmbLottoType, daOptions.Tables["cmbLottoType"]);
            SetCombobox(ref cmbCompareType, daOptions.Tables["cmbCompareType"]);
            SetCombobox(ref cmbCompare01, daOptions.Tables["cmbCompare"]);
            SetCombobox(ref cmbCompare02, daOptions.Tables["cmbCompare"]);
            SetCombobox(ref cmbCompare03, daOptions.Tables["cmbCompare"]);
            SetCombobox(ref cmbCompare04, daOptions.Tables["cmbCompare"]);
            SetCombobox(ref cmbCompare05, daOptions.Tables["cmbCompare"]);
            SetCombobox(ref cmbNextNums, daOptions.Tables["cmbNextNums"]);
            SetCombobox(ref cmbNextStep, daOptions.Tables["cmbNextStep"]);
            //cmbQueryRangeStart.SelectedIndex = cmbDataRangeStart.SelectedIndex + gDataSetGlotto.MinStartDate;
            cmbNextNums.SelectedIndex = 0;
            cmbNextStep.SelectedIndex = 0;
            chkGeneral.IsChecked = true;
            chkField.IsChecked = false;
            chkNextNums.IsChecked = false;
            chkPeriod.IsChecked = false;
            chkshowGraphic.IsChecked = false;
            chkShowProcess.IsChecked = false;
            #endregion

            #region Set stuRibbonSearchOption
            stuRibbonSearchOption.LottoType = gDataSetGlotto.DataType;
            stuRibbonSearchOption.StrCompareType = cmbCompareType.SelectedValue.ToString();
            stuRibbonSearchOption.BoolGeneralMode = (bool)chkGeneral.IsChecked;
            stuRibbonSearchOption.BoolFieldMode = (bool)chkField.IsChecked;
            stuRibbonSearchOption.BoolNextNumsMode = (bool)chkNextNums.IsChecked;
            stuRibbonSearchOption.BoolPeriodMode = (bool)chkPeriod.IsChecked;
            stuRibbonSearchOption.LngDataEnd = long.Parse(cmbDataRangeEnd.SelectedValue.ToString());
            stuRibbonSearchOption.LngCurrentData = long.Parse(cmbDataRangeEnd.SelectedValue.ToString());
            stuRibbonSearchOption.IntMatchMin = int.Parse(txtDataLimit.Text);
            stuRibbonSearchOption.IntDataOffset = int.Parse(txtDataOffset.Text);
            stuRibbonSearchOption.IntSearchLimit = int.Parse(txtSearchLimit.Text);
            stuRibbonSearchOption.IntSearchOffset = int.Parse(txtSearchOffset.Text);
            stuRibbonSearchOption.IntTestTimes = int.Parse(txtTestTimes.Text);
            stuRibbonSearchOption.StrCompares = "gen";
            stuRibbonSearchOption.StrComparesDetail = "none";
            stuRibbonSearchOption.StrNextNums = "none";
            stuRibbonSearchOption.StrNextNumSpe = "none";
            stuRibbonSearchOption.IntNextNums = int.Parse(cmbNextNums.SelectedValue.ToString());
            stuRibbonSearchOption.IntNextStep = int.Parse(cmbNextStep.SelectedValue.ToString());
            if (chkShowProcess.IsChecked == true) { stuRibbonSearchOption.showProcess = ShowProcess.Visible; } else { stuRibbonSearchOption.showProcess = ShowProcess.Hide; }
            if (chkshowGraphic.IsChecked == true) { stuRibbonSearchOption.showGraphic = ShowGraphic.Visible; } else { stuRibbonSearchOption.showGraphic = ShowGraphic.Hide; }
            #endregion Set stuRibbonSearchOption
        }
        /// <summary>
        /// Read From XML file
        /// </summary>
        /// <returns></returns>
        private DataSet SetOptionsFromXML()
        {
            DataSet dsOptions = new DataSet("Options");
            //string strXmlFileName;
            //strXmlFileName = string.Format("{0}\\{1}", Environment.CurrentDirectory, "OptionSearch.xml");            
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.LoadXml(Properties.Resources.OptionSearch);
            //xmldoc.Load(@strXmlFileName);
            foreach (XmlNode node in xmldoc.DocumentElement)
            {
                if (node.HasChildNodes)
                {
                    string TableName = "";
                    DataColumn dcDataColumn;
                    DataRow drDataRow;
                    foreach (XmlNode child in node.ChildNodes)
                    {
                        if (child.Name == "name")
                        {
                            TableName = child.InnerText;
                            dsOptions.Tables.Add(TableName);
                            if (dsOptions.Tables[TableName].Columns.Count == 0)
                            {
                                dcDataColumn = new DataColumn()
                                {
                                    Caption = "id",
                                    ColumnName = "id",
                                    DataType = typeof(string)
                                };
                                dsOptions.Tables[TableName].Columns.Add(dcDataColumn);
                                dcDataColumn = new DataColumn()
                                {
                                    Caption = "description",
                                    ColumnName = "description",
                                    DataType = typeof(string)
                                };
                                dsOptions.Tables[TableName].Columns.Add(dcDataColumn);
                            }
                        }
                        else
                        {
                            if (child.HasChildNodes)
                            {
                                Dictionary<string, string> dicDataRow = new Dictionary<string, string>();
                                foreach (XmlNode child1 in child.ChildNodes)
                                {
                                    dicDataRow.Add(child1.Name, child1.InnerText);
                                    if (child1.Name == "description")
                                    {
                                        drDataRow = dsOptions.Tables[TableName].NewRow();
                                        drDataRow["id"] = dicDataRow["id"];
                                        drDataRow["description"] = dicDataRow["description"];
                                        dsOptions.Tables[TableName].Rows.Add(drDataRow);
                                        dicDataRow.Clear();
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return dsOptions;
        }
        private void SetCombobox(ref System.Windows.Controls.ComboBox comboBox, DataTable dataTable)
        {

            if (comboBox.ItemsSource != null && (comboBox.Name == "cmbLottoType"))
            {
                return;
            }
            comboBox.ItemsSource = dataTable.DefaultView;
            comboBox.SelectedValuePath = "id";
            comboBox.DisplayMemberPath = "description";
            comboBox.SelectedIndex = 0;
        }
        private void DataRangeStart_init(TableType lottotype)
        {
            string strlngTotalSN;
            DataTable dtDataTable = new DataTable();
            StuGLData stuData00 = new StuGLData()
            {
                LottoType = lottotype,
                DataBaseType = DatabaseType.Data,
                StrSELECT = "[lngTotalSN] , [lngDateSN]",
                StrFROM = "[tblData] ",
                StrWHERE = "[lngL1] = 0 ",
                StrORDER = "[lngTotalSN] ASC; "
            };

            #region get the first 0 in ingL1 as th last ingTotalSN
            dtDataTable = stuData00.GetSourceData(stuData00);
            strlngTotalSN = dtDataTable.Rows[0]["lngTotalSN"].ToString();
            dtDataTable.Clear();
            #endregion

            #region Get datastart,dataend tables
            stuData00 = new StuGLData()
            {
                LottoType = lottotype,
                DataBaseType = DatabaseType.Data,
                StrSELECT = "[lngTotalSN] , [lngDateSN]",
                StrFROM = "[tblData] ",
                StrWHERE = string.Format("[lngTotalSN] <= {0} ", strlngTotalSN),
                StrORDER = "[lngTotalSN] DESC; "
            };
            dtDataTable = stuData00.GetSourceData(stuData00);

            //strSQL = string.Format("SELECT [lngTotalSN] , [lngDateSN]  FROM {0} WHERE [lngTotalSN] <= {1} ORDER BY [lngTotalSN] DESC; "
            //                            , "tblData", strlngTotalSN);

            //odaAdapter = new SqlDataAdapter(strSQL, sqlConnect);
            //odaAdapter.Fill(dtDataTable);
            cmbDataRangeEnd.ItemsSource = dtDataTable.DefaultView;
            cmbDataRangeEnd.SelectedValuePath = "lngTotalSN";
            cmbDataRangeEnd.DisplayMemberPath = "lngDateSN";
            cmbDataRangeEnd.SelectedIndex = 0;
            int DataTableCount = dtDataTable.Rows.Count;
            stuRibbonSearchOption.LngDataStart = long.Parse(dtDataTable.Rows[DataTableCount - 1]["lngTotalSN"].ToString());
            #endregion
        }
        #endregion initialization part

        #region Option Part
        private void CmbLottoType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch ((int)cmbLottoType.SelectedIndex + 1)
            {
                case (int)TableType.LottoBig:
                    gDataSetGlotto = new CGLDataSet(TableType.LottoBig);
                    stuRibbonSearchOption.LottoType = TableType.LottoBig;
                    //stuRibbonSearchOption.IntMinStartDate = gDataSetGlotto.MinStartDate;
                    DataRangeStart_init(TableType.LottoBig);
                    break;
                case (int)TableType.Lotto539:
                    gDataSetGlotto = new CGLDataSet(TableType.Lotto539);
                    stuRibbonSearchOption.LottoType = TableType.Lotto539;
                    //stuRibbonSearchOption.IntMinStartDate = gDataSetGlotto.MinStartDate;
                    DataRangeStart_init(TableType.Lotto539);
                    break;
                case (int)TableType.LottoWeli:
                    gDataSetGlotto = new CGLDataSet(TableType.LottoWeli);
                    stuRibbonSearchOption.LottoType = TableType.LottoWeli;
                    //stuRibbonSearchOption.IntMinStartDate = gDataSetGlotto.MinStartDate;
                    DataRangeStart_init(TableType.LottoWeli);
                    break;
                case (int)TableType.LottoSix:
                    gDataSetGlotto = new CGLDataSet(TableType.LottoSix);
                    stuRibbonSearchOption.LottoType = TableType.LottoSix;
                    //stuRibbonSearchOption.IntMinStartDate = gDataSetGlotto.MinStartDate;
                    DataRangeStart_init(TableType.LottoSix);
                    break;
                case (int)TableType.LottoDafu:
                    gDataSetGlotto = new CGLDataSet(TableType.LottoDafu);
                    stuRibbonSearchOption.LottoType = TableType.LottoDafu;
                    //stuRibbonSearchOption.IntMinStartDate = gDataSetGlotto.MinStartDate;
                    DataRangeStart_init(TableType.LottoDafu);
                    break;
            }
            //cmbQueryRangeStart.SelectedIndex = cmbDataRangeStart.SelectedIndex + gDataSetGlotto.MinStartDate;

        }
        private void ChkField_Checked(object sender, RoutedEventArgs e)
        {
            rgCompare.IsEnabled = true;
            rgCompare.IsExpanded = true;
            stuRibbonSearchOption.BoolFieldMode = true;
        }
        private void ChkField_Unchecked(object sender, RoutedEventArgs e)
        {
            rgCompare.IsEnabled = false;
            rgCompare.IsExpanded = false;
            stuRibbonSearchOption.BoolFieldMode = false;
        }
        private void ChkNextNums_Checked(object sender, RoutedEventArgs e)
        {
            rgNextNums.IsEnabled = true;
            rgNextNums.IsExpanded = true;
            stuRibbonSearchOption.BoolNextNumsMode = true;
            cmbNextNums.SelectedIndex = 1;
        }
        private void ChkNextNums_Unchecked(object sender, RoutedEventArgs e)
        {
            rgNextNums.IsEnabled = false;
            rgNextNums.IsExpanded = false;
            stuRibbonSearchOption.BoolNextNumsMode = false;
            cmbNextNums.SelectedIndex = 0;
        }
        private void ChkPeriod_Checked(object sender, RoutedEventArgs e)
        {
            stuRibbonSearchOption.BoolPeriodMode = true;
        }
        private void ChkPeriod_Unchecked(object sender, RoutedEventArgs e)
        {
            stuRibbonSearchOption.BoolPeriodMode = false;
        }
        private void ChkRecalc_Checked(object sender, RoutedEventArgs e)
        {
            stuRibbonSearchOption.BoolRecalc = true;
        }
        private void ChkRecalc_Unchecked(object sender, RoutedEventArgs e)
        {
            stuRibbonSearchOption.BoolRecalc = false;
        }
        private void ChkGeneral_Unchecked(object sender, RoutedEventArgs e)
        {
            stuRibbonSearchOption.BoolGeneralMode = false;
        }
        private void ChkGeneral_Checked(object sender, RoutedEventArgs e)
        {
            stuRibbonSearchOption.BoolGeneralMode = true;
        }
        private void ChkShowProcess_Unchecked(object sender, RoutedEventArgs e)
        {
            stuRibbonSearchOption.showProcess = ShowProcess.Hide;
        }
        private void ChkshowGraphic_Checked(object sender, RoutedEventArgs e)
        {
            stuRibbonSearchOption.showGraphic = ShowGraphic.Visible;
        }
        private void ChkshowGraphic_Unchecked(object sender, RoutedEventArgs e)
        {
            stuRibbonSearchOption.showGraphic = ShowGraphic.Hide;
        }
        private void CmbCompareType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            stuRibbonSearchOption.StrCompareType = cmbCompareType.SelectedValue.ToString();
        }
        private void ChkShowProcess_Checked(object sender, RoutedEventArgs e)
        {
            stuRibbonSearchOption.showProcess = ShowProcess.Visible;
        }
        private void TxtDataLimit_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if ((e.Key >= Key.D0 && e.Key <= Key.D9) || (e.Key >= Key.NumPad0 && e.Key <= Key.NumPad9))
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
            if (e.Key == Key.Escape)
            {
                txtDataLimit.Text = "0";
            }
        }
        private void TxtDataOffset_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if ((e.Key >= Key.D0 && e.Key <= Key.D9) || (e.Key >= Key.NumPad0 && e.Key <= Key.NumPad9))
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
            if (e.Key == Key.Escape)
            {
                txtDataOffset.Text = "0";
            }
        }
        private void TxtSearchLimit_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if ((e.Key >= Key.D0 && e.Key <= Key.D9) || (e.Key >= Key.NumPad0 && e.Key <= Key.NumPad9))
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
            if (e.Key == Key.Escape)
            {
                txtSearchLimit.Text = "0";
            }
        }
        private void TxtSearchOffset_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if ((e.Key >= Key.D0 && e.Key <= Key.D9) || (e.Key >= Key.NumPad0 && e.Key <= Key.NumPad9))
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
            if (e.Key == Key.Escape)
            {
                txtSearchOffset.Text = "0";
            }
        }
        private void TxtTestTimes_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if ((e.Key >= Key.D0 && e.Key <= Key.D9) || (e.Key >= Key.NumPad0 && e.Key <= Key.NumPad9))
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
            if (e.Key == Key.Escape)
            {
                txtTestTimes.Text = "1";
            }
        }
        //private void cmbDataRangeStart_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{
        //    cmbQueryRangeStart.SelectedIndex = cmbDataRangeStart.SelectedIndex + gDataSetGlotto.MinStartDate;
        //    if (cmbDataRangeStart.SelectedValue != null) stuRibbonSearchOption.lngDataStart = long.Parse(cmbDataRangeStart.SelectedValue.ToString());
        //    if (cmbQueryRangeStart.SelectedValue != null) stuRibbonSearchOption.lngQueryStart = long.Parse(cmbQueryRangeStart.SelectedValue.ToString());
        //}
        private void CmbDataRangeEnd_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //cmbQueryRangeEnd.SelectedIndex = cmbDataRangeEnd.SelectedIndex;
            if (cmbDataRangeEnd.SelectedValue != null)
                stuRibbonSearchOption.LngDataEnd = long.Parse(cmbDataRangeEnd.SelectedValue.ToString());
            if (cmbDataRangeEnd.SelectedValue != null)
                stuRibbonSearchOption.LngCurrentData = long.Parse(cmbDataRangeEnd.SelectedValue.ToString());
            //if (cmbQueryRangeEnd.SelectedValue != null) stuRibbonSearchOption.lngQueryEnd = long.Parse(cmbQueryRangeEnd.SelectedValue.ToString());
        }
        //private void cmbQueryRangeStart_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{
        //    if (cmbQueryRangeStart.SelectedValue != null) stuRibbonSearchOption.lngQueryStart = long.Parse(cmbQueryRangeStart.SelectedValue.ToString());
        //}
        //private void cmbQueryRangeEnd_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{
        //    if (cmbQueryRangeEnd.SelectedValue != null) stuRibbonSearchOption.lngQueryEnd = long.Parse(cmbQueryRangeEnd.SelectedValue.ToString());
        //}
        private void CmbNextNums_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            stuRibbonSearchOption.IntNextNums = int.Parse(cmbNextNums.SelectedValue.ToString());
        }
        private void CmbNextStep_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            stuRibbonSearchOption.IntNextStep = int.Parse(cmbNextStep.SelectedValue.ToString());
        }
        private void TxtDataLimit_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (int.TryParse(txtDataLimit.Text, out int intTemp))
            {
                stuRibbonSearchOption.IntDataLimit = intTemp;
            }
            else
            {
                txtDataLimit.Text = "0";
            }
        }
        private void TxtDataLimit_LostFocus(object sender, RoutedEventArgs e)
        {
            if (int.TryParse(txtDataLimit.Text, out int intTemp))
            {
                stuRibbonSearchOption.IntDataLimit = intTemp;
                txtDataLimit.Text = intTemp.ToString();
            }
            else
            {
                txtDataLimit.Text = "0";
            }
        }
        private void TxtDataOffset_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (int.TryParse(txtDataOffset.Text, out int intTemp))
            {
                stuRibbonSearchOption.IntDataOffset = intTemp;
            }
        }
        private void TxtSearchLimit_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (int.TryParse(txtSearchLimit.Text, out int intTemp))
            {
                stuRibbonSearchOption.IntSearchLimit = intTemp;
            }
        }
        private void TxtSearchOffset_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (int.TryParse(txtSearchOffset.Text, out int intTemp))
            {
                stuRibbonSearchOption.IntSearchOffset = intTemp;
            }
        }
        private void TxtTestTimes_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (int.TryParse(txtTestTimes.Text, out int intTemp))
            {
                stuRibbonSearchOption.IntTestTimes = intTemp;
            }
        }
        private void TbMatchCount_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (int.TryParse(tbMatchCount.Text, out int intTemp))
            {
                stuRibbonSearchOption.IntMatchCount = intTemp;
            }
            else
            {
                txtDataLimit.Text = "1";
            }
        }
        private void TbMatchMin_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (int.TryParse(tbMatchMin.Text, out int intTemp))
            {
                stuRibbonSearchOption.IntMatchMin = intTemp;
            }
            else
            {
                txtDataLimit.Text = "1";
            }
        }
        private void ChkshowMonitor_Checked(object sender, RoutedEventArgs e)
        {
            bool IsWinMonitorExist = false;
            foreach (Window wintest in System.Windows.Application.Current.Windows)
            {
                if (wintest.Name == "ProcedureMonitor")
                {
                    IsWinMonitorExist = true;
                    wintest.Show();
                }
            }
            if (!IsWinMonitorExist)
            {
                Monitor winMonitor = new Monitor();
                winMonitor.Show();
            }
        }
        private void ChkshowMonitor_Unchecked(object sender, RoutedEventArgs e)
        {
            foreach (Window wintest in System.Windows.Application.Current.Windows)
            {
                if (wintest.Name == "ProcedureMonitor")
                {
                    wintest.Close();
                }
            }
        }
        private void ChkshowInWeb_Checked(object sender, RoutedEventArgs e)
        {
            stuRibbonSearchOption.BoolShowInWeb = true;
        }
        private void ChkshowInWeb_Unchecked(object sender, RoutedEventArgs e)
        {
            stuRibbonSearchOption.BoolShowInWeb = false;
        }

        #endregion Option Part

        private StuGLSearch ButtonClick(StuGLSearch stuSearch00)
        {
            stuSearch00 = SetCompareString(stuSearch00);
            //stuSearch00 = new CGLSearch().InitSearch(stuSearch00);
            return stuSearch00;
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e) //重設
        {
            cmbCompareType.SelectedIndex = 0;
            chkGeneral.IsChecked = true;
            chkField.IsChecked = false;
            chkNextNums.IsChecked = false;
            chkPeriod.IsChecked = false;
            //cmbDataRangeStart.SelectedIndex = 0;
            cmbDataRangeEnd.SelectedIndex = 0;
            cmbCompare01.SelectedIndex = 0;
            cmbCompare02.SelectedIndex = 0;
            cmbCompare03.SelectedIndex = 0;
            cmbCompare04.SelectedIndex = 0;
            cmbCompare05.SelectedIndex = 0;
            cmbNextNums.SelectedIndex = 0;
            cmbNextStep.SelectedIndex = 0;
            txtDataLimit.Text = "0";
            txtDataOffset.Text = "0";
            txtSearchLimit.Text = "0";
            txtSearchOffset.Text = "0";
            txtTestTimes.Text = "1";
            chkshowGraphic.IsChecked = false;
            chkShowProcess.IsChecked = false;
        }

        #region btnFreq 
        private void BtnFreq_Click(object sender, RoutedEventArgs e)
        {
            #region 設定 stuGLSearch
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSearchFreq = stuRibbonSearchOption;
            #endregion 設定 stuGLSearch

            #region Setting BackgroundWorker
            bwSearchFreq = new BackgroundWorker()
            {
                WorkerReportsProgress = true
            };
            bwSearchFreq.DoWork += BwSearchFreq_DoWork;
            bwSearchFreq.ProgressChanged += BwSearchFreq_ProgressChanged;
            bwSearchFreq.RunWorkerCompleted += BwSearchFreq_RunWorkerCompleted;
            bwSearchFreq.RunWorkerAsync(stuSearchFreq);
            btnFreq.Content = "計算中...";
            btnFreq.IsEnabled = false;

            #endregion Setting BackgroundWorker
        }
        private void BwSearchFreq_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Dictionary<string, object> dicFreq = new Dictionary<string, object>();
            dicFreq = (Dictionary<string, object>)e.Result;
            ShowResult((StuGLSearch)dicFreq["stuSearch"], (Dictionary<string, object>)dicFreq["dicFreq"]);
            this.tbStatusTextBlock.Text = "Ready";
            btnFreq.Content = "查詢頻率";
            btnFreq.IsEnabled = true;
            System.Media.SoundPlayer player = new System.Media.SoundPlayer()
            {
                SoundLocation = "Resources/Finish.wav"
            };
            player.Load();
            player.Play();
            player.Dispose();
            BackgroundWorker bwMyWork = (BackgroundWorker)sender;
            bwMyWork.Dispose();
        }
        private void BwSearchFreq_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 1)
            {
                this.tbStatusTextBlock.Text = "Work...";
            }
        }
        private void BwSearchFreq_DoWork(object sender, DoWorkEventArgs e)
        {
            //Console.WriteLine("bwFreq Thread :{0}", Thread.CurrentThread.ManagedThreadId);
            StuGLSearch stuSearchFreq = (StuGLSearch)e.Argument;
            BackgroundWorker bwMyWork = (BackgroundWorker)sender;
            bwMyWork.ReportProgress(1);
            Dictionary<string, object> dicFreq = new Dictionary<string, object>
            {
                { "stuSearch", stuSearchFreq },
                { "dicFreq", new CGLFreq().GetFreqdic(stuSearchFreq) }
            };
            e.Result = dicFreq;
        }
        /// <summary>
        /// 顯然查詢結果
        /// </summary>
        /// <param name="gstuGLSearch"></param>
        /// <param name="colResultFreq"></param>
        private void ShowResult(StuGLSearch gstuGLSearch, Dictionary<string, object> colResultFreq)
        {
            Dictionary<string, int> dicCurrentNums = new CGLSearch().GetCurrentDataNums(gstuGLSearch);
            #region Set ProcessBar
            double value = 0;
            Boolean blProcWin = false;
            Window winMain = new Window();
            System.Windows.Controls.ProgressBar pbProcessBar = new System.Windows.Controls.ProgressBar();
            foreach (Window wintest in System.Windows.Application.Current.Windows)
            {
                if (wintest.Name == "ProcedureMonitor")
                {
                    winMain = wintest;
                    blProcWin = true;
                }
            }

            if (blProcWin)
            {
                StackPanel spMain = (StackPanel)winMain.Content;
                StackPanel spStackPanel = new StackPanel();
                Grid gGrid = new Grid();
                TextBlock tbTextBlock = new TextBlock();
                foreach (Grid child00 in spMain.Children)
                {
                    if (child00.Name == "gShowFreq")
                    {
                        gGrid = child00;
                        pbProcessBar = (System.Windows.Controls.ProgressBar)gGrid.Children[0];
                        pbProcessBar.Minimum = 0;
                        pbProcessBar.Value = 0;
                        tbTextBlock = (TextBlock)gGrid.Children[1];
                        MultiBinding myBinding = new MultiBinding();
                        System.Windows.Data.Binding binding01 = new System.Windows.Data.Binding("Value");
                        System.Windows.Data.Binding binding02 = new System.Windows.Data.Binding("Maximum");
                        myBinding.StringFormat = "ShowFreq: {0}/{1}";
                        binding01.Source = pbProcessBar;
                        binding02.Source = pbProcessBar;
                        myBinding.Bindings.Add(binding01);
                        myBinding.Bindings.Add(binding02);
                        tbTextBlock.SetBinding(TextBlock.TextProperty, myBinding);
                    }
                }
            }
            UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(pbProcessBar.SetValue);
            #endregion Set ProcessBar

            #region Set Windows winResult 
            Window winFreq = new Window()
            {
                Title = gDataSetGlotto.LottoDescription,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.SingleBorderWindow
            };
            Grid gridResult = new Grid();
            ScrollViewer scrollviewer = new ScrollViewer { HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            StackPanel stackPanel = new StackPanel { Orientation = System.Windows.Controls.Orientation.Vertical };
            #endregion

            #region Set Title 
            System.Windows.Controls.Label lblTitle = new System.Windows.Controls.Label()
            {
                Background = Brushes.Yellow,
                Foreground = Brushes.Red,
                Content = gstuGLSearch.StrTitle
            };
            stackPanel.Children.Add(lblTitle);
            #endregion

            #region Set Expander CurrentData
            Expander expCurrentData = new Expander()
            {
                Header = "Current Data",
                ExpandDirection = ExpandDirection.Down,
                IsExpanded = true,
                Background = Brushes.YellowGreen
            };

            expCurrentData.Content = new CGLSearch().GetCurrentDataSp(gstuGLSearch);
            stackPanel.Children.Add(expCurrentData);
            #endregion

            #region Define the Expander
            Expander expFreqResult;
            System.Windows.Controls.DataGrid dgFreqResult;
            DataTable dtFreqResult;
            DataGridTextColumn dgtColumnFreq;

            Expander expProcessResult;
            System.Windows.Controls.DataGrid dgProcessResult;
            DataTable dtProcessResult;
            Dictionary<string, DataTable> dicProcessResult = (Dictionary<string, DataTable>)colResultFreq["Process"];
            DataGridTextColumn dgtColumnFreqProcess;

            Expander expGraphicResult;
            Dictionary<string, Dictionary<string, int>> dicGraphicResult =
                            (Dictionary<string, Dictionary<string, int>>)colResultFreq["NoSortResult"];
            Dictionary<string, int> dicFreqGraphic;
            Chart chtGraphic;
            ColumnSeries csGraphic;

            #endregion

            #region Define the Numbers brush type
            Dictionary<string, Style> dicNumBrushType = new Dictionary<string, Style>();
            Dictionary<string, Style> dicNumBrushType00 = new CNumberBrushType().Todictionary();

            foreach (var KeyPair in dicCurrentNums)
            {
                if (KeyPair.Value > 0)
                {
                    dicNumBrushType.Add(KeyPair.Value.ToString(), dicNumBrushType00[KeyPair.Key]);
                }
            }
            #endregion

            foreach (var dicFreqResult in colResultFreq["SortResult"] as Dictionary<string, Dictionary<string, int>>)
            {
                if (blProcWin)
                {
                    pbProcessBar.Maximum = colResultFreq.Count;
                    value++;
                    this.Dispatcher.Invoke(updatePbDelegate, DispatcherPriority.Background,
                                                new object[] { System.Windows.Controls.ProgressBar.ValueProperty, value });
                }

                #region Set Expander Frequency Result
                expFreqResult = new Expander()
                {
                    Header = string.Format("{0} Frequency", dicFreqResult.Key),
                    ExpandDirection = ExpandDirection.Down,
                    IsExpanded = true,
                    Background = Brushes.Pink
                };
                #region Set DataGrid dgFreqResult
                dgFreqResult = new System.Windows.Controls.DataGrid()
                {
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    RowDetailsVisibilityMode = DataGridRowDetailsVisibilityMode.Visible,
                    CanUserAddRows = true,
                    CanUserDeleteRows = false,
                    CanUserResizeColumns = true,
                    CanUserResizeRows = true,
                    CanUserSortColumns = true,
                    IsReadOnly = true,
                    AutoGenerateColumns = false,
                    HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch
                };
                #endregion

                #region Convert dictionary dicFreqResult to table dtFreqResult
                dtFreqResult = new CGLFunc().CDicTOTable(new CGLFunc().CDic_intTOstring(dicFreqResult.Value));
                #endregion

                dgFreqResult.ItemsSource = dtFreqResult.DefaultView;

                #region Set Columns of DataGrid dgFreqResult 

                if (dgFreqResult.Columns.Count == 0)
                {
                    foreach (var KeyPair in dicFreqResult.Value)
                    {
                        dgtColumnFreq = new DataGridTextColumn()
                        {
                            Header = string.Format("{0:00}", int.Parse(KeyPair.Key.Substring(4))),
                            Binding = new System.Windows.Data.Binding(KeyPair.Key),
                            IsReadOnly = true
                        };
                        if (dicNumBrushType.ContainsKey(string.Format("{0}", KeyPair.Key.Substring(4))))
                        {
                            dgtColumnFreq.HeaderStyle = dicNumBrushType[string.Format("{0}", KeyPair.Key.Substring(4))];
                        }
                        dgFreqResult.Columns.Add(dgtColumnFreq);
                    }
                }
                #endregion

                expFreqResult.Content = dgFreqResult;
                stackPanel.Children.Add(expFreqResult);
                #endregion
                #region Set Expander Graphic
                if (gstuGLSearch.showGraphic == ShowGraphic.Visible)
                {
                    expGraphicResult = new Expander()
                    {
                        Header = string.Format("{0} Frequency Graphic", dicFreqResult.Key),
                        ExpandDirection = ExpandDirection.Down,
                        IsExpanded = false,
                        Background = Brushes.LightGreen
                    };
                    dicFreqGraphic = new Dictionary<string, int>();
                    foreach (var KeyValuePair in dicGraphicResult[dicFreqResult.Key])
                    {
                        dicFreqGraphic.Add(string.Format("{0}", KeyValuePair.Key.Substring(4)), KeyValuePair.Value);
                    }
                    chtGraphic = new Chart()
                    {
                        Name = "gen",
                        VerticalAlignment = VerticalAlignment.Stretch,
                        HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch,
                        Height = 300
                    };
                    //chtGraphic.DataContext = dicFreqGraphic;
                    csGraphic = new ColumnSeries()
                    {
                        Title = "gen",
                        ItemsSource = dicFreqGraphic,
                        IndependentValuePath = "Key",
                        DependentValuePath = "Value",
                        IsSelectionEnabled = true
                    };
                    chtGraphic.Series.Add(csGraphic);
                    expGraphicResult.Content = chtGraphic;
                    stackPanel.Children.Add(expGraphicResult);
                }
                #endregion
                #region Set Expander Process
                if (gstuGLSearch.showProcess == ShowProcess.Visible)
                {
                    expProcessResult = new Expander()
                    {
                        Header = string.Format("{0} Process", dicFreqResult.Key),
                        ExpandDirection = ExpandDirection.Down,
                        IsExpanded = false,
                        Background = Brushes.Aqua
                    };
                    dgProcessResult = new System.Windows.Controls.DataGrid();
                    dtProcessResult = new CGLFunc().CTableShow(dicProcessResult[dicFreqResult.Key]);

                    #region Set DataGrid dgProcessResult
                    dgProcessResult.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
                    dgProcessResult.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
                    dgProcessResult.RowDetailsVisibilityMode = DataGridRowDetailsVisibilityMode.Visible;
                    dgProcessResult.CanUserAddRows = true;
                    dgProcessResult.CanUserDeleteRows = false;
                    dgProcessResult.CanUserResizeColumns = true;
                    dgProcessResult.CanUserResizeRows = true;
                    dgProcessResult.CanUserSortColumns = true;
                    dgProcessResult.IsReadOnly = true;
                    dgProcessResult.AutoGenerateColumns = false;
                    dgProcessResult.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                    #endregion

                    expProcessResult.Header += string.Format(" ({0}期)", dtProcessResult.Rows.Count);
                    dtProcessResult.DefaultView.Sort = "[lngTotalSN] DESC ";
                    dgProcessResult.ItemsSource = dtProcessResult.DefaultView;

                    #region Set Columns of DataGrid dgGeneralProcess
                    if (dgProcessResult.Columns.Count == 0)
                    {
                        foreach (DataColumn cFreqProcess in dtProcessResult.Columns)
                        {
                            dgtColumnFreqProcess = new DataGridTextColumn
                            {
                                Header = new CGLFunc().ConvertFieldNameID(cFreqProcess.Caption, 1),
                                Binding = new Binding(cFreqProcess.ColumnName),
                                IsReadOnly = true
                            };
                            #region Set Trigrer
                            if (cFreqProcess.ColumnName.Substring(0, 4) == "lngL")
                            {
                                Style Style00 = new Style()
                                {
                                    TargetType = typeof(DataGridCell)
                                };
                                foreach (var KeyValuePair in dicCurrentNums)
                                {
                                    DataTrigger dtTrig00 = new DataTrigger()
                                    {
                                        Binding = dgtColumnFreqProcess.Binding,
                                        Value = string.Format("{0}", KeyValuePair.Value)
                                    };
                                    Setter setter00 = new Setter()
                                    {
                                        Property = ForegroundProperty,
                                        Value = Brushes.Yellow
                                    };
                                    dtTrig00.Setters.Add(setter00);
                                    setter00 = new Setter()
                                    {
                                        Property = BackgroundProperty
                                    };
                                    switch (KeyValuePair.Key)
                                    {
                                        case "lngL1":
                                            setter00.Value = Brushes.DarkGreen;
                                            break;
                                        case "lngL2":
                                            setter00.Value = Brushes.DarkBlue;
                                            break;
                                        case "lngL3":
                                            setter00.Value = Brushes.Red;
                                            break;
                                        case "lngL4":
                                            setter00.Value = Brushes.DeepPink;
                                            break;
                                        case "lngL5":
                                            setter00.Value = Brushes.DarkRed;
                                            break;
                                        case "lngL6":
                                            setter00.Value = Brushes.Purple;
                                            break;
                                        case "lngL7":
                                            setter00.Value = Brushes.Maroon;
                                            break;
                                        case "lngS1":
                                            setter00.Value = Brushes.LightSeaGreen;
                                            break;
                                    }
                                    dtTrig00.Setters.Add(setter00);
                                    Style00.Triggers.Add(dtTrig00);
                                }
                                dgtColumnFreqProcess.CellStyle = Style00;
                            }
                            #endregion Set Trigrer
                            dgProcessResult.Columns.Add(dgtColumnFreqProcess);
                        }
                    }
                    #endregion

                    expProcessResult.Content = dgProcessResult;
                    expProcessResult.MaxHeight = 500;
                    stackPanel.Children.Add(expProcessResult);
                }

                #endregion
            }
            scrollviewer.Content = stackPanel;
            gridResult.Children.Add(scrollviewer);
            winFreq.Content = gridResult;
            #region Set ProcessBar
            if (blProcWin)
            {
                value = 0;
                this.Dispatcher.Invoke(updatePbDelegate, DispatcherPriority.Background,
                                            new object[] { System.Windows.Controls.ProgressBar.ValueProperty, value });
            }
            #endregion Set ProcessBar
            winFreq.Show();
        }
        #endregion btnFreq 

        #region btnFreqAll
        private void BtnFreqAll_Click(object sender, RoutedEventArgs e)
        {
            IsSearchAllGo = !IsSearchAllGo;
            #region Setting bwSearchFreqAll01
            bwSearchFreqAll01 = new BackgroundWorker()
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };
            bwSearchFreqAll01.DoWork += BwSearchFreqAll_DoWork;
            bwSearchFreqAll01.RunWorkerCompleted += BwSearchFreqAll_RunWorkerCompleted;
            bwSearchFreqAll01.ProgressChanged += BwSearchFreqAll_ProgressChanged;
            #endregion Setting bwSearchFreqAll01

            if (IsSearchAllGo && !bwSearchFreqAll01.IsBusy)
            {
                StuGLSearch stuSearchAll01 = ButtonClick(stuRibbonSearchOption);
                #region 設定 stuSearchAll01
                //stuSearchAll01 = new CGLSearch().InitSearch(stuSearchAll01);
                //stuSearchAll01 = new CGLSearch().GetMethodSN(stuSearchAll01);
                btnFreqAll.Content = "取消尋找";
                pbStatusProcess.Value = 0;
                pbStatusProcess.Visibility = Visibility.Visible;
                //taskBarItemInfo1.ProgressState = System.Windows.Shell.TaskbarItemProgressState.Normal;
                #endregion 設定 stuSearchAll01
                bwSearchFreqAll01.RunWorkerAsync(stuSearchAll01);
            }
            else
            {
                btnFreqAll.Content = "尋找全部";
                tbStatusTextBlock.Text = "Ready";
                pbStatusProcess.Visibility = Visibility.Hidden;
                IsSearchAllGo = false;
                bwSearchFreqAll01.CancelAsync();
                //bwSearchFreqAll02.CancelAsync();
            }

        }
        private void BwSearchFreqAll_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Dictionary<string, string> dicReport = (Dictionary<string, string>)e.UserState;
            switch (e.ProgressPercentage)
            {
                case 0:
                    pbStatusProcess.Value = e.ProgressPercentage;
                    pbStatusProcess.Maximum = double.Parse(dicReport["Max"]);
                    break;
                case 1:
                    pbStatusProcess.Value = pbStatusProcess.Value + e.ProgressPercentage;
                    break;
            }
            tbStatusTextBlock.Text = string.Format("{0} =>{1} {2}/{3} ", dicReport["Method"], dicReport["Title"], pbStatusProcess.Value, pbStatusProcess.Maximum);
        }
        private void BwSearchFreqAll_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            BackgroundWorker bwmyWorker = (BackgroundWorker)sender;
            if (bwmyWorker.CancellationPending)
            {
                MessageBox.Show("Canceled.");
            }
            pbStatusProcess.Visibility = Visibility.Hidden;
            //taskBarItemInfo1.ProgressState = System.Windows.Shell.TaskbarItemProgressState.None;
            btnFreqAll.Content = "尋找全部";
            tbStatusTextBlock.Text = "Ready";
            IsSearchAllGo = false;
            System.Media.SoundPlayer player = new System.Media.SoundPlayer()
            {
                SoundLocation = "Resources/Finish.wav"
            };
            player.Load();
            player.Play();
            player.Dispose();
            bwmyWorker.Dispose();

        }
        private void BwSearchFreqAll_DoWork(object sender, DoWorkEventArgs e)
        {
            #region 原始參數設定
            BackgroundWorker bwMyWork = (BackgroundWorker)sender;
            StuGLSearch stuSearch00 = (StuGLSearch)e.Argument;
            CGLDataSet DataSet00 = new CGLDataSet(stuSearch00.LottoType);
            #endregion 原始參數設定

            StuGLSearch stuSearchTemp = stuSearch00; //for test use
            string strDataBaseCS = CGLFunc.SetDataBase(TableType.DataMethod, DatabaseType.Method);
            using (SqlConnection sqlMethodConn = new SqlConnection(strDataBaseCS))
            {
                string strAscDesc = "ASC";
                switch (MessageBox.Show("YES for ASC mode \n NO for DESC mode", "Calculat in ASC or DESC mode", MessageBoxButton.YesNo, MessageBoxImage.Question))
                {
                    case MessageBoxResult.Yes:
                        strAscDesc = "ASC";
                        break;
                    case MessageBoxResult.No:
                        strAscDesc = "DESC";
                        break;
                }
                string qryMethod = string.Format("SELECT * FROM [{0}] ORDER BY [lngMethodSN] {1}", "tblMethod", strAscDesc);
                SqlCommand scCommand = new SqlCommand(qryMethod, sqlMethodConn);
                sqlMethodConn.Open();
                SqlDataReader sdrReader = scCommand.ExecuteReader();
                while (sdrReader.Read())
                {
                    StuGLMethod stuMethod00 = new StuGLMethod()
                    {
                        LngMethodSN = long.Parse(sdrReader[0].ToString())  //read lngMethodSN
                    };
                    stuMethod00 = stuMethod00.RebuildstuMethod(stuMethod00);               //get the stuMethod structure
                    stuSearchTemp = new CGLSearch().FromstuMethod(stuSearchTemp, stuMethod00); //rebuild the stuSearchTemp
                    stuSearchTemp.LngCurrentData = stuSearch00.LngCurrentData - 1;  //test the data of previous date
                    stuSearchTemp.IntInvestNums = 10;                               //that intInestNums equal to 10
                    if (!new CGLSearch().HasHitAllData(stuSearchTemp))                //Does it have Data
                    {
                        int intCount;
                        Dictionary<string, string> dicReport = new Dictionary<string, string>();
                        #region 開始計算

                        #region 計算頻率
                        #region 找出搜尋範圍
                        StuGLData stuData00 = new StuGLData()
                        {
                            LottoType = stuSearchTemp.LottoType,
                            DataBaseType = DatabaseType.Data,
                            StrSELECT = "*",
                            StrFROM = "tblFreq",
                            StrWHERE = string.Format("[lngMethodSN] = {0} ", stuSearchTemp.LngMethodSN.ToString()),
                            StrORDER = string.Format(" [lngTotalSN] DESC ;")
                        };
                        int intQueryStart = (int)stuSearch00.LngSearchStart;
                        int intQueryEnd = (int)stuSearch00.LngSearchEnd;
                        DataTable dtTest = stuData00.GetSourceData(stuData00);

                        if (dtTest.Rows.Count > 0)
                        {
                            if (int.Parse(dtTest.Rows[0]["lngTotalSN"].ToString()) >= intQueryStart)
                            {
                                intQueryStart = int.Parse(dtTest.Rows[0]["lngTotalSN"].ToString());
                            }
                        }
                        #endregion 找出搜尋範圍
                        intCount = 0;
                        dicReport.Add("Title", "Freq");
                        dicReport.Add("Method", stuSearchTemp.LngMethodSN.ToString());
                        dicReport.Add("Max", string.Format("{0}", intQueryEnd - intQueryStart + 1));
                        bwMyWork.ReportProgress(0, dicReport);
                        strDataBaseCS = CGLFunc.SetDataBase(stuSearchTemp.LottoType, DatabaseType.Freq);
                        using (SqlConnection sqlFreqConn = new SqlConnection(strDataBaseCS))
                        {
                            DataTable dtFreq = new DataTable("dtDataFreq");
                            string strQueryFreq = string.Format("SELECT * FROM {0} ", DataSet00.TableFreq);
                            SqlDataAdapter sdaFreq = new SqlDataAdapter(strQueryFreq, sqlFreqConn);
                            SqlCommandBuilder scbCBuilderFreq = new SqlCommandBuilder(sdaFreq);
                            sdaFreq.InsertCommand = scbCBuilderFreq.GetInsertCommand();
                            sdaFreq.DeleteCommand = scbCBuilderFreq.GetDeleteCommand();
                            sdaFreq.UpdateCommand = scbCBuilderFreq.GetUpdateCommand();
                            sdaFreq.FillSchema(dtFreq, SchemaType.Mapped);
                            sqlFreqConn.Close();

                            for (int intTotalSN = intQueryStart; intTotalSN <= intQueryEnd; intTotalSN++)
                            {
                                bwMyWork.ReportProgress(1, dicReport);
                                //Console.WriteLine("Freq : {0}/{1}", intTotalSN, intQueryEnd);
                                stuSearchTemp.LngCurrentData = (long)intTotalSN;
                                //stuSearchTemp = new CGLSearch().InitSearch(stuSearchTemp);

                                if (!new CGLFreq().HasFreqData(stuSearchTemp))
                                {
                                    //Console.WriteLine("SearchFreq : {0}", intTotalSN);
                                    foreach (DataRow drRow in new CGLFreq().GetFreqdt(stuSearchTemp).Rows)
                                    {
                                        DataRow drRow00 = dtFreq.NewRow();
                                        for (int i = 1; i < dtFreq.Columns.Count; i++)
                                        {
                                            drRow00[i] = drRow[i];
                                        }
                                        dtFreq.Rows.Add(drRow00);
                                    }
                                    if (intCount % 300 == 0)
                                    {
                                        sqlFreqConn.Open();
                                        sdaFreq.Update(dtFreq);
                                        dtFreq.Clear();
                                        sqlFreqConn.Close();
                                    }
                                    intCount++;
                                }
                                if (IsSearchAllGo == false)
                                {
                                    sqlFreqConn.Open();
                                    sdaFreq.Update(dtFreq);
                                    dtFreq.Dispose();
                                    scbCBuilderFreq.Dispose();
                                    sdaFreq.Dispose();
                                    sqlFreqConn.Dispose();
                                    bwMyWork.CancelAsync();
                                    return;
                                }
                            }
                            sqlFreqConn.Open();
                            sdaFreq.Update(dtFreq);
                            dtFreq.Dispose();
                            scbCBuilderFreq.Dispose();
                            sdaFreq.Dispose();
                        }
                        #endregion 計算頻率
                        #region 計算中獎率
                        Dictionary<string, double> dicSortFreq = new Dictionary<string, double>();
                        #region 找出搜尋範圍
                        stuData00 = new StuGLData()
                        {
                            LottoType = stuSearchTemp.LottoType,
                            DataBaseType = DatabaseType.Hit,
                            StrSELECT = "*",
                            StrFROM = string.Format("{0}", "tblHit"),
                            StrWHERE = string.Format("[lngMethodSN] = {0} AND [intInvestNums] =10 ", stuSearchTemp.LngMethodSN.ToString()),
                            StrORDER = string.Format("[lngTotalSN] DESC ;")
                        };
                        intQueryStart = (int)stuSearch00.LngSearchStart;
                        intQueryEnd = (int)stuSearch00.LngSearchEnd;
                        dtTest = stuData00.GetSourceData(stuData00);
                        if (dtTest.Rows.Count > 0)
                        {
                            if (int.Parse(dtTest.Rows[0]["lngTotalSN"].ToString()) >= intQueryStart)
                            {
                                intQueryStart = int.Parse(dtTest.Rows[0]["lngTotalSN"].ToString());
                            }
                        }
                        #endregion 找出搜尋範圍
                        intCount = 0;
                        dicReport["Title"] = "Hit";
                        dicReport["Max"] = string.Format("{0}", intQueryEnd - intQueryStart + 1);
                        bwMyWork.ReportProgress(0, dicReport);
                        strDataBaseCS = CGLFunc.SetDataBase(stuSearchTemp.LottoType, DatabaseType.Hit);
                        using (SqlConnection sqlHitConn = new SqlConnection(strDataBaseCS))
                        {
                            DataTable dtHit = new DataTable("dtDataHit");
                            string strQueryHit = string.Format("SELECT * FROM {0} ", DataSet00.TableHit);
                            SqlDataAdapter sdaHit = new SqlDataAdapter(strQueryHit, sqlHitConn);
                            SqlCommandBuilder scbCBuilderHit = new SqlCommandBuilder(sdaHit);
                            sdaHit.InsertCommand = scbCBuilderHit.GetInsertCommand();
                            sdaHit.DeleteCommand = scbCBuilderHit.GetDeleteCommand();
                            sdaHit.UpdateCommand = scbCBuilderHit.GetUpdateCommand();
                            sdaHit.FillSchema(dtHit, SchemaType.Mapped);
                            sqlHitConn.Close();

                            for (int intTotalSN = intQueryStart; intTotalSN <= intQueryEnd; intTotalSN++)
                            {
                                bwMyWork.ReportProgress(1, dicReport);
                                stuSearchTemp.LngCurrentData = (long)intTotalSN;
                                //stuSearchTemp = new CGLSearch().InitSearch(stuSearchTemp);

                                #region 處理所有非零的頻率
                                dicSortFreq.Clear();
                                dicSortFreq = new CGLFreq().GetSortedFreqdic(stuSearchTemp);
                                string[] strNoneZeroFreq = new CGLFunc().GetNonZeroArray(dicSortFreq);
                                #endregion 處理所有非零的頻率

                                for (int intInvest = DataSet00.CountNumber; intInvest <= 10; intInvest++)
                                {
                                    stuSearchTemp.IntInvestNums = intInvest;
                                    if (!new CGLSearch().HasHitData(stuSearchTemp) && !new CGLSearch().IsCurrentNumsZero(stuSearchTemp))
                                    {
                                        if ((intInvest <= strNoneZeroFreq.Length) && (strNoneZeroFreq.Length > 0))
                                        {
                                            #region 一般搜尋
                                            string[] strInvest = new string[intInvest];
                                            for (int i = 0; i < intInvest; i++)
                                            {
                                                strInvest[i] = strNoneZeroFreq[i];
                                            }
                                            stuSearchTemp.StrHitTestNums = String.Join("#", strInvest);
                                            //Console.WriteLine("SearchHit : {0}/{1}", intInvest, intTotalSN);
                                            foreach (DataRow drRow in new CGLSearch().SearchHitTable(stuSearchTemp).Rows)
                                            {
                                                DataRow drRow00 = dtHit.NewRow();
                                                for (int i = 1; i < dtHit.Columns.Count; i++)
                                                {
                                                    drRow00[i] = drRow[i];
                                                }
                                                dtHit.Rows.Add(drRow00);
                                                if (intCount % 300 == 0)
                                                {
                                                    sqlHitConn.Open();
                                                    sdaHit.Update(dtHit);
                                                    dtHit.Clear();
                                                    sqlHitConn.Close();
                                                    //Thread.Sleep(1000);
                                                }
                                                intCount++;
                                            }
                                            #endregion 一般搜尋
                                        }
                                        else
                                        {
                                            #region 增加一筆零中獎率
                                            DataRow drRow00 = dtHit.NewRow();
                                            drRow00["lngTotalSN"] = stuSearchTemp.LngCurrentData;
                                            drRow00["lngMethodSN"] = stuSearchTemp.LngMethodSN;
                                            drRow00["strHitTestNums"] = "none";
                                            drRow00["intInvestNums"] = intInvest;
                                            foreach (DataColumn dcColumn in dtHit.Columns)
                                            {
                                                if (dcColumn.ColumnName != "lngHitSN" && dcColumn.ColumnName != "lngTotalSN"
                                                    && dcColumn.ColumnName != "lngMethodSN" && dcColumn.ColumnName != "strHitTestNums"
                                                    && dcColumn.ColumnName != "intInvestNums")
                                                {
                                                    drRow00[dcColumn] = 0;
                                                }
                                            }
                                            dtHit.Rows.Add(drRow00);
                                            if (intCount % 300 == 0)
                                            {
                                                sqlHitConn.Open();
                                                sdaHit.Update(dtHit);
                                                dtHit.Clear();
                                                sqlHitConn.Close();
                                            }
                                            intCount++;
                                            #endregion 增加一筆零中獎率
                                        }
                                    }
                                }
                                if (IsSearchAllGo == false)
                                {
                                    sqlHitConn.Open();
                                    sdaHit.Update(dtHit);
                                    dtHit.Dispose();
                                    sdaHit.Dispose();
                                    scbCBuilderHit.Dispose();
                                    sqlHitConn.Dispose();
                                    bwMyWork.CancelAsync();
                                    return;
                                }
                            }
                            sqlHitConn.Open();
                            sdaHit.Update(dtHit);
                            dtHit.Dispose();
                            scbCBuilderHit.Dispose();
                            sdaHit.Dispose();
                        }
                        #endregion 計算中獎率
                        #region 計算總中獎率
                        #region 找出搜尋範圍
                        stuData00 = new StuGLData()
                        {
                            LottoType = stuSearchTemp.LottoType,
                            DataBaseType = DatabaseType.Hit,
                            StrSELECT = "*",
                            StrFROM = "tblHitAll",
                            StrWHERE = string.Format("[lngMethodSN] = {0} AND [intInvestNums] =10 ", stuSearchTemp.LngMethodSN.ToString()),
                            StrORDER = string.Format("[lngTotalSN] DESC ;")
                        };
                        intQueryStart = (int)stuSearch00.LngSearchStart;
                        intQueryEnd = (int)stuSearch00.LngSearchEnd;
                        dtTest = stuData00.GetSourceData(stuData00);
                        if (dtTest.Rows.Count > 0)
                        {
                            if (int.Parse(dtTest.Rows[0]["lngTotalSN"].ToString()) >= intQueryStart)
                            {
                                intQueryStart = int.Parse(dtTest.Rows[0]["lngTotalSN"].ToString());
                            }
                        }
                        #endregion 找出搜尋範圍
                        intCount = 0;
                        dicReport["Title"] = "HitAll";
                        dicReport["Max"] = string.Format("{0}", intQueryEnd - intQueryStart + 1);
                        bwMyWork.ReportProgress(0, dicReport);
                        DataTable dtHitAll = new DataTable("dtDataHitAll");
                        strDataBaseCS = CGLFunc.SetDataBase(stuSearchTemp.LottoType, DatabaseType.Hit);
                        using (SqlConnection sqlHitAllConn = new SqlConnection(strDataBaseCS))
                        {
                            string strQueryHitAll = string.Format("SELECT * FROM {0} ", DataSet00.TableHitAll);
                            SqlDataAdapter sdaHitAll = new SqlDataAdapter(strQueryHitAll, sqlHitAllConn);
                            SqlCommandBuilder scbCBuilderHitAll = new SqlCommandBuilder(sdaHitAll);
                            sdaHitAll.InsertCommand = scbCBuilderHitAll.GetInsertCommand();
                            sdaHitAll.DeleteCommand = scbCBuilderHitAll.GetDeleteCommand();
                            sdaHitAll.UpdateCommand = scbCBuilderHitAll.GetUpdateCommand();
                            sdaHitAll.FillSchema(dtHitAll, SchemaType.Mapped);
                            sqlHitAllConn.Close();

                            for (int intTotalSN = intQueryStart; intTotalSN <= intQueryEnd; intTotalSN++)
                            {
                                bwMyWork.ReportProgress(1, dicReport);
                                //Console.WriteLine("HitAll : {0}/{1}", intTotalSN, intQueryEnd);
                                stuSearchTemp.LngCurrentData = (long)intTotalSN;
                                //stuSearchTemp = new CGLSearch().InitSearch(stuSearchTemp);

                                for (int intInvest = DataSet00.CountNumber; intInvest <= 10; intInvest++)
                                {
                                    stuSearchTemp.IntInvestNums = intInvest;
                                    if (!new CGLSearch().IsCurrentNumsZero(stuSearchTemp) && !new CGLSearch().HasHitAllData(stuSearchTemp))
                                    {
                                        //Console.WriteLine("SearchHitAll : {0}/{1}", intInvest, intTotalSN);
                                        DataTable dtReturn = new CGLSearch().SearchHitAllTable(stuSearchTemp);
                                        foreach (DataRow drRow in dtReturn.Rows)
                                        {
                                            //dtHitAll.ImportRow(drRow);
                                            DataRow drRow00 = dtHitAll.NewRow();
                                            for (int i = 1; i < dtHitAll.Columns.Count; i++)
                                            {
                                                drRow00[i] = drRow[i];
                                            }
                                            dtHitAll.Rows.Add(drRow00);
                                        }
                                        if (intCount % 300 == 0)
                                        {
                                            sqlHitAllConn.Open();
                                            sdaHitAll.Update(dtHitAll);
                                            dtHitAll.Clear();
                                            sqlHitAllConn.Close();
                                        }
                                        intCount++;
                                    }
                                }
                                if (IsSearchAllGo == false)
                                {
                                    sqlHitAllConn.Open();
                                    sdaHitAll.Update(dtHitAll);
                                    dtHitAll.Dispose();
                                    scbCBuilderHitAll.Dispose();
                                    sdaHitAll.Dispose();
                                    sqlHitAllConn.Dispose();
                                    bwMyWork.CancelAsync();
                                    return;

                                }
                            }
                            sqlHitAllConn.Open();
                            sdaHitAll.Update(dtHitAll);
                            dtHitAll.Dispose();
                            scbCBuilderHitAll.Dispose();
                            sdaHitAll.Dispose();
                        }
                        #endregion 計算總中獎率

                        #endregion 開始計算
                    }
                }
                sdrReader.Close();
            }

        }
        #endregion btnFreqAll

        #region btnHit

        private void BtnHit_Click(object sender, RoutedEventArgs e)
        {
            #region 設定 stuSearchHit
            StuGLSearch stuSearchHit = ButtonClick(stuRibbonSearchOption);
            //stuSearchHit = new CGLSearch().InitSearch(stuSearchHit);
            //stuSearchHit = new CGLSearch().GetMethodSN(stuSearchHit);
            #endregion 設定 stuSearchHit
            bwHit = new BackgroundWorker()
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };
            bwHit.DoWork += BwHit_DoWork;
            bwHit.ProgressChanged += BwHit_ProgressChanged;
            bwHit.RunWorkerCompleted += BwHit_RunWorkerCompleted;
            bwHit.RunWorkerAsync(stuSearchHit);
            btnHit.Content = "計算中...";
            btnHit.IsEnabled = false;

        }
        private void BwHit_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Dictionary<string, object> dicHitResult = new Dictionary<string, object>();
            dicHitResult = (Dictionary<string, object>)e.Result;
            ShowHitResult((StuGLSearch)dicHitResult["stuSearch"], dicHitResult);
            this.tbStatusTextBlock.Text = "Ready";
            btnHit.Content = "中獎率";
            btnHit.IsEnabled = true;
            BackgroundWorker bwMywork = (BackgroundWorker)sender;
            bwMywork.Dispose();
        }
        private void BwHit_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 1)
            {
                this.tbStatusTextBlock.Text = "計算中...";
            }
        }
        private void BwHit_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bwMyWork = (BackgroundWorker)sender;
            bwMyWork.ReportProgress(1);
            StuGLSearch stuSearch00 = (StuGLSearch)e.Argument;
            CGLDataSet DataSet00 = new CGLDataSet(stuSearch00.LottoType);
            Dictionary<string, object> dicHitResult = new Dictionary<string, object>();
            Dictionary<string, object> dicCurrentHitResult = new Dictionary<string, object>();
            Dictionary<string, object> dicPreviousHitResult = new Dictionary<string, object>();

            #region Get Current ResultDic
            Dictionary<string, double> dicCurrentSortResult = new Dictionary<string, double>();
            string[] strCurrentNoneZero;

            dicCurrentSortResult = new CGLFreq().GetSortedFreqdic(stuSearch00);
            strCurrentNoneZero = new CGLFunc().GetNonZeroArray(dicCurrentSortResult);

            #region Set Current data
            dicCurrentHitResult.Add("CurrentData", new CGLSearch().GetCurrentData(stuSearch00));
            #endregion Set Current data

            dicCurrentHitResult.Add("TestNums:", string.Join("#", strCurrentNoneZero));

            //Console.WriteLine("Hit {0}:", stuSearch00.lngCurrentData);
            for (int intInvest = DataSet00.CountNumber; intInvest <= strCurrentNoneZero.Length; intInvest++)
            {
                string[] strCurrentSubTest = new string[intInvest];
                for (int i = 0; i < intInvest; i++)
                {
                    strCurrentSubTest[i] = strCurrentNoneZero[i];
                }

                stuSearch00.StrHitTestNums = String.Join("#", strCurrentSubTest);
                stuSearch00.IntInvestNums = intInvest;
                if (!new CGLSearch().HasHitData(stuSearch00) && !new CGLSearch().IsCurrentNumsZero(stuSearch00))
                {
                    //Console.WriteLine("Search Hit {0}:", stuSearch00.lngCurrentData);
                    new CGLSearch().SearchHit(stuSearch00);
                }
                StuGLHit stuCurrentHit = new StuGLHit();
                stuCurrentHit = new CGLSearch().SetHit(stuSearch00);
                stuCurrentHit = stuCurrentHit.GetHitSN(stuCurrentHit);

                dicCurrentHitResult.Add(intInvest.ToString() + " > " + stuCurrentHit.StrHitTestNums, stuCurrentHit.ToDicionary(stuCurrentHit));
            }


            #endregion Get Current ResultDic

            #region Get Previous Nums and ResultDic
            Dictionary<string, double> dicPreviousSortResult = new Dictionary<string, double>();
            string[] strPreviousNoneZero;
            StuGLSearch stuPreviosSearch = stuSearch00;
            stuPreviosSearch.LngCurrentData = stuSearch00.LngCurrentData - 1;
            //stuPreviosSearch = new CGLSearch().InitSearch(stuPreviosSearch);

            dicPreviousSortResult = new CGLFreq().GetSortedFreqdic(stuPreviosSearch);
            strPreviousNoneZero = new CGLFunc().GetNonZeroArray(dicPreviousSortResult);

            #region Set Previous data
            dicPreviousHitResult.Add("PreviousData", new CGLSearch().GetCurrentData(stuPreviosSearch));
            #endregion Set Previous data

            dicPreviousHitResult.Add("TestNums: ", string.Join("#", strPreviousNoneZero));

            //stuGLMethod stuPreviousMethod = new stuGLMethod();
            //Console.WriteLine("Hit {0}:", stuPreviosSearch.lngCurrentData);
            for (int intInvest = DataSet00.CountNumber; intInvest <= strPreviousNoneZero.Length; intInvest++)
            {
                string[] strPreviousSubTest = new string[intInvest];
                for (int i = 0; i < intInvest; i++)
                {
                    strPreviousSubTest[i] = strPreviousNoneZero[i];
                }
                stuPreviosSearch.StrHitTestNums = String.Join("#", strPreviousSubTest);
                stuPreviosSearch.IntInvestNums = intInvest;
                if (!new CGLSearch().HasHitData(stuPreviosSearch) && !new CGLSearch().IsCurrentNumsZero(stuPreviosSearch))
                {
                    //Console.WriteLine("Search Hit {0}:", stuPreviosSearch.lngCurrentData);
                    new CGLSearch().SearchHit(stuPreviosSearch);
                }
                StuGLHit stuPreviousHit = new StuGLHit();
                stuPreviousHit = new CGLSearch().SetHit(stuPreviosSearch);
                stuPreviousHit = stuPreviousHit.GetHitSN(stuPreviousHit);

                dicPreviousHitResult.Add(intInvest.ToString() + " > " + stuPreviousHit.StrHitTestNums, stuPreviousHit.ToDicionary(stuPreviousHit));
            }
            #endregion Get Previous Nums and ResultDic

            dicHitResult.Add("stuSearch", stuSearch00);
            dicHitResult.Add("Current", dicCurrentHitResult);
            dicHitResult.Add("Previous", dicPreviousHitResult);

            e.Result = dicHitResult;
        }
        private void ShowHitResult(StuGLSearch stuSearchHit, Dictionary<string, object> result)
        {
            #region Set Current and Previous HitResult
            CGLDataSet DataSet00 = new CGLDataSet(stuSearchHit.LottoType);
            Dictionary<string, object> dicCurrentHitResult = (Dictionary<string, object>)result["Current"];
            Dictionary<string, object> dicPreviousHitResult = (Dictionary<string, object>)result["Previous"];
            #endregion Set Current and Previous HitResult

            #region Set Windows winHitResult 
            Window winResult = new Window()
            {
                Title = DataSet00.LottoDescription,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.SingleBorderWindow,
                Name = "HitResult"
            };
            Grid gridResult = new Grid();
            ScrollViewer scrollviewer = new ScrollViewer { HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            StackPanel stackPanel = new StackPanel { Orientation = System.Windows.Controls.Orientation.Vertical };
            #endregion

            #region Set Title 
            System.Windows.Controls.Label lblTitle = new System.Windows.Controls.Label()
            {
                Background = Brushes.Yellow,
                Foreground = Brushes.Red,
                Content = string.Format("{0}: ", DataSet00.LottoDescription)
            };
            StuGLSearch testdate = new StuGLSearch();
            testdate = stuSearchHit;
            testdate.LngCurrentData = testdate.LngDataStart;
            Dictionary<string, string> dicCurrentData = new Dictionary<string, string>();
            dicCurrentData = new CGLSearch().GetCurrentData(testdate);
            lblTitle.Content += dicCurrentData["lngDateSN"] + " => ";
            testdate.LngCurrentData = testdate.LngDataEnd;
            dicCurrentData = new CGLSearch().GetCurrentData(testdate);
            lblTitle.Content += dicCurrentData["lngDateSN"];

            if (stuSearchHit.BoolFieldMode)
            {
                if (stuSearchHit.StrCompares.Length > 0)
                {
                    string[] strCompare = stuSearchHit.StrCompares.Split('#');
                    lblTitle.Content += " 相同 {";
                    foreach (string strCompareOption in strCompare)
                    {
                        lblTitle.Content += " " + new CGLFunc().ConvertFieldNameID(strCompareOption, 1) + " ,";
                    }
                    lblTitle.Content = lblTitle.Content.ToString().TrimEnd(',') + " }";
                }
            }
            if (stuSearchHit.BoolNextNumsMode)
            {
                lblTitle.Content += string.Format(" 間隔 {0} 期 , {1} 星托牌 ", stuSearchHit.IntNextStep, stuSearchHit.IntNextNums);
            }

            lblTitle.Content += string.Format(" 期數限制:{0}", stuSearchHit.IntMatchMin.ToString());
            lblTitle.Content += string.Format(" 查詢限制:{0}", stuSearchHit.IntSearchLimit.ToString());

            stackPanel.Children.Add(lblTitle);
            #endregion

            #region Set Expander CurrentData
            int intIndex = 0;
            Expander expCurrentData = new Expander()
            {
                Header = "Current Data",
                ExpandDirection = ExpandDirection.Down,
                IsExpanded = true,
                Background = Brushes.YellowGreen
            };
            StackPanel spCurrent = new StackPanel()
            {
                Name = "CurrentData"
            };
            foreach (var KeyValuePair in dicCurrentHitResult)
            {
                System.Windows.Controls.Label lblCurrent = new System.Windows.Controls.Label()
                {
                    Content = KeyValuePair.Key.ToString(),
                    Name = "lblCurrent" + intIndex.ToString()
                };

                if (KeyValuePair.Value.GetType() == Type.GetType("System.String"))
                {
                    lblCurrent.Content += KeyValuePair.Value.ToString();
                    spCurrent.Children.Add(lblCurrent);
                }
                else
                {
                    spCurrent.Children.Add(lblCurrent);
                    #region Set DataGrid dgItem
                    Dictionary<string, string> dicItems = (Dictionary<string, string>)KeyValuePair.Value;
                    System.Windows.Controls.DataGrid dgItem = new System.Windows.Controls.DataGrid()
                    {
                        Name = "dgCurrent" + intIndex.ToString(),
                        HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                        VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                        RowDetailsVisibilityMode = DataGridRowDetailsVisibilityMode.Visible,
                        CanUserAddRows = true,
                        CanUserDeleteRows = false,
                        CanUserResizeColumns = true,
                        CanUserResizeRows = true,
                        CanUserSortColumns = true,
                        IsReadOnly = true,
                        AutoGenerateColumns = false,
                        //dgCurrentData.VerticalAlignment = VerticalAlignment.Stretch;
                        HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch
                    };
                    #endregion

                    #region Convert dictionary dicItems to table dtItem
                    DataTable dtItem = new CGLFunc().CDicTOTable(dicItems);
                    #endregion

                    dgItem.ItemsSource = dtItem.DefaultView;

                    #region Set Columns of DataGrid dgItem 
                    DataGridTextColumn dgtColumn;
                    if (dgItem.Columns.Count == 0)
                    {
                        foreach (var KeyPair in dicItems)
                        {
                            dgtColumn = new DataGridTextColumn
                            {
                                Header = new CGLFunc().ConvertFieldNameID(KeyPair.Key, 1)
                            };
                            System.Windows.Data.Binding dgtBinding = new System.Windows.Data.Binding(KeyPair.Key);
                            dgtColumn.Binding = dgtBinding;
                            dgtColumn.IsReadOnly = true;
                            dgItem.Columns.Add(dgtColumn);
                        }
                    }
                    #endregion
                    spCurrent.Children.Add(dgItem);
                }
                intIndex++;
            }
            expCurrentData.Content = spCurrent;
            stackPanel.Children.Add(expCurrentData);
            #endregion

            #region Set Expander PreviousData
            intIndex = 0;
            Expander expPreviousData = new Expander()
            {
                Header = "Previous Data",
                ExpandDirection = ExpandDirection.Down,
                IsExpanded = true,
                Background = Brushes.LightPink
            };
            StackPanel spPrevious = new StackPanel()
            {
                Name = "PreviousData"
            };
            foreach (var KeyValuePair in dicPreviousHitResult)
            {
                System.Windows.Controls.Label lblPrevious = new System.Windows.Controls.Label()
                {
                    Content = KeyValuePair.Key.ToString(),
                    Name = "lblPrevious" + intIndex.ToString()
                };
                if (KeyValuePair.Value.GetType() == Type.GetType("System.String"))
                {
                    lblPrevious.Content += KeyValuePair.Value.ToString();
                    spPrevious.Children.Add(lblPrevious);

                }
                else
                {
                    spPrevious.Children.Add(lblPrevious);
                    Dictionary<string, string> dicItems = (Dictionary<string, string>)KeyValuePair.Value;
                    #region Set DataGrid dgItem
                    DataGrid dgItem = new DataGrid()
                    {
                        Name = "dgPrevious" + intIndex.ToString(),
                        HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                        VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                        RowDetailsVisibilityMode = DataGridRowDetailsVisibilityMode.Visible,
                        CanUserAddRows = true,
                        CanUserDeleteRows = false,
                        CanUserResizeColumns = true,
                        CanUserResizeRows = true,
                        CanUserSortColumns = true,
                        IsReadOnly = true,
                        AutoGenerateColumns = false,
                        HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch
                        //dgCurrentData.VerticalAlignment = VerticalAlignment.Stretch;
                    };
                    #endregion

                    #region Convert dictionary dicItems to table dtItem
                    DataTable dtItem = new CGLFunc().CDicTOTable(dicItems);
                    #endregion

                    dgItem.ItemsSource = dtItem.DefaultView;

                    #region Set Columns of DataGrid dgItem 
                    DataGridTextColumn dgtColumn;
                    if (dgItem.Columns.Count == 0)
                    {
                        foreach (var KeyPair in dicItems)
                        {
                            dgtColumn = new DataGridTextColumn
                            {
                                Header = new CGLFunc().ConvertFieldNameID(KeyPair.Key, 1)
                            };
                            System.Windows.Data.Binding dgtBinding = new System.Windows.Data.Binding(KeyPair.Key);
                            dgtColumn.Binding = dgtBinding;
                            dgtColumn.IsReadOnly = true;
                            dgItem.Columns.Add(dgtColumn);
                        }
                    }
                    #endregion
                    spPrevious.Children.Add(dgItem);
                }
                intIndex++;
            }
            expPreviousData.Content = spPrevious;
            stackPanel.Children.Add(expPreviousData);
            #endregion

            scrollviewer.Content = stackPanel;
            gridResult.Children.Add(scrollviewer);
            winResult.Content = gridResult;

            winResult.Show();

        }

        #endregion btnHit

        #region btnHitAll
        private void BtnHitAll_Click(object sender, RoutedEventArgs e)
        {
            #region 
            //taskBarItemInfo1.ProgressState = System.Windows.Shell.TaskbarItemProgressState.Normal;
            //taskBarItemInfo1.ProgressValue = 0;

            StuGLSearch stuSearchHitAll = ButtonClick(stuRibbonSearchOption);
            //stuSearchHitAll = new CGLSearch().InitSearch(stuSearchHitAll);
            //stuSearchHitAll = stuSearchHitAll.GetMethodSN(stuSearchHitAll);


            Window WinSearchHitAll = new Window()
            {
                WindowState = WindowState.Minimized,
                Visibility = Visibility.Hidden,
                WindowStyle = WindowStyle.None
            };
            GetSearchFreqDelegate delegateGetsearchHit = new GetSearchFreqDelegate(RunSearchHitAll);
            var Result = WinSearchHitAll.Dispatcher.Invoke(delegateGetsearchHit, DispatcherPriority.Background, stuSearchHitAll);

            //ShowHitResult(stuRun, (Dictionary<string, object>)Result);
            //taskBarItemInfo1.ProgressState = System.Windows.Shell.TaskbarItemProgressState.None;
            ShowHitResult(stuSearchHitAll, (Dictionary<string, object>)Result);
            #endregion
        }

        private Dictionary<string, object> RunSearchHitAll(StuGLSearch stuSearch00)
        {
            Dictionary<string, object> dicHitResult = new Dictionary<string, object>();

            return dicHitResult;
        }
        #endregion btnHitAll

        #region LastHit 前期中獎率
        private void BtnLastHit_Click(object sender, RoutedEventArgs e)
        {
            #region 設定 stuSearchLastHit
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSearchLastHit = stuRibbonSearchOption;
            stuSearchLastHit.IntMatchMin = 0;
            stuSearchLastHit.IntDataOffset = 0;
            stuSearchLastHit.IntSearchLimit = 0;
            stuSearchLastHit.IntSearchOffset = 0;
            #endregion 設定 stuSearchLastHit

            #region Setting BwLastHit
            bwBackgroundWorker00 = new BackgroundWorker()
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };
            bwBackgroundWorker00.DoWork += BwLastHit_DoWork;
            bwBackgroundWorker00.ProgressChanged += BwLastHit_ProgressChanged;
            bwBackgroundWorker00.RunWorkerCompleted += BwLastHit_RunWorkerCompleted;
            #endregion Setting BwLastHit

            bwBackgroundWorker00.RunWorkerAsync(stuSearchLastHit);
            btnLastHit.Content = "計算中...";
            btnLastHit.IsEnabled = false;

        }
        private void BwLastHit_DoWork(object sender, DoWorkEventArgs e)
        {
            #region Argument Setup
            StuGLSearch stuSearchLastHit = (StuGLSearch)e.Argument;
            BackgroundWorker bwMyWork = (BackgroundWorker)sender;
            bwMyWork.ReportProgress(1);
            Dictionary<string, object> dicArgument = new Dictionary<string, object>
            {
                { "stuSearch", stuSearchLastHit }
            };
            #endregion Setup
            #region 檢查 LastHit 是否有資料 
            if (!new CGLSearch().HasLastHitData(stuSearchLastHit, "tblLastHit00"))
            {
                new CGLSearch().SearchLastHit00(stuSearchLastHit);
            }
            if (!new CGLSearch().HasLastHitData(stuSearchLastHit, "tblLastHit01"))
            {
                new CGLSearch().SearchLastHit01(stuSearchLastHit);
            }
            if (!new CGLSearch().HasLastHitData(stuSearchLastHit, "tblLastHit02", 0, false))
            {
                new CGLSearch().SearchLastHit02(stuSearchLastHit);
            }
            if (!new CGLSearch().HasLastHitData(stuSearchLastHit, "tblLastHit03", 1, false))
            {
                new CGLSearch().SearchLastHitP(stuSearchLastHit, 1);
                new CGLSearch().SearchLastHitP(stuSearchLastHit, 2);
            }
            #endregion 檢查是否有資料
            e.Result = dicArgument;
        }
        private void BwLastHit_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 1)
            {
                this.tbStatusTextBlock.Text = "計算中...";
            }
        }
        private void BwLastHit_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            BackgroundWorker bwMywork = (BackgroundWorker)sender;
            Dictionary<string, object> dicArgument = (Dictionary<string, object>)e.Result;
            //ShowSum((stuGLSearch)dicArgument["stuSearch"], (DataSet)dicArgument["Freq"]);
            tbStatusTextBlock.Text = "Ready";
            btnLastHit.Content = "前期中獎率";
            btnLastHit.IsEnabled = true;
            bwMywork.Dispose();
        }
        #endregion LastHit

        // 表格部份 *****

        #region btnTableOddEven 奇數-偶數表
        private void BtnTableOddEven_Click(object sender, RoutedEventArgs e) //奇數-偶數表 
        {
            #region 設定 stuSearchOddEven
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSearchOddEven = stuRibbonSearchOption;
            #endregion 設定 stuSearchOddEven
            SearchDelegate OddEven = new SearchDelegate(SearchOddEven);
            Window wCalculate = new Window();
            wCalculate.Dispatcher.Invoke(OddEven, DispatcherPriority.Background, stuSearchOddEven);
            wCalculate.Close();
        }
        private void SearchOddEven(StuGLSearch stuSearchOddEven)
        {
            //CGLDataSet Dataset00 = new CGLDataSet(stuSearchOddEven.LottoType);


            #region 導入 200 期 數值
            //Dictionary<string, string> dicCurrentData = new CGLSearch().GetCurrentData(stuSearchOddEven);
            DataTable dtDataTable = new CGLOddEven().GetOddEven(stuSearchOddEven).Rows.Cast<DataRow>().Take(200).CopyToDataTable();
            #endregion 導入 200 期 數值

            //int[] intSection = { 5, 10, 25, 50, 100 };
            #region 計算 當期 奇數偶數 資料
            DataTable dtDataNext = new CGLOddEven().GetOddEvenNext(stuSearchOddEven, dtDataTable);
            #endregion

            #region Show Window

            #region Expander dgData

            #region DataGrid
            DataGrid dgOddEven = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Visible,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                AutoGenerateColumns = false
            };
            #region Set Columns of DataGrid dgOddEven
            DataGridTextColumn dgtColumnOddEven;
            if (dgOddEven.Columns.Count == 0)
            {
                foreach (DataColumn dcdgColumn in dtDataTable.Columns)
                {
                    dgtColumnOddEven = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1),
                        Binding = new Binding(dcdgColumn.ColumnName),
                        IsReadOnly = true
                    };
                    dgOddEven.Columns.Add(dgtColumnOddEven);
                }
            }
            #endregion
            dtDataTable.DefaultView.Sort = "[lngTotalSN] DESC ";
            dgOddEven.ItemsSource = dtDataTable.DefaultView;
            #endregion DataGrid

            Expander expOddEven = new Expander()
            {
                Header = "200 期 奇數偶數資料",
                ExpandDirection = ExpandDirection.Down,
                IsExpanded = true,
                Background = Brushes.Aqua,
                VerticalAlignment = VerticalAlignment.Top,
                HorizontalAlignment = HorizontalAlignment.Left,
                MaxHeight = 400,
                MaxWidth = 1400,
                Content = dgOddEven
            };
            #endregion Expander

            #region Expander dgDataNext

            #region DataGrid
            DataGrid dgOddEvenNext = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Visible,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                AutoGenerateColumns = false
            };
            #region Set Columns of DataGrid dgOddEven
            DataGridTextColumn dgtColumnOddEveNext;
            if (dgOddEvenNext.Columns.Count == 0)
            {
                foreach (DataColumn dcdgColumn in dtDataNext.Columns)
                {
                    dgtColumnOddEveNext = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1),
                        Binding = new System.Windows.Data.Binding(dcdgColumn.ColumnName),
                        IsReadOnly = true
                    };
                    dgOddEvenNext.Columns.Add(dgtColumnOddEveNext);
                }
            }
            #endregion
            dgOddEvenNext.ItemsSource = dtDataNext.DefaultView;
            #endregion DataGrid

            Expander expOddEvenNext = new Expander()
            {
                Header = "當期奇數偶數預測表",
                ExpandDirection = ExpandDirection.Down,
                IsExpanded = true,
                Background = Brushes.Aqua,
                VerticalAlignment = VerticalAlignment.Top,
                HorizontalAlignment = HorizontalAlignment.Left,
                MaxHeight = 400,
                MaxWidth = 1400,
                Content = dgOddEvenNext
            };
            #endregion Expander

            #region Expander Graphic

            #region Graphic OddEven
            Chart chtGraphic = new Chart()
            {
                Name = "gen",
                VerticalAlignment = VerticalAlignment.Stretch,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                Height = 300,
                Width = 25 * 200
            };
            ColumnSeries csGraphic = new ColumnSeries()
            {
                Title = "Odds",
                //dtDataTable.DefaultView.Sort = "[lngTotalSN] ASC";
                ItemsSource = dtDataTable.DefaultView,
                IndependentValuePath = "lngTotalSN",
                DependentValuePath = "intOdd",
                IsSelectionEnabled = true
            };
            chtGraphic.Series.Add(csGraphic);

            csGraphic = new ColumnSeries()
            {
                Title = "Evens",
                ItemsSource = dtDataTable.DefaultView,
                IndependentValuePath = "lngTotalSN",
                DependentValuePath = "intEven",
                IsSelectionEnabled = true
            };
            chtGraphic.Series.Add(csGraphic);
            #endregion Graphic

            ScrollViewer svViewerGraphic = new ScrollViewer()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = chtGraphic
            };
            Expander expOddEvenGraphic = new Expander()
            {
                Header = "200 期 奇數偶數圖表",
                ExpandDirection = ExpandDirection.Down,
                IsExpanded = false,
                Background = Brushes.Aqua,
                MaxHeight = 700,
                Content = svViewerGraphic
            };
            #endregion Expander Graphic

            #region StackPanel
            StackPanel spMain = new StackPanel()
            {
                CanVerticallyScroll = true,
                CanHorizontallyScroll = true
            };
            spMain.Children.Add(new CGLSearch().GetCurrentDataSp(stuSearchOddEven));
            spMain.Children.Add(expOddEven);
            spMain.Children.Add(expOddEvenNext);
            spMain.Children.Add(expOddEvenGraphic);
            //spMain.Children.Add(expOddEvenGraphic05);
            //spMain.Children.Add(expOddEvenGraphic10);
            //spMain.Children.Add(expOddEvenGraphic25);
            #endregion StackPanel

            #region ScrollViewer
            ScrollViewer svViewer = new ScrollViewer()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = spMain
            };
            #endregion ScrollViewer

            #region Window
            Window wOddEven = new Window()
            {
                Name = "tblOddEven",
                Title = string.Format("奇數-偶數表 ({0})", stuSearchOddEven.StrTitle),
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.SingleBorderWindow
            };
            wOddEven.Show();
            wOddEven.Content = svViewer;
            #endregion Window

            #endregion Show Window
        }

        #endregion btnTableOddEven

        private void BtnTableHighLow_Click(object sender, RoutedEventArgs e) //大數-小數表
        {

            #region 設定 stuSearchHighLow
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSearchHighLow = stuRibbonSearchOption;
            stuSearchHighLow.IntMatchMin = 0;
            stuSearchHighLow.IntDataOffset = 0;
            stuSearchHighLow.IntSearchLimit = 0;
            stuSearchHighLow.IntSearchOffset = 0;
            #endregion 設定 stuSearchHighLow
            CGLDataSet Dataset00 = new CGLDataSet(stuSearchHighLow.LottoType);
            #region 檢查是否有資料
            StuGLData stuData00 = new StuGLData()
            {
                LottoType = stuSearchHighLow.LottoType,
                DataBaseType = DatabaseType.HighLow,
                StrSELECT = " * ",
                StrFROM = "[qryHighLow]",
                StrWHERE = string.Format("[{0}] = {1} AND [{2}] = {3} ;",
                                                    "lngMethodSN", stuSearchHighLow.LngMethodSN,
                                                    "lngTotalSN", stuSearchHighLow.LngCurrentData - 1),
                StrORDER = ""
            };
            DataTable dtDataTable = stuData00.GetSourceData(stuData00); ;
            if (dtDataTable.Rows.Count == 0) { new CGLSearch().SearchHighLow(stuSearchHighLow); }
            #endregion 檢查是否有資料

            int[] intSection = { 5, 10, 25, 50, 100 };

            #region 導入 200 期 數值

            #region Query String
            Dictionary<string, string> dicCurrentData = new Dictionary<string, string>();
            dicCurrentData = new CGLSearch().GetCurrentData(stuSearchHighLow);
            stuData00 = new StuGLData()
            {
                LottoType = stuSearchHighLow.LottoType,
                DataBaseType = DatabaseType.HighLow,
                StrSELECT = " TOP 200 ",
                StrFROM = "[qryHighLow] ",
                StrWHERE = string.Format("[{0}] = {1} AND [{2}] < {3} ",
                                        "lngMethodSN", stuSearchHighLow.LngMethodSN,
                                        "lngTotalSN", stuSearchHighLow.LngCurrentData),
                StrORDER = "[lngTotalSN] DESC ;"
            };
            #region set select
            stuData00.StrSELECT += "[lngTotalSN] , [lngMethodSN] , [lngDateSN] , ";
            for (int i = 1; i <= Dataset00.CountNumber; i++)
            {
                stuData00.StrSELECT += string.Format("[lngL{0}] , ", i.ToString());
            }
            if (Dataset00.DataType != TableType.LottoWeli && Dataset00.SCountNumber > 0)
            {
                stuData00.StrSELECT += "[lngS1] , ";
            }
            stuData00.StrSELECT += "[intHigh] , [intLow] , ";
            foreach (int ints in intSection)
            {
                stuData00.StrSELECT += string.Format("[intTHighA{0}] , [intTLowA{0}] , [sglPHLA{0}] , ", ints.ToString("D2"));
                stuData00.StrSELECT += string.Format("[intTHighB{0}] , [intTLowB{0}] , [sglPHLB{0}] , ", ints.ToString("D2"));
            }
            stuData00.StrSELECT += "[sglTPHLB] ";
            #endregion

            #region  Set Compare Option 
            if (stuSearchHighLow.BoolFieldMode && stuSearchHighLow.StrCompares != "gen")
            {
                string[] strCompare = stuSearchHighLow.StrCompares.Split('#');
                foreach (string strCompareOption in strCompare)
                {
                    stuData00.StrWHERE += string.Format(" {0} [{1}] = '{2}' "
                                                        , stuSearchHighLow.StrCompareType
                                                        , strCompareOption
                                                        , dicCurrentData[strCompareOption]);
                }
            }
            #endregion

            #endregion Query String
            dtDataTable = stuData00.GetSourceData(stuData00);

            #endregion 導入 200 期 數值

            #region 計算下一期

            #region 建立 dtDataNext
            DataTable dtDataNext = new DataTable();


            #region 建立 Columns
            DataColumn dcColumn = new DataColumn()
            {
                ColumnName = "intHigh",
                Caption = "intHigh",
                DataType = typeof(int)
            };
            dtDataNext.Columns.Add(dcColumn);

            dcColumn = new DataColumn()
            {
                ColumnName = "intLow",
                Caption = "intLow",
                DataType = typeof(int)
            };
            dtDataNext.Columns.Add(dcColumn);

            foreach (int ints in intSection)
            {
                dcColumn = new DataColumn()
                {
                    ColumnName = string.Format("intTHighA{0}", ints.ToString("D2")),
                    Caption = string.Format("intTHighA{0}", ints.ToString("D2")),
                    DataType = typeof(int)
                };
                dtDataNext.Columns.Add(dcColumn);

                dcColumn = new DataColumn()
                {
                    ColumnName = string.Format("intTLowA{0}", ints.ToString("D2")),
                    Caption = string.Format("intTLowA{0}", ints.ToString("D2")),
                    DataType = typeof(int)
                };
                dtDataNext.Columns.Add(dcColumn);

                dcColumn = new DataColumn()
                {
                    ColumnName = string.Format("sglPHLA{0}", ints.ToString("D2")),
                    Caption = string.Format("sglPHLA{0}", ints.ToString("D2")),
                    DataType = typeof(float)
                };
                dtDataNext.Columns.Add(dcColumn);

                dcColumn = new DataColumn()
                {
                    ColumnName = string.Format("intTHighB{0}", ints.ToString("D2")),
                    Caption = string.Format("intTHighB{0}", ints.ToString("D2")),
                    DataType = typeof(int)
                };
                dtDataNext.Columns.Add(dcColumn);

                dcColumn = new DataColumn()
                {
                    ColumnName = string.Format("intTLowB{0}", ints.ToString("D2")),
                    Caption = string.Format("intTLowB{0}", ints.ToString("D2")),
                    DataType = typeof(int)
                };
                dtDataNext.Columns.Add(dcColumn);

                dcColumn = new DataColumn()
                {
                    ColumnName = string.Format("sglPHLB{0}", ints.ToString("D2")),
                    Caption = string.Format("sglPHLB{0}", ints.ToString("D2")),
                    DataType = typeof(float)
                };
                dtDataNext.Columns.Add(dcColumn);
            }
            dcColumn = new DataColumn()
            {
                ColumnName = "sglTPHLB",
                Caption = "sglTPHLB",
                DataType = typeof(float)
            };
            dtDataNext.Columns.Add(dcColumn);


            #endregion 建立 Columns

            #endregion 建立 tblDataNext

            #region 計算迴圈
            for (int i = 0; i <= Dataset00.CountNumber; i++)
            {
                DataRow drRow = dtDataNext.NewRow();
                int intHigh = i, intLow = Dataset00.CountNumber - i;
                drRow["intHigh"] = intHigh; drRow["intLow"] = intLow;
                float sglTPHLB = 0;
                foreach (int ints in intSection)
                {
                    int intTHighTemp, intTLowTemp;
                    float sglPHighLowTemp;
                    #region Part A
                    intTHighTemp = int.Parse(dtDataTable.Rows[0][string.Format("intTHighA{0}", ints.ToString("D2"))].ToString()) -
                                 int.Parse(dtDataTable.Rows[ints - 1]["intHigh"].ToString()) +
                                 intHigh;

                    intTLowTemp = int.Parse(dtDataTable.Rows[0][string.Format("intTLowA{0}", ints.ToString("D2"))].ToString()) -
                                 int.Parse(dtDataTable.Rows[ints - 1]["intLow"].ToString()) +
                                 intLow;

                    sglPHighLowTemp = (float)((intTHighTemp - intTLowTemp) * 100f) / (float)(intTHighTemp + intTLowTemp);
                    drRow[string.Format("intTHighA{0}", ints.ToString("D2"))] = intTHighTemp;
                    drRow[string.Format("intTLowA{0}", ints.ToString("D2"))] = intTLowTemp;
                    drRow[string.Format("sglPHLA{0}", ints.ToString("D2"))] = sglPHighLowTemp;
                    sglTPHLB += sglPHighLowTemp;
                    #endregion Part A
                    #region Part B
                    intTHighTemp = int.Parse(dtDataTable.Rows[0][string.Format("intTHighB{0}", ints.ToString("D2"))].ToString()) -
                                  int.Parse(dtDataTable.Rows[(ints * 2) - 1]["intHigh"].ToString()) +
                                  int.Parse(dtDataTable.Rows[ints - 1]["intHigh"].ToString());

                    intTLowTemp = int.Parse(dtDataTable.Rows[0][string.Format("intTLowB{0}", ints.ToString("D2"))].ToString()) -
                                   int.Parse(dtDataTable.Rows[(ints * 2) - 1]["intLow"].ToString()) +
                                   int.Parse(dtDataTable.Rows[ints - 1]["intLow"].ToString());

                    sglPHighLowTemp = (float)((intTHighTemp - intTLowTemp) * 100f) / (float)(intTHighTemp + intTLowTemp);
                    drRow[string.Format("intTHighB{0}", ints.ToString("D2"))] = intTHighTemp;
                    drRow[string.Format("intTLowB{0}", ints.ToString("D2"))] = intTLowTemp;
                    drRow[string.Format("sglPHLB{0}", ints.ToString("D2"))] = sglPHighLowTemp;
                    sglTPHLB += sglPHighLowTemp;
                    #endregion Part B
                }
                drRow["sglTPHLB"] = sglTPHLB;
                dtDataNext.Rows.Add(drRow);
            }

            #endregion 計算迴圈

            #endregion 計算下一期

            #region Show Window

            #region Expander dgData

            #region DataGrid
            DataGrid dgHighLow = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Visible,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                AutoGenerateColumns = false
            };
            #region Set Columns of DataGrid dgHighLow
            DataGridTextColumn dgtColumnHighLow;
            if (dgHighLow.Columns.Count == 0)
            {
                foreach (DataColumn dcdgColumn in dtDataTable.Columns)
                {
                    dgtColumnHighLow = new DataGridTextColumn
                    {
                        //MaxWidth = 50,
                        Header = new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1),
                        Binding = new System.Windows.Data.Binding(dcdgColumn.ColumnName),
                        IsReadOnly = true
                    };
                    string[] cName = { "intHigh", "intLow", "sglPHLA05", "sglPHLB05", "sglPHLA10", "sglPHLB10", "sglPHLA25", "sglPHLB25", "sglPHLA50", "sglPHLB50", "sglPHLA100", "sglPHLB100", "sglTPHLB" };
                    if (cName.Contains(dcdgColumn.ColumnName))
                    {
                        Style Style00 = new Style()
                        {
                            TargetType = typeof(DataGridCell)
                        };
                        Setter setter00 = new Setter()
                        {
                            Property = BackgroundProperty,
                            Value = Brushes.LightYellow
                        };
                        Style00.Setters.Add(setter00);
                        setter00 = new Setter()
                        {
                            Property = ForegroundProperty,
                            Value = Brushes.Black
                        };
                        Style00.Setters.Add(setter00);

                        #region set Trigger
                        DataTrigger dtTrig00 = new DataTrigger()
                        {
                            Binding = dgtColumnHighLow.Binding,
                            Value = "0"
                        };
                        setter00 = new Setter()
                        {
                            Property = ForegroundProperty,
                            Value = Brushes.Black
                        };
                        dtTrig00.Setters.Add(setter00);
                        setter00 = new Setter()
                        {
                            Property = BackgroundProperty,
                            Value = Brushes.GreenYellow
                        };
                        dtTrig00.Setters.Add(setter00);
                        Style00.Triggers.Add(dtTrig00);
                        #endregion set Trigger

                        dgtColumnHighLow.CellStyle = Style00;
                    }

                    dgHighLow.Columns.Add(dgtColumnHighLow);
                }
            }
            #endregion
            //dgHighLow.ColumnHeaderHeight = 50;
            dtDataTable.DefaultView.Sort = "lngTotalSN DESC ";
            dgHighLow.ItemsSource = dtDataTable.DefaultView;
            #endregion DataGrid

            Expander expHighLow = new Expander()
            {
                Header = "100 期 大數-小數資料",
                ExpandDirection = ExpandDirection.Down,
                IsExpanded = true,
                Background = Brushes.Aqua,
                VerticalAlignment = VerticalAlignment.Top,
                HorizontalAlignment = HorizontalAlignment.Left,
                MaxHeight = 400,
                MaxWidth = 1400,
                Content = dgHighLow
            };
            #endregion Expander

            #region Expander dgDataNext

            #region DataGrid
            DataGrid dgHighLowNext = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Visible,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                AutoGenerateColumns = false
            };
            #region Set Columns of DataGrid dgHighLow
            DataGridTextColumn dgtColumnHighLowext;
            if (dgHighLowNext.Columns.Count == 0)
            {
                foreach (DataColumn dcdgColumn in dtDataNext.Columns)
                {
                    dgtColumnHighLowext = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1),
                        Binding = new Binding(dcdgColumn.ColumnName),
                        IsReadOnly = true
                    };
                    string[] cName = { "intHigh", "intLow", "sglPHLA05", "sglPHLB05", "sglPHLA10", "sglPHLB10", "sglPHLA25", "sglPHLB25", "sglPHLA50", "sglPHLB50", "sglPHLA100", "sglPHLB100", "sglTPHLB" };
                    if (cName.Contains(dcdgColumn.ColumnName))
                    {
                        Style Style00 = new Style() { TargetType = typeof(DataGridCell) };
                        Setter setter00 = new Setter()
                        {
                            Property = BackgroundProperty,
                            Value = Brushes.LightYellow
                        };
                        Style00.Setters.Add(setter00);
                        setter00 = new Setter()
                        {
                            Property = ForegroundProperty,
                            Value = Brushes.Black
                        };
                        Style00.Setters.Add(setter00);

                        #region set Trigger
                        DataTrigger dtTrig00 = new DataTrigger()
                        {
                            Binding = dgtColumnHighLowext.Binding,
                            Value = "0"
                        };
                        setter00 = new Setter()
                        {
                            Property = ForegroundProperty,
                            Value = Brushes.Black
                        };
                        dtTrig00.Setters.Add(setter00);
                        setter00 = new Setter()
                        {
                            Property = BackgroundProperty,
                            Value = Brushes.GreenYellow
                        };
                        dtTrig00.Setters.Add(setter00);
                        Style00.Triggers.Add(dtTrig00);
                        #endregion set Trigger

                        dgtColumnHighLowext.CellStyle = Style00;
                    }
                    dgHighLowNext.Columns.Add(dgtColumnHighLowext);
                }
            }
            #endregion
            dgHighLowNext.ItemsSource = dtDataNext.DefaultView;
            #endregion DataGrid

            Expander expHighLowNext = new Expander()
            {
                Header = "100 期 大數-小數資料",
                ExpandDirection = ExpandDirection.Down,
                IsExpanded = true,
                Background = Brushes.Aqua,
                VerticalAlignment = VerticalAlignment.Top,
                HorizontalAlignment = HorizontalAlignment.Left,
                MaxHeight = 400,
                MaxWidth = 1400,
                Content = dgHighLowNext
            };
            #endregion Expander

            #region Expander Graphic

            #region Graphic HighLow
            Chart chtGraphic = new Chart()
            {
                Name = "gen",
                VerticalAlignment = VerticalAlignment.Stretch,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                Height = 300,
                Width = 25 * 200
            };
            ColumnSeries csGraphic = new ColumnSeries()
            {
                Title = "Highs",
                //dtDataTable.DefaultView.Sort = "[lngTotalSN] ASC";
                ItemsSource = dtDataTable.DefaultView,
                IndependentValuePath = "lngTotalSN",
                DependentValuePath = "intHigh",
                IsSelectionEnabled = true
            };
            chtGraphic.Series.Add(csGraphic);

            csGraphic = new ColumnSeries()
            {
                Title = "Lows",
                ItemsSource = dtDataTable.DefaultView,
                IndependentValuePath = "lngTotalSN",
                DependentValuePath = "intLow",
                IsSelectionEnabled = true
            };
            chtGraphic.Series.Add(csGraphic);
            #endregion Graphic

            ScrollViewer svViewerGraphic = new ScrollViewer()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = chtGraphic
            };
            Expander expHighLowGraphic = new Expander()
            {
                Header = "100 期 大數-小數資料圖表",
                ExpandDirection = ExpandDirection.Down,
                IsExpanded = false,
                Background = Brushes.Aqua,
                MaxHeight = 700,
                Content = svViewerGraphic
            };
            #endregion Expander Graphic

            #region Expander Graphic05

            //#region Graphic HighLow05
            //Chart chtGraphic05 = new Chart();
            //chtGraphic05.Name = "Total05";
            //chtGraphic05.VerticalAlignment = VerticalAlignment.Stretch;
            //chtGraphic05.HorizontalAlignment = HorizontalAlignment.Stretch;
            //chtGraphic05.Height = 300;
            //chtGraphic05.Width = 25 * 200;

            //LineSeries lsGraphic = new LineSeries();
            //lsGraphic.Title = "intTHigh05";
            //lsGraphic.ItemsSource = dtDataTable.DefaultView;
            //lsGraphic.IndependentValuePath = "lngTotalSN";
            //lsGraphic.DependentValuePath = "intTHigh05";
            //lsGraphic.IsSelectionEnabled = true;
            //chtGraphic05.Series.Add(lsGraphic);

            //lsGraphic = new LineSeries();
            //lsGraphic.Title = "intTLow05";
            //lsGraphic.ItemsSource = dtDataTable.DefaultView;
            //lsGraphic.IndependentValuePath = "lngTotalSN";
            //lsGraphic.DependentValuePath = "intTLow05";
            //lsGraphic.IsSelectionEnabled = true;
            //chtGraphic05.Series.Add(lsGraphic);

            //lsGraphic = new LineSeries();
            //lsGraphic.Title = "sglPHighLow05";
            //lsGraphic.ItemsSource = dtDataTable.DefaultView;
            //lsGraphic.IndependentValuePath = "lngTotalSN";
            //lsGraphic.DependentValuePath = "sglPHighLow05";
            //lsGraphic.IsSelectionEnabled = true;
            //chtGraphic05.Series.Add(lsGraphic);
            //#endregion Graphic HighLow05
            //ScrollViewer svViewerGraphi05 = new ScrollViewer();
            //svViewerGraphi05.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
            //svViewerGraphi05.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            //svViewerGraphi05.Content = chtGraphic05;

            //Expander expHighLowGraphic05 = new Expander();
            //expHighLowGraphic05.Header = "100 期 奇數偶數圖表05";
            //expHighLowGraphic05.ExpandDirection = ExpandDirection.Down;
            //expHighLowGraphic05.IsExpanded = false;
            //expHighLowGraphic05.Background = Brushes.Aqua;
            //expHighLowGraphic05.MaxHeight = 700;
            //expHighLowGraphic05.Content = svViewerGraphi05;

            #endregion Expander Graphic05

            #region Expander Graphic10
            //#region Graphic HighLow10
            //Chart chtGraphic10 = new Chart();
            //chtGraphic10.Name = "Total0";
            //chtGraphic10.VerticalAlignment = VerticalAlignment.Stretch;
            //chtGraphic10.HorizontalAlignment = HorizontalAlignment.Stretch;
            //chtGraphic10.Height = 300;
            //chtGraphic10.Width = 25 * 200;

            //lsGraphic = new LineSeries();
            //lsGraphic.Title = "intTHigh10";
            //lsGraphic.ItemsSource = dtDataTable.DefaultView;
            //lsGraphic.IndependentValuePath = "lngTotalSN";
            //lsGraphic.DependentValuePath = "intTHigh10";
            //lsGraphic.IsSelectionEnabled = true;
            //chtGraphic10.Series.Add(lsGraphic);

            //lsGraphic = new LineSeries();
            //lsGraphic.Title = "intTLow10";
            //lsGraphic.ItemsSource = dtDataTable.DefaultView;
            //lsGraphic.IndependentValuePath = "lngTotalSN";
            //lsGraphic.DependentValuePath = "intTLow10";
            //lsGraphic.IsSelectionEnabled = true;
            //chtGraphic10.Series.Add(lsGraphic);

            //lsGraphic = new LineSeries();
            //lsGraphic.Title = "sglPHighLow10";
            //lsGraphic.ItemsSource = dtDataTable.DefaultView;
            //lsGraphic.IndependentValuePath = "lngTotalSN";
            //lsGraphic.DependentValuePath = "sglPHighLow10";
            //lsGraphic.IsSelectionEnabled = true;
            //chtGraphic10.Series.Add(lsGraphic);
            //#endregion Graphic HighLow10
            //ScrollViewer svViewerGraphi10 = new ScrollViewer();
            //svViewerGraphi10.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
            //svViewerGraphi10.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            //svViewerGraphi10.Content = chtGraphic10;

            //Expander expHighLowGraphic10 = new Expander();
            //expHighLowGraphic10.Header = "100 期 奇數偶數圖表10";
            //expHighLowGraphic10.ExpandDirection = ExpandDirection.Down;
            //expHighLowGraphic10.IsExpanded = false;
            //expHighLowGraphic10.Background = Brushes.Aqua;
            //expHighLowGraphic10.MaxHeight = 700;
            //expHighLowGraphic10.Content = svViewerGraphi10;
            #endregion Expander Graphic10

            #region Expander Graphic25
            //#region Graphic HighLow25
            //Chart chtGraphic25 = new Chart();
            //chtGraphic25.Name = "Total0";
            //chtGraphic25.VerticalAlignment = VerticalAlignment.Stretch;
            //chtGraphic25.HorizontalAlignment = HorizontalAlignment.Stretch;
            //chtGraphic25.Height = 300;
            //chtGraphic25.Width = 25 * 200;

            //lsGraphic = new LineSeries();
            //lsGraphic.Title = "intTHigh25";
            //lsGraphic.ItemsSource = dtDataTable.DefaultView;
            //lsGraphic.IndependentValuePath = "lngTotalSN";
            //lsGraphic.DependentValuePath = "intTHigh25";
            //lsGraphic.IsSelectionEnabled = true;
            //chtGraphic25.Series.Add(lsGraphic);

            //lsGraphic = new LineSeries();
            //lsGraphic.Title = "intTLow25";
            //lsGraphic.ItemsSource = dtDataTable.DefaultView;
            //lsGraphic.IndependentValuePath = "lngTotalSN";
            //lsGraphic.DependentValuePath = "intTLow25";
            //lsGraphic.IsSelectionEnabled = true;
            //chtGraphic25.Series.Add(lsGraphic);

            //lsGraphic = new LineSeries();
            //lsGraphic.Title = "sglPHighLow25";
            //lsGraphic.ItemsSource = dtDataTable.DefaultView;
            //lsGraphic.IndependentValuePath = "lngTotalSN";
            //lsGraphic.DependentValuePath = "sglPHighLow25";
            //lsGraphic.IsSelectionEnabled = true;
            //chtGraphic25.Series.Add(lsGraphic);
            //#endregion Graphic HighLow25
            //ScrollViewer svViewerGraphi25 = new ScrollViewer();
            //svViewerGraphi25.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
            //svViewerGraphi25.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            //svViewerGraphi25.Content = chtGraphic25;

            //Expander expHighLowGraphic25 = new Expander();
            //expHighLowGraphic25.Header = "100 期 奇數偶數圖表25";
            //expHighLowGraphic25.ExpandDirection = ExpandDirection.Down;
            //expHighLowGraphic25.IsExpanded = false;
            //expHighLowGraphic25.Background = Brushes.Aqua;
            //expHighLowGraphic25.MaxHeight = 700;
            //expHighLowGraphic25.Content = svViewerGraphi25;
            #endregion Expander Graphic25

            #region StackPanel
            StackPanel spMain = new StackPanel()
            {
                CanVerticallyScroll = true,
                CanHorizontallyScroll = true
            };
            spMain.Children.Add(expHighLow);
            spMain.Children.Add(expHighLowNext);
            spMain.Children.Add(expHighLowGraphic);
            //spMain.Children.Add(expHighLowGraphic05);
            //spMain.Children.Add(expHighLowGraphic10);
            //spMain.Children.Add(expHighLowGraphic25);
            #endregion StackPanel

            #region ScrollViewer
            ScrollViewer svViewer = new ScrollViewer()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = spMain
            };
            #endregion ScrollViewer

            #region Window
            Window wHighLow = new Window()
            {
                Name = "tblHighLow",
                Title = string.Format("大數-小數表 ({0})", stuSearchHighLow.StrTitle),
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.SingleBorderWindow,
                Content = svViewer
            };
            wHighLow.Show();
            #endregion Window

            #endregion Show Window
        }

        #region btnTableSum 和數值表
        private void BtnTableSum_Click(object sender, RoutedEventArgs e)
        {
            #region 設定 stuSearchSum
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSearchSum = stuRibbonSearchOption;
            #endregion 設定 stuSearchSum

            #region Setting BwSum
            bwBackgroundWorker00 = new BackgroundWorker()
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };
            bwBackgroundWorker00.DoWork += BwSum_DoWork;
            bwBackgroundWorker00.ProgressChanged += BwSum_ProgressChanged;
            bwBackgroundWorker00.RunWorkerCompleted += BwSum_RunWorkerCompleted;
            #endregion Setting bwTablePercent

            bwBackgroundWorker00.RunWorkerAsync(stuSearchSum);
            btnTableSum.Content = "計算中...";
            btnTableSum.IsEnabled = false;

        }
        private void BwSum_DoWork(object sender, DoWorkEventArgs e)
        {
            #region Argument Setup
            StuGLSearch stuSearchSum = (StuGLSearch)e.Argument;
            BackgroundWorker bwMyWork = (BackgroundWorker)sender;
            bwMyWork.ReportProgress(1);
            Dictionary<string, object> dicArgument = new Dictionary<string, object>
            {
                { "stuSearch", stuSearchSum }
            };
            CGLDataSet Dataset00 = new CGLDataSet(stuSearchSum.LottoType);
            #endregion Setup

            #region Check tblMissAll ,Does it have Data?
            if (!new CGLSum().HasSumData(stuSearchSum))
            {
                new CGLSum().GetSumDic(stuSearchSum);
            }
            #endregion 檢查是否有資料

            int[] intSections = { 5, 10, 25, 50, 100 };

            DataTable dtDataTable, dtDataTemp = new DataTable();
            #region 導入 200 期 數值
            #region Query String
            Dictionary<string, string> dicCurrentData = new Dictionary<string, string>();
            dicCurrentData = new CGLSearch().GetCurrentData(stuSearchSum);
            StuGLData stuData00 = new StuGLData()
            {
                LottoType = stuSearchSum.LottoType,
                DataBaseType = DatabaseType.Sum,
                StrSELECT = " TOP 200 ",
                StrFROM = "[qrySum] ",
                StrWHERE = string.Format("[{0}] = {1} AND [{2}] < {3} ",
                                        "lngMethodSN", stuSearchSum.LngMethodSN,
                                        "lngTotalSN", stuSearchSum.LngCurrentData),
                StrORDER = "[lngTotalSN] DESC ;"
            };

            #region set select 
            stuData00.StrSELECT += "[lngTotalSN] , [lngMethodSN] , [lngDateSN] , ";
            for (int i = 1; i <= Dataset00.CountNumber; i++)
            {
                stuData00.StrSELECT += string.Format("[lngL{0}] , ", i.ToString());
            }
            if (Dataset00.DataType != TableType.LottoWeli && Dataset00.SCountNumber > 0)
            {
                stuData00.StrSELECT += "[lngS1] , ";
            }
            stuData00.StrSELECT += "[intSum] , [sglAvg] , [sglTPA] , ";
            foreach (int ints in intSections)
            {
                stuData00.StrSELECT += string.Format("[sglAvg{0}] , [sglPA{0}] ,", ints.ToString("D2"));
            }
            stuData00.StrSELECT = stuData00.StrSELECT.TrimEnd(',');
            #endregion

            #region  Set Compare Option 
            if (stuSearchSum.BoolFieldMode && stuSearchSum.StrCompares != "gen")
            {
                string[] strCompare = stuSearchSum.StrCompares.Split('#');
                foreach (string strCompareOption in strCompare)
                {
                    stuData00.StrWHERE += string.Format(" {0} [{1}] = '{2}' "
                                                        , stuSearchSum.StrCompareType
                                                        , strCompareOption
                                                        , dicCurrentData[strCompareOption]);
                }
            }
            #endregion

            #endregion Query String
            dtDataTable = stuData00.GetSourceData(stuData00);
            #endregion 導入 200 期 數值

            #region 計算下一期

            #region 建立 dtDataNext
            DataTable dtDataNext = new DataTable();

            #region 建立 Columns
            DataColumn dcColumn = new DataColumn()
            {
                ColumnName = "intSum",
                Caption = "intSum",
                DataType = typeof(int)
            };
            dtDataNext.Columns.Add(dcColumn);

            dcColumn = new DataColumn()
            {
                ColumnName = "sglAvg",
                Caption = "sglAvg",
                DataType = typeof(float)
            };
            dtDataNext.Columns.Add(dcColumn);

            dcColumn = new DataColumn()
            {
                ColumnName = "sglTPA",
                Caption = "sglTPA",
                DataType = typeof(float)
            };
            dtDataNext.Columns.Add(dcColumn);

            foreach (int ints in intSections)
            {
                dcColumn = new DataColumn()
                {
                    ColumnName = string.Format("sglAvg{0}", ints.ToString("D2")),
                    Caption = string.Format("sglAvg{0}", ints.ToString("D2")),
                    DataType = typeof(float)
                };
                dtDataNext.Columns.Add(dcColumn);

                dcColumn = new DataColumn()
                {
                    ColumnName = string.Format("sglPA{0}", ints.ToString("D2")),
                    Caption = string.Format("sglPA{0}", ints.ToString("D2")),
                    DataType = typeof(float)
                };
                dtDataNext.Columns.Add(dcColumn);
            }

            #endregion 建立 Columns

            #endregion 建立 tblDataNext

            #region 計算迴圈
            int intStar, intEnd, intAvg, intMax, intMin;
            intMin = (1 + Dataset00.CountNumber) * Dataset00.CountNumber / 2;
            intMax = (2 * Dataset00.LottoNumbers - Dataset00.CountNumber + 1) * Dataset00.CountNumber / 2;
            intAvg = (intMax + intMin) / 2;
            intStar = intAvg - (intAvg * 2 / 3);
            intEnd = intAvg + (intAvg * 2 / 3);
            for (int i = intStar; i <= intEnd; i++)
            {
                double dblAvgTemp, dblTotalTemp, intSum, sglAvg;
                DataRow drRow = dtDataNext.NewRow();
                int intSumTemp = i;
                drRow["intSum"] = intSumTemp;

                #region 計算 sglAvg
                #region Query String
                stuData00 = new StuGLData()
                {
                    LottoType = stuSearchSum.LottoType,
                    DataBaseType = DatabaseType.Data,
                    StrSELECT = " * ",
                    StrFROM = string.Format("[{0}]", Dataset00.QueryData),
                    StrWHERE = string.Format("[lngTotalSN] < {0} ", stuSearchSum.LngCurrentData),
                    StrORDER = "[lngTotalSN] DESC ;"
                };
                #region  Set Compare Option 
                if (stuSearchSum.BoolFieldMode && stuSearchSum.StrCompares != "gen")
                {
                    string[] strCompare = stuSearchSum.StrCompares.Split('#');
                    foreach (string strCompareOption in strCompare)
                    {
                        stuData00.StrWHERE += string.Format(" {0} [{1}] = '{2}' "
                                                            , stuSearchSum.StrCompareType
                                                            , strCompareOption
                                                            , dicCurrentData[strCompareOption]);
                    }
                }
                #endregion
                #endregion Query String
                dtDataTemp = stuData00.GetSourceData(stuData00);

                dblAvgTemp = double.Parse(dtDataTemp.Compute(string.Format("Avg([{0}])", "intSum"), string.Empty).ToString());
                dblTotalTemp = dblAvgTemp * (double)dtDataTemp.Rows.Count;
                dblAvgTemp = (dblTotalTemp + (double)intSumTemp) / (double)(dtDataTemp.Rows.Count + 1);
                drRow["sglAvg"] = Math.Round(double.Parse(dblAvgTemp.ToString()), 2);
                intSum = double.Parse(drRow["intSum"].ToString());
                sglAvg = double.Parse(drRow["sglAvg"].ToString());
                drRow["sglTPA"] = Math.Round(double.Parse((intSum - sglAvg).ToString()), 2);
                #endregion 計算 sglAvg

                #region 計算 sglAvg 05-100
                foreach (int Section in intSections)
                {
                    #region Query String
                    stuData00 = new StuGLData()
                    {
                        LottoType = stuSearchSum.LottoType,
                        DataBaseType = DatabaseType.Data,
                        StrSELECT = string.Format(" TOP {0} * ", (Section - 1).ToString("D2")),
                        StrFROM = string.Format("[{0}]", Dataset00.QueryData),
                        StrWHERE = string.Format("[lngTotalSN] < {0} ", stuSearchSum.LngCurrentData),
                        StrORDER = "[lngTotalSN] DESC ;"
                    };
                    #region  Set Compare Option 
                    if (stuSearchSum.BoolFieldMode && stuSearchSum.StrCompares != "gen")
                    {
                        string[] strCompare = stuSearchSum.StrCompares.Split('#');
                        foreach (string strCompareOption in strCompare)
                        {
                            stuData00.StrWHERE += string.Format(" {0} [{1}] = '{2}' "
                                                                , stuSearchSum.StrCompareType
                                                                , strCompareOption
                                                                , dicCurrentData[strCompareOption]);
                        }
                    }
                    #endregion
                    #endregion Query String
                    dtDataTemp = stuData00.GetSourceData(stuData00);

                    dblAvgTemp = double.Parse(dtDataTemp.Compute(string.Format("Avg([{0}])", "intSum"), string.Empty).ToString());
                    dblTotalTemp = dblAvgTemp * (double)(Section - 1);
                    dblAvgTemp = (dblTotalTemp + (double)intSumTemp) / (double)(Section);

                    string FiledAvg = string.Format("sglAvg{0}", Section.ToString("D2"));
                    string FiledPA = string.Format("sglPA{0}", Section.ToString("D2"));

                    drRow[FiledAvg] = Math.Round(double.Parse(dblAvgTemp.ToString()), 2);
                    intSum = double.Parse(drRow["intSum"].ToString());
                    sglAvg = double.Parse(drRow[FiledAvg].ToString());
                    drRow[FiledPA] = Math.Round(double.Parse((intSum - sglAvg).ToString()), 2);
                }
                #endregion 計算 sglAvg 05-100

                dtDataNext.Rows.Add(drRow);
            }

            #endregion 計算迴圈

            #endregion 計算下一期

            DataSet dsSum = new DataSet();
            dtDataTable.TableName = "TableSum";
            dsSum.Tables.Add(dtDataTable);
            dtDataNext.TableName = "DataNext";
            dsSum.Tables.Add(dtDataNext);
            dicArgument.Add("Freq", dsSum);
            e.Result = dicArgument;

        }
        private void BwSum_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 1)
            {
                this.tbStatusTextBlock.Text = "計算中...";
            }
        }
        private void BwSum_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            BackgroundWorker bwMywork = (BackgroundWorker)sender;
            Dictionary<string, object> dicArgument = (Dictionary<string, object>)e.Result;
            ShowSum((StuGLSearch)dicArgument["stuSearch"], (DataSet)dicArgument["Freq"]);
            tbStatusTextBlock.Text = "Ready";
            btnTableSum.Content = "和數值表";
            btnTableSum.IsEnabled = true;
            bwMywork.Dispose();
        }
        private void ShowSum(StuGLSearch stuSearch00, DataSet dsSum)
        {
            DataTable dtDataTable = dsSum.Tables["TableSum"];
            DataTable dtDataNext = dsSum.Tables["DataNext"];
            Dictionary<string, string> dicCurrentData = new CGLSearch().GetCurrentData(stuSearch00);
            #region Show Window

            #region Set DataGrid dgCurrentData
            #region Convert dictionary dicCurrentData to table dtCurrentData
            DataTable dtCurrentData = new CGLFunc().CDicTOTable(dicCurrentData);
            #endregion
            DataGrid dgCurrentData = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                RowDetailsVisibilityMode = DataGridRowDetailsVisibilityMode.Visible,
                CanUserAddRows = true,
                CanUserDeleteRows = false,
                CanUserResizeColumns = true,
                CanUserResizeRows = true,
                CanUserSortColumns = true,
                IsReadOnly = true,
                AutoGenerateColumns = false,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch
            };
            #region Set Columns of DataGrid dgCurrentData 
            DataGridTextColumn dgtColumn;
            if (dgCurrentData.Columns.Count == 0)
            {
                foreach (var KeyPair in dicCurrentData)
                {
                    dgtColumn = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(KeyPair.Key, 1),
                        Binding = new Binding(KeyPair.Key),
                        IsReadOnly = true
                    };
                    dgCurrentData.Columns.Add(dgtColumn);
                }
            }
            #endregion
            dgCurrentData.ItemsSource = dtCurrentData.DefaultView;
            #endregion

            #region Expander dgData

            #region DataGrid
            DataGrid dgSum = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                AutoGenerateColumns = false
            };
            #region Set Columns of DataGrid dgSum
            DataGridTextColumn dgtColumnSum;
            if (dgSum.Columns.Count == 0)
            {
                foreach (DataColumn dcdgColumn in dtDataTable.Columns)
                {
                    dgtColumnSum = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1),
                        Binding = new Binding(dcdgColumn.ColumnName),
                        IsReadOnly = true
                    };
                    string[] cName = { "intSum", "sglAvg", "sglAvg05", "sglAvg10", "sglAvg25", "sglAvg50", "sglAvg100" };
                    if (cName.Contains(dcdgColumn.ColumnName))
                    {
                        Style Style00 = new Style()
                        {
                            TargetType = typeof(DataGridCell)
                        };
                        Setter setter00 = new Setter()
                        {
                            Property = BackgroundProperty,
                            Value = Brushes.LightYellow
                        };
                        Style00.Setters.Add(setter00);
                        setter00 = new Setter()
                        {
                            Property = ForegroundProperty,
                            Value = Brushes.Black
                        };
                        Style00.Setters.Add(setter00);

                        #region set Trigger
                        DataTrigger dtTrig00 = new DataTrigger()
                        {
                            Binding = dgtColumnSum.Binding,
                            Value = "0"
                        };
                        setter00 = new Setter()
                        {
                            Property = ForegroundProperty,
                            Value = Brushes.Black
                        };
                        dtTrig00.Setters.Add(setter00);
                        setter00 = new Setter()
                        {
                            Property = BackgroundProperty,
                            Value = Brushes.GreenYellow
                        };
                        dtTrig00.Setters.Add(setter00);
                        Style00.Triggers.Add(dtTrig00);
                        #endregion set Trigger

                        dgtColumnSum.CellStyle = Style00;
                    }
                    dgSum.Columns.Add(dgtColumnSum);
                }
            }
            #endregion
            dgSum.ItemsSource = dtDataTable.DefaultView;
            dgSum.LoadingRow += DgSum_LoadingRow;
            #endregion DataGrid

            Expander expSum = new Expander()
            {
                Header = "200 期 和數值資料",
                ExpandDirection = ExpandDirection.Down,
                IsExpanded = true,
                Background = Brushes.Aqua,
                VerticalAlignment = VerticalAlignment.Top,
                HorizontalAlignment = HorizontalAlignment.Left,
                MaxHeight = 400,
                MaxWidth = 1400,
                Content = dgSum
            };

            #endregion Expander

            #region Expander dgDataNext

            #region DataGrid
            DataGrid dgSumNext = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Visible,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                AutoGenerateColumns = false
            };
            #region Set Columns of DataGrid dgSum
            DataGridTextColumn dgtColumnSumext;
            if (dgSumNext.Columns.Count == 0)
            {
                foreach (DataColumn dcdgColumn in dtDataNext.Columns)
                {
                    dgtColumnSumext = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1),
                        Binding = new System.Windows.Data.Binding(dcdgColumn.ColumnName),
                        IsReadOnly = true
                    };
                    string[] cName = { "sglAvg", "sglAvg05", "sglAvg10", "sglAvg25", "sglAvg50", "sglAvg100" };
                    if (cName.Contains(dcdgColumn.ColumnName))
                    {
                        Style Style00 = new Style()
                        {
                            TargetType = typeof(DataGridCell)
                        };
                        #region set Trigger
                        DataTrigger dtTrig00 = new DataTrigger()
                        {
                            Binding = dgtColumnSumext.Binding,
                            Value = dtDataTable.Rows[0][dcdgColumn.ColumnName].ToString()
                        };
                        Setter setter00 = new Setter()
                        {
                            Property = ForegroundProperty,
                            Value = Brushes.Black
                        };
                        dtTrig00.Setters.Add(setter00);
                        setter00 = new Setter()
                        {
                            Property = BackgroundProperty,
                            Value = Brushes.GreenYellow
                        };
                        dtTrig00.Setters.Add(setter00);
                        Style00.Triggers.Add(dtTrig00);
                        #endregion set Trigger
                        dgtColumnSumext.CellStyle = Style00;
                    }
                    dgSumNext.Columns.Add(dgtColumnSumext);
                }
            }
            #endregion
            dgSumNext.ItemsSource = dtDataNext.DefaultView;
            #endregion DataGrid

            Expander expSumNext = new Expander()
            {
                Header = "100 期 和數值資料",
                ExpandDirection = ExpandDirection.Down,
                IsExpanded = true,
                Background = Brushes.Aqua,
                VerticalAlignment = VerticalAlignment.Top,
                HorizontalAlignment = HorizontalAlignment.Left,
                MaxHeight = 400,
                MaxWidth = 1400,
                Content = dgSumNext
            };
            #endregion Expander

            #region Expander Graphic

            #region Graphic Sum
            Chart chtGraphic = new Chart()
            {
                Name = "gen",
                VerticalAlignment = VerticalAlignment.Stretch,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                Height = 300,
                Width = 25 * 200
            };
            ColumnSeries csGraphic = new ColumnSeries()
            {
                Title = "Sum",
                //dtDataTable.DefaultView.Sort = "[lngTotalSN] ASC";
                ItemsSource = dtDataTable.DefaultView,
                IndependentValuePath = "lngTotalSN",
                DependentValuePath = "intSum",
                IsSelectionEnabled = true
            };
            chtGraphic.Series.Add(csGraphic);

            csGraphic = new ColumnSeries()
            {
                Title = "Avg",
                ItemsSource = dtDataTable.DefaultView,
                IndependentValuePath = "lngTotalSN",
                DependentValuePath = "sglAvg",
                IsSelectionEnabled = true
            };
            chtGraphic.Series.Add(csGraphic);

            //LineSeries lsGraphic = new LineSeries();
            //lsGraphic.Title = "sglAvg05";
            //lsGraphic.ItemsSource = dtDataTable.DefaultView;
            //lsGraphic.IndependentValuePath = "lngTotalSN";
            //lsGraphic.DependentValuePath = "sglAvg05";
            //lsGraphic.IsSelectionEnabled = true;
            //chtGraphic.Series.Add(lsGraphic);

            //lsGraphic = new LineSeries();
            //lsGraphic.Title = "sglAvg10";
            //lsGraphic.ItemsSource = dtDataTable.DefaultView;
            //lsGraphic.IndependentValuePath = "lngTotalSN";
            //lsGraphic.DependentValuePath = "sglAvg10";
            //lsGraphic.IsSelectionEnabled = true;
            //chtGraphic.Series.Add(lsGraphic);

            //lsGraphic = new LineSeries();
            //lsGraphic.Title = "sglAvg25";
            //lsGraphic.ItemsSource = dtDataTable.DefaultView;
            //lsGraphic.IndependentValuePath = "lngTotalSN";
            //lsGraphic.DependentValuePath = "sglAvg25";
            //lsGraphic.IsSelectionEnabled = true;
            //chtGraphic.Series.Add(lsGraphic);

            //lsGraphic = new LineSeries();
            //lsGraphic.Title = "sglAvg50";
            //lsGraphic.ItemsSource = dtDataTable.DefaultView;
            //lsGraphic.IndependentValuePath = "lngTotalSN";
            //lsGraphic.DependentValuePath = "sglAvg50";
            //lsGraphic.IsSelectionEnabled = true;
            //chtGraphic.Series.Add(lsGraphic);

            //lsGraphic = new LineSeries();
            //lsGraphic.Title = "sglAvg100";
            //lsGraphic.ItemsSource = dtDataTable.DefaultView;
            //lsGraphic.IndependentValuePath = "lngTotalSN";
            //lsGraphic.DependentValuePath = "sglAvg100";
            //lsGraphic.IsSelectionEnabled = true;
            //chtGraphic.Series.Add(lsGraphic);

            #endregion Graphic

            //ScrollViewer svViewerGraphic = new ScrollViewer();
            //svViewerGraphic.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
            //svViewerGraphic.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            //svViewerGraphic.Content = chtGraphic;


            Expander expSumGraphic = new Expander()
            {
                Header = "100 期 和數值圖表",
                ExpandDirection = ExpandDirection.Down,
                IsExpanded = true,
                Background = Brushes.Aqua,
                MaxHeight = 700,
                //expSumGraphic.Content = svViewerGraphic;
                Content = chtGraphic
            };
            #endregion Expander Graphic

            #region StackPanel
            StackPanel spMain = new StackPanel()
            {
                CanVerticallyScroll = true,
                CanHorizontallyScroll = true
            };
            spMain.Children.Add(dgCurrentData);
            spMain.Children.Add(expSum);
            spMain.Children.Add(expSumNext);
            spMain.Children.Add(expSumGraphic);
            #endregion StackPanel

            #region ScrollViewer
            ScrollViewer svViewer = new ScrollViewer()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = spMain
            };
            #endregion ScrollViewer

            #region Window
            Window wSum = new Window()
            {
                Name = "tblSum",
                Title = string.Format("和數值資料表 ({0})", stuSearch00.StrTitle),
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.SingleBorderWindow,
                Content = svViewer
            };
            wSum.Show();
            #endregion Window

            #endregion Show Window

        }
        private void DgSum_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            DataGridRow dgrRow = e.Row;
            DataRowView drvGridRow = (DataRowView)dgrRow.Item;
            double intSum = double.Parse(drvGridRow.Row.ItemArray[8].ToString());
            double intAvg = double.Parse(drvGridRow.Row.ItemArray[9].ToString());
            if (intAvg > intSum)
            {
                e.Row.Background = Brushes.Honeydew;
            }
            if (intAvg < intSum)
            {
                e.Row.Background = Brushes.LavenderBlush;
            }
            if (intAvg == intSum)
            {
                e.Row.Background = Brushes.PaleGoldenrod;
            }
        }

        #endregion btnTableSum

        #region btnTableInterval 數字區間表
        private void BtnTableInterval_Click(object sender, RoutedEventArgs e)
        {
            #region 設定 stuSearchInterval
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSearchInterval = stuRibbonSearchOption;
            #endregion 設定 stuSearchInterval
            #region Setting bwTablePercent
            bwBackgroundWorker00 = new BackgroundWorker()
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };
            bwBackgroundWorker00.DoWork += BwInterval_DoWork;
            bwBackgroundWorker00.ProgressChanged += BwInterval_ProgressChanged;
            bwBackgroundWorker00.RunWorkerCompleted += BwInterval_RunWorkerCompleted;
            #endregion Setting bwTablePercent
            bwBackgroundWorker00.RunWorkerAsync(stuSearchInterval);
            btnTableInterval.Content = "計算中...";
            btnTableInterval.IsEnabled = false;
        }
        private void BwInterval_DoWork(object sender, DoWorkEventArgs e)
        {
            StuGLSearch stuSearchInterval = (StuGLSearch)e.Argument;
            BackgroundWorker bwMyWork = (BackgroundWorker)sender;
            bwMyWork.ReportProgress(1);
            Dictionary<string, object> dicArgument = new Dictionary<string, object>
            {
                { "stuSearch", stuSearchInterval }
            };
            CGLDataSet Dataset00 = new CGLDataSet(stuSearchInterval.LottoType);
            Dictionary<string, string> dicCurrentData = new CGLSearch().GetCurrentData(stuSearchInterval);

            #region 建立分區
            int intInterval01, intCountNum;
            if (Dataset00.DataType != TableType.LottoWeli)
            {
                intCountNum = Dataset00.CountNumber + Dataset00.SCountNumber;
                intInterval01 = Dataset00.LottoNumbers / intCountNum;
                if (intInterval01 * intCountNum < Dataset00.LottoNumbers) intInterval01++;
            }
            else
            {
                intCountNum = Dataset00.CountNumber;
                intInterval01 = Dataset00.LottoNumbers / intCountNum;
                if (intInterval01 * intCountNum < Dataset00.LottoNumbers) intInterval01++;
            }

            #endregion 建立分區

            DataSet dsIntervals = new DataSet();
            DataTable dtDataInterval = SearchInterval(stuSearchInterval, intInterval01);
            dtDataInterval.TableName = "Interval";
            dsIntervals.Tables.Add(dtDataInterval);
            DataTable dtDataInterval10 = SearchInterval(stuSearchInterval, 10);
            dtDataInterval10.TableName = "Interval10";
            dsIntervals.Tables.Add(dtDataInterval10);
            dicArgument.Add("Freq", dsIntervals);
            e.Result = dicArgument;

        }
        private void BwInterval_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 1)
            {
                this.tbStatusTextBlock.Text = "計算中...";
            }
        }
        private void BwInterval_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            BackgroundWorker bwMywork = (BackgroundWorker)sender;
            Dictionary<string, object> dicArgument = (Dictionary<string, object>)e.Result;
            ShowTableInterval((StuGLSearch)dicArgument["stuSearch"], (DataSet)dicArgument["Freq"]);
            tbStatusTextBlock.Text = "Ready";
            btnTableInterval.Content = "數字區間表";
            btnTableInterval.IsEnabled = true;
            bwMywork.Dispose();
        }
        private DataTable SearchInterval(StuGLSearch stuSearchInterval, int intInterval01)
        {
            CGLDataSet Dataset00 = new CGLDataSet(stuSearchInterval.LottoType);

            #region 導入 100 期 數值
            Dictionary<string, string> dicCurrentData = new Dictionary<string, string>();
            dicCurrentData = new CGLSearch().GetCurrentData(stuSearchInterval);

            #region Query String
            StuGLData stuData00 = new StuGLData()
            {
                LottoType = stuSearchInterval.LottoType,
                DataBaseType = DatabaseType.Data,
                StrSELECT = "TOP 100 * ",
                StrFROM = string.Format("[{0}]", Dataset00.QueryData),
                StrWHERE = string.Format("[lngTotalSN] >= {0} And [lngTotalSN] < {1} "
                                                                , stuSearchInterval.LngDataStart
                                                                , stuSearchInterval.LngCurrentData),
                StrORDER = "[lngTotalSN] DESC ;"
            };
            #region  Set Compare Option 
            if (stuSearchInterval.BoolFieldMode && stuSearchInterval.StrCompares != "gen")
            {
                string[] strCompare = stuSearchInterval.StrCompares.Split('#');
                foreach (string strCompareOption in strCompare)
                {
                    stuData00.StrWHERE += string.Format(" {0} [{1}] = '{2}' "
                                                        , stuSearchInterval.StrCompareType
                                                        , strCompareOption
                                                        , dicCurrentData[strCompareOption]);
                }
            }
            #endregion


            #endregion Query String

            DataTable dtDataTable = stuData00.GetSourceData(stuData00);
            #endregion 導入 100 期 數值

            int intIntervalNum = Dataset00.LottoNumbers / intInterval01;
            if (intIntervalNum * intInterval01 < Dataset00.LottoNumbers) intIntervalNum++;
            #region 建立 Table dtDataInterval 
            DataTable dtDataInterval = new DataTable();

            #region Create Columns
            DataColumn dcColumn = new DataColumn();
            string strFieldName = "lngTotalSN";
            dcColumn.ColumnName = dtDataTable.Columns[strFieldName].ColumnName;
            dcColumn.Caption = dtDataTable.Columns[strFieldName].Caption;
            dcColumn.DataType = dtDataTable.Columns[strFieldName].DataType;
            dtDataInterval.Columns.Add(dcColumn);

            dcColumn = new DataColumn();
            strFieldName = "lngDateSN";
            dcColumn.ColumnName = dtDataTable.Columns[strFieldName].ColumnName;
            dcColumn.Caption = dtDataTable.Columns[strFieldName].Caption;
            dcColumn.DataType = dtDataTable.Columns[strFieldName].DataType;
            dtDataInterval.Columns.Add(dcColumn);

            for (int i = 1; i <= Dataset00.CountNumber; i++)
            {
                dcColumn = new DataColumn();
                strFieldName = string.Format("lngL{0}", i.ToString());
                dcColumn.ColumnName = dtDataTable.Columns[strFieldName].ColumnName;
                dcColumn.Caption = dtDataTable.Columns[strFieldName].Caption;
                dcColumn.DataType = dtDataTable.Columns[strFieldName].DataType;
                dtDataInterval.Columns.Add(dcColumn);
            }

            if (Dataset00.SCountNumber > 0)
            {
                dcColumn = new DataColumn();
                strFieldName = "lngS1";
                dcColumn.ColumnName = dtDataTable.Columns[strFieldName].ColumnName;
                dcColumn.Caption = dtDataTable.Columns[strFieldName].Caption;
                dcColumn.DataType = dtDataTable.Columns[strFieldName].DataType;
                dtDataInterval.Columns.Add(dcColumn);
            }

            for (int i = 1; i <= intIntervalNum; i++)
            {
                dcColumn = new DataColumn();
                strFieldName = string.Format("intInt{0}", i.ToString());
                dcColumn.ColumnName = strFieldName;
                string strFieldCaption = "";
                if (Dataset00.DataType != TableType.LottoWeli)
                {
                    int intIntStar = (i - 1) * intInterval01;
                    if (intIntStar <= 0) intIntStar = 1;
                    int intIntEnd = (i * intInterval01) - 1;
                    if (intIntEnd > Dataset00.LottoNumbers) intIntEnd = Dataset00.LottoNumbers;
                    strFieldCaption = string.Format("{0}-{1}", intIntStar.ToString("D2"), intIntEnd.ToString("D2"));
                }
                dcColumn.Caption = strFieldCaption;
                dcColumn.DataType = typeof(Int16);
                dcColumn.DefaultValue = 0;
                dtDataInterval.Columns.Add(dcColumn);
            }
            #endregion Create Columns

            #region 計算 Interval
            foreach (DataRow drRow in dtDataTable.Rows)
            {
                DataRow drNewRow = dtDataInterval.NewRow();
                drNewRow["lngTotalSN"] = drRow["lngTotalSN"];
                drNewRow["lngDateSN"] = drRow["lngDateSN"];
                for (int L = 1; L <= Dataset00.CountNumber; L++)
                {
                    strFieldName = string.Format("lngL{0}", L.ToString());
                    drNewRow[strFieldName] = drRow[strFieldName];
                    for (int i01 = 1; i01 <= intIntervalNum; i01++)
                    {
                        string strIntFieldName = string.Format("intInt{0}", i01.ToString());
                        if (Dataset00.DataType != TableType.LottoWeli)
                        {
                            int intIntStar = (i01 - 1) * intInterval01 + 1;
                            int intIntEnd = i01 * intInterval01;
                            int intCheckNum = int.Parse(drRow[strFieldName].ToString());
                            if (intIntEnd > Dataset00.LottoNumbers) intIntEnd = Dataset00.LottoNumbers;
                            if (intCheckNum >= intIntStar && intCheckNum <= intIntEnd)
                            {
                                drNewRow[strIntFieldName] = int.Parse(drNewRow[strIntFieldName].ToString()) + 1;
                            }
                        }
                    }
                }
                if (Dataset00.SCountNumber > 0)
                {
                    strFieldName = "lngS1";
                    drNewRow[strFieldName] = drRow[strFieldName];
                    for (int i01 = 1; i01 <= intIntervalNum; i01++)
                    {
                        string strIntFieldName = string.Format("intInt{0}", i01.ToString());
                        if (Dataset00.DataType != TableType.LottoWeli)
                        {
                            int intIntStar = (i01 - 1) * intInterval01;
                            if (intIntStar <= 0) intIntStar = 1;
                            int intIntEnd = (i01 * intInterval01) - 1;
                            if (intIntEnd > Dataset00.LottoNumbers) intIntEnd = Dataset00.LottoNumbers; int intCheckNum = int.Parse(drRow[strFieldName].ToString());
                            if (intCheckNum >= intIntStar && intCheckNum <= intIntEnd)
                            {
                                drNewRow[strIntFieldName] = int.Parse(drNewRow[strIntFieldName].ToString()) + 1;
                            }
                        }
                    }
                }
                dtDataInterval.Rows.Add(drNewRow);
            }

            #endregion 計算 Interval

            #endregion 建立 Table dtDataInterval
            return dtDataInterval;
        }
        private void ShowTableInterval(StuGLSearch stuSearch00, DataSet dsIntervals)
        {
            DataTable dtDataInterval = dsIntervals.Tables["Interval"];
            DataTable dtDataInterval10 = dsIntervals.Tables["Interval10"];
            CGLDataSet Dataset00 = new CGLDataSet(stuSearch00.LottoType);
            Dictionary<string, string> dicCurrentData = new CGLSearch().GetCurrentData(stuSearch00);
            #region Set lstCurreny
            List<int> lstCurrentData = new List<int>();
            for (int intCount = 1; intCount <= Dataset00.CountNumber; intCount++)
            {
                lstCurrentData.Add(int.Parse(dicCurrentData[string.Format("lngL{0}", intCount)].ToString()));
            }
            if (Dataset00.DataType != TableType.LottoWeli && Dataset00.SCountNumber > 0)
            {
                lstCurrentData.Add(int.Parse(dicCurrentData[string.Format("lngS1")].ToString()));
            }
            #endregion Set lstCurreny

            #region Show Table dtDataInterval

            #region Expander dtDataInterval

            #region DataGrid dtDataInterval
            DataGrid dgInterval = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Visible,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                AutoGenerateColumns = false
            };
            #region Set Columns of DataGrid dgInterval
            DataGridTextColumn dgtColumnInterval;
            if (dgInterval.Columns.Count == 0)
            {
                foreach (DataColumn dcdgColumn in dtDataInterval.Columns)
                {
                    dgtColumnInterval = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1),
                        Binding = new System.Windows.Data.Binding(dcdgColumn.ColumnName),
                        IsReadOnly = true
                    };
                    dgInterval.Columns.Add(dgtColumnInterval);
                }
            }
            #endregion
            dgInterval.ItemsSource = dtDataInterval.DefaultView;
            #endregion DataGrid

            Expander expInterval = new Expander()
            {
                Header = "100 期 數字區間表 ",
                ExpandDirection = ExpandDirection.Down,
                IsExpanded = true,
                Background = Brushes.Aqua,
                MaxHeight = 500,
                Content = dgInterval
            };
            #endregion Expander

            #region Expander dgData Interval 10

            #region DataGrid
            DataGrid dgInterval10 = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Visible,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                AutoGenerateColumns = false
            };
            #region Set Columns of DataGrid dgInterval
            DataGridTextColumn dgtColumnInterval10;
            if (dgInterval10.Columns.Count == 0)
            {
                foreach (DataColumn dcdgColumn in dtDataInterval10.Columns)
                {
                    dgtColumnInterval10 = new DataGridTextColumn()
                    {
                        Header = dcdgColumn.Caption,
                        Binding = new System.Windows.Data.Binding(dcdgColumn.ColumnName),
                        IsReadOnly = true
                    };
                    dgInterval10.Columns.Add(dgtColumnInterval10);
                }
            }
            #endregion
            dgInterval10.ItemsSource = dtDataInterval10.DefaultView;
            #endregion DataGrid

            Expander expInterval10 = new Expander()
            {
                Header = string.Format("100 期 數字區間表 ({0})", 10),
                ExpandDirection = ExpandDirection.Down,
                IsExpanded = true,
                Background = Brushes.Aqua,
                MaxHeight = 500,
                Content = dgInterval10
            };
            #endregion Expander

            #region StackPanel
            StackPanel spMain = new StackPanel()
            {
                CanVerticallyScroll = true,
                CanHorizontallyScroll = true
            };
            spMain.Children.Add(expInterval);
            spMain.Children.Add(expInterval10);
            #endregion StackPanel

            #region Scrollview 
            ScrollViewer svScroll = new ScrollViewer()
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = spMain
            };
            #endregion Scrollview 

            #region Window
            Window wInterval = new Window()
            {
                Name = "tblInterval",
                Title = string.Format("100 期 數字區間表 ({0})", stuSearch00.StrTitle),
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.SingleBorderWindow,
                Content = svScroll
            };
            wInterval.Show();
            #endregion Window
            #endregion Show Table dtDataInterval
        }
        #endregion btnTableInterval

        /// <summary>
        /// 冷熱門數字表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnTableActive_Click(object sender, RoutedEventArgs e) //冷熱門數字表
        {

            #region 設定 stuSearchActive
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSearchActive = stuRibbonSearchOption;
            stuSearchActive.IntDataOffset = 0;
            stuSearchActive.IntSearchLimit = 0;
            stuSearchActive.IntSearchOffset = 0;
            stuSearchActive = new CGLSearch().GetMethodSN(stuSearchActive);
            #endregion 設定 stuSearchActive

            CGLDataSet Dataset00 = new CGLDataSet(stuSearchActive.LottoType);
            Dictionary<string, string> dicCurrentData = new Dictionary<string, string>();

            dicCurrentData = new CGLSearch().GetCurrentData(stuSearchActive);

            #region 檢查 tblMissAll 是否有資料
            if (!new CGLMissAll().HasMissAllData(stuSearchActive))
            {
                new CGLMissAll().GetMissAlldic(stuSearchActive);
            }
            #endregion 檢查是否有資料

            #region 檢查 tblActive是否有資料
            if (!new CGLSearch().HasActiveData(stuSearchActive))
            {
                new CGLSearch().SearchActive(stuSearchActive);
            }
            #endregion 檢查是否有資料

            #region 導入 100 期 數值

            #region Query String
            StuGLData stuData00 = new StuGLData()
            {
                LottoType = stuSearchActive.LottoType,
                DataBaseType = DatabaseType.Active,
                StrSELECT = "TOP 100 ",
                StrFROM = "[qryActive] ",
                StrWHERE = string.Format("[{0}] = {1} AND [{2}] < {3} ",
                                        "lngMethodSN", stuSearchActive.LngMethodSN,
                                        "lngTotalSN", stuSearchActive.LngCurrentData),
                StrORDER = "[lngTotalSN] DESC ;"
            };
            #region set select
            stuData00.StrSELECT += "[lngTotalSN] , [lngMethodSN] , [lngDateSN] , [strWeekend] , ";
            for (int i = 1; i <= Dataset00.CountNumber; i++)
            {
                stuData00.StrSELECT += string.Format("[lngL{0}] ,", i.ToString());
            }
            if (Dataset00.DataType != TableType.LottoWeli && Dataset00.SCountNumber > 0)
            {
                stuData00.StrSELECT += "[lngS1] ,";
            }
            for (int i = 1; i <= Dataset00.CountNumber; i++)
            {
                stuData00.StrSELECT += string.Format("[intAL{0}] ,", i.ToString());
            }
            if (Dataset00.DataType != TableType.LottoWeli && Dataset00.SCountNumber > 0)
            {
                stuData00.StrSELECT += "[intAS1] ,";
            }
            stuData00.StrSELECT += "[intLess10] , [intSum] , [sglAvg] ";
            #endregion
            #region  Set Compare Option 
            if (stuSearchActive.BoolFieldMode && stuSearchActive.StrCompares != "gen")
            {
                string[] strCompare = stuSearchActive.StrCompares.Split('#');
                foreach (string strCompareOption in strCompare)
                {
                    stuData00.StrWHERE += string.Format("{0} [{1}] = '{2}' "
                                                        , stuSearchActive.StrCompareType
                                                        , strCompareOption
                                                        , dicCurrentData[strCompareOption]);
                }
            }
            #endregion
            #endregion Query String
            DataTable dtDataTable = stuData00.GetSourceData(stuData00);
            #endregion 導入 100 期 數值

            #region Show Window

            #region Expander dgData

            #region DataGrid
            DataGrid dgActive = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Visible,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                AutoGenerateColumns = false
            };
            #region Set Columns of DataGrid dgActive
            DataGridTextColumn dgtColumnActive;
            if (dgActive.Columns.Count == 0)
            {
                foreach (DataColumn dcdgColumn in dtDataTable.Columns)
                {
                    dgtColumnActive = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1),
                        Binding = new System.Windows.Data.Binding(dcdgColumn.ColumnName),
                        IsReadOnly = true
                    };
                    dgActive.Columns.Add(dgtColumnActive);
                }
            }
            #endregion
            dgActive.ItemsSource = dtDataTable.DefaultView;
            #endregion DataGrid

            Expander expActive = new Expander()
            {
                Header = "100 期 熱門冷門數字資料",
                ExpandDirection = ExpandDirection.Down,
                IsExpanded = true,
                Background = Brushes.Aqua,
                MaxHeight = 500,
                Content = dgActive
            };
            #endregion Expander

            #region StackPanel
            //StackPanel spMain = new StackPanel();
            //spMain.CanVerticallyScroll = true;
            //spMain.CanHorizontallyScroll = true;
            //spMain.Children.Add(expActive);
            #endregion StackPanel

            #region ScrollViewer
            //ScrollViewer svViewer = new ScrollViewer();
            //svViewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
            //svViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            //svViewer.Content = spMain;
            #endregion ScrollViewer

            #region Window
            Window wActive = new Window()
            {
                Name = "tblActive",
                Title = string.Format("熱門冷門數字資料 ({0})", stuSearchActive.StrTitle),
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.SingleBorderWindow,
                Content = dgActive
            };
            wActive.Show();
            #endregion Window

            #endregion Show Window
        }

        /// <summary>
        /// 末位數字表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnTableDigits_Click(object sender, RoutedEventArgs e) //末位數字表
        {
            #region 設定 stuSearchDigits
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSearchDigits = stuRibbonSearchOption;
            #endregion 設定 stuSearchDigits

            #region 設定
            CGLDataSet Dataset00 = new CGLDataSet(stuSearchDigits.LottoType);
            Dictionary<string, string> dicCurrentData = new Dictionary<string, string>();
            dicCurrentData = new CGLSearch().GetCurrentData(stuSearchDigits);
            Dictionary<string, int>[] dicALCount = new Dictionary<string, int>[6];
            DataTable[] dtTableDigits = new DataTable[6];
            DataTable[] dtDataShow = new DataTable[6];
            #endregion 設定

            #region 計算
            for (int intTrace = 0; intTrace <= 5; intTrace++)
            {
                dicALCount[intTrace] = new Dictionary<string, int>();
                #region 導入數值

                #region Query String
                StuGLData stuData00 = new StuGLData()
                {
                    LottoType = stuSearchDigits.LottoType,
                    DataBaseType = DatabaseType.Data,
                    StrSELECT = string.Format("TOP {0} ", (intTrace + 5).ToString()),
                    StrFROM = string.Format("[{0}] ", "qryData"),
                    StrWHERE = string.Format("[{0}] < {1} ", "lngTotalSN", stuSearchDigits.LngCurrentData),
                    StrORDER = "[lngTotalSN] DESC ;"
                };

                #region set select
                stuData00.StrSELECT += "[lngTotalSN] , [lngDateSN] , ";
                for (int i = 1; i <= Dataset00.CountNumber; i++)
                {
                    stuData00.StrSELECT += string.Format("[lngL{0}] ,", i.ToString());
                }
                if (Dataset00.DataType != TableType.LottoWeli && Dataset00.SCountNumber > 0)
                {
                    stuData00.StrSELECT += "[lngS1] ,";
                }
                stuData00.StrSELECT = stuData00.StrSELECT.TrimEnd(',');
                #endregion

                #region  Set Compare Option 
                if (stuSearchDigits.BoolFieldMode && stuSearchDigits.StrCompares != "gen")
                {
                    string[] strCompare = stuSearchDigits.StrCompares.Split('#');
                    foreach (string strCompareOption in strCompare)
                    {
                        stuData00.StrWHERE += string.Format(" {0} [{1}] = '{2}' "
                                                            , stuSearchDigits.StrCompareType
                                                            , strCompareOption
                                                            , dicCurrentData[strCompareOption]);
                    }
                }
                #endregion
                #endregion Query String
                dtTableDigits[intTrace] = stuData00.GetSourceData(stuData00);
                #endregion 導入數值
                for (int i = 0; i <= 9; i++) { dicALCount[intTrace].Add(i.ToString(), 0); }

                #region 開始計算
                foreach (DataRow drRow in dtTableDigits[intTrace].Rows)
                {
                    for (int i = 1; i <= Dataset00.CountNumber; i++)
                    {
                        int intDigitsCount = int.Parse(drRow[string.Format("lngL{0}", i.ToString())].ToString());
                        dicALCount[intTrace][(intDigitsCount % 10).ToString()]++;

                    }
                    if (Dataset00.DataType != TableType.LottoWeli && Dataset00.SCountNumber > 0)
                    {
                        int intDigitsCount = int.Parse(drRow["lngS1"].ToString());
                        dicALCount[intTrace][(intDigitsCount % 10).ToString()]++;
                    }
                }
                #endregion 開始計算
                #region 按頻率小到大排序

                //var sortDictionary = from keys in dicALCount[intTrace] orderby keys.Value ascending select keys;
                //Dictionary<string, int> dicSortResult = new Dictionary<string, int>();
                //foreach (var KeyValuePair in sortDictionary)
                //{
                //    dicSortResult.Add(KeyValuePair.Key, KeyValuePair.Value);
                //}
                Dictionary<string, int> dicSortResult = dicALCount[intTrace].OrderBy(x => x.Value).ToDictionary(x => x.Key, x => x.Value);
                #endregion 按頻率小到大排序
                #region 建立 dtDataShow
                dtDataShow[intTrace] = new DataTable();
                #region 建立 Columns
                dtDataShow[intTrace].Columns.Add("intDigits", typeof(int));
                dtDataShow[intTrace].Columns.Add("intDigitsCount", typeof(int));
                dtDataShow[intTrace].Columns.Add("strNums", typeof(string));
                #endregion 建立 Columns
                #endregion 建立 dtDataShow
                #region 填入資料
                foreach (var KeyValuePair in dicSortResult)
                {
                    DataRow drRowShow = dtDataShow[intTrace].NewRow();
                    drRowShow["intDigits"] = int.Parse(KeyValuePair.Key);
                    drRowShow["intDigitsCount"] = KeyValuePair.Value;
                    string strNums = "";
                    foreach (DataRow drRow in dtTableDigits[intTrace].Rows)
                    {
                        for (int i = 1; i <= Dataset00.CountNumber; i++)
                        {
                            int intDigitsCount = int.Parse(drRow[string.Format("lngL{0}", i.ToString())].ToString());
                            if ((intDigitsCount % 10) == int.Parse(KeyValuePair.Key))
                            {
                                strNums += string.Format("{0}#", intDigitsCount.ToString("D2"));
                            }
                        }
                        if (Dataset00.DataType != TableType.LottoWeli && Dataset00.SCountNumber > 0)
                        {
                            int intDigitsCount = int.Parse(drRow["lngS1"].ToString());
                            if ((intDigitsCount % 10) == int.Parse(KeyValuePair.Key))
                            {
                                strNums += string.Format("{0}#", intDigitsCount.ToString("D2"));
                            }
                        }
                    }
                    strNums = strNums.TrimEnd('#');
                    string[] strTemp = strNums.Split('#');
                    string[] strTempDist = strTemp.Distinct().ToArray();
                    drRowShow["strNums"] = string.Join("#", strTempDist);

                    dtDataShow[intTrace].Rows.Add(drRowShow);
                }
                #endregion 填入資料
            }
            #endregion 計算

            #region 顯示
            #region StackPanel
            StackPanel spMain = new StackPanel()
            {
                CanVerticallyScroll = true,
                CanHorizontallyScroll = true
            };
            #endregion StackPanel

            for (int intTrace = 0; intTrace <= 5; intTrace++)
            {

                #region DataGrid
                DataGrid dgDigits = new DataGrid()
                {
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Visible,
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    VerticalAlignment = VerticalAlignment.Top,
                    CanUserAddRows = false,
                    CanUserDeleteRows = false,
                    AutoGenerateColumns = false
                };
                #region Set Columns of DataGrid dgDigits
                DataGridTextColumn dgtColumnDigits;
                if (dgDigits.Columns.Count == 0)
                {
                    foreach (DataColumn dcdgColumn in dtDataShow[intTrace].Columns)
                    {
                        dgtColumnDigits = new DataGridTextColumn()
                        {
                            Header = dcdgColumn.Caption,
                            Binding = new System.Windows.Data.Binding(dcdgColumn.ColumnName),
                            IsReadOnly = true
                        };
                        dgDigits.Columns.Add(dgtColumnDigits);
                    }
                }
                #endregion
                dgDigits.ItemsSource = dtDataShow[intTrace].DefaultView;
                #endregion DataGrid

                #region Expander dgData
                Expander expDigits = new Expander()
                {
                    Header = string.Format(" {0}期 末位數字表", (intTrace + 5).ToString("D2")),
                    ExpandDirection = ExpandDirection.Down,
                    IsExpanded = true,
                    Background = Brushes.Aqua,
                    MaxHeight = 500,
                    Content = dgDigits
                };
                #endregion Expander

                spMain.Children.Add(expDigits);
            }

            #region ScrollViewer
            ScrollViewer svViewer = new ScrollViewer()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = spMain
            };
            #endregion ScrollViewer
            #region Window
            Window wDigits = new Window()
            {
                Name = "tblDigits",
                Title = string.Format("末位數字表 ({0})", stuSearchDigits.StrTitle),
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.SingleBorderWindow,
                Content = svViewer
            };
            wDigits.Show();
            #endregion Window
            #endregion 顯示
        }

        #region btnTableDataN 數字振盪表
        private void BtnTableDataN_Click(object sender, RoutedEventArgs e)
        {
            #region 設定 stuSearchDataN
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSearchDataN = stuRibbonSearchOption;
            #endregion 設定 stuSearchsmart
            #region Setting bwBackgroundWorker00
            bwBackgroundWorker00 = new BackgroundWorker()
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };
            bwBackgroundWorker00.DoWork += BwsDataN_DoWork;
            bwBackgroundWorker00.ProgressChanged += BwDataN_ProgressChanged;
            bwBackgroundWorker00.RunWorkerCompleted += BwDataN_RunWorkerCompleted;
            #endregion Setting bwTablePercent
            bwBackgroundWorker00.RunWorkerAsync(stuSearchDataN);
            btnTableDataN.Content = "計算中...";
            btnTableDataN.IsEnabled = false;
        }
        private void BwsDataN_DoWork(object sender, DoWorkEventArgs e)
        {
            #region 設定
            StuGLSearch stuSearchDataN = (StuGLSearch)e.Argument;
            BackgroundWorker bwMyWork = (BackgroundWorker)sender;
            bwMyWork.ReportProgress(1);
            CGLDataSet Dataset00 = new CGLDataSet(stuSearchDataN.LottoType);
            Dictionary<string, object> dicArgument = new Dictionary<string, object>();
            DataSet dsDataSet = new DataSet();
            #endregion
            dicArgument.Add("stuSearch", stuSearchDataN);
            #region 檢查 DataN 是否有資料 
            if (!new CGLDataN().HasDataNData(stuSearchDataN))
            {
                new CGLDataN().GetDataNdic(stuSearchDataN);
            }
            #endregion 檢查是否有資料

            List<int> lstCurrentNums = new CGLSearch().GetlstCurrentNums(stuSearchDataN);
            int[] intSections = { 5, 10, 25, 50, 100 };

            #region Loading Top200 Datas
            #region set Query string
            StuGLData stuData00 = new StuGLData()
            {
                LottoType = stuSearchDataN.LottoType,
                DataBaseType = DatabaseType.DataN,
                StrSELECT = "TOP 200 ",
                StrFROM = "[qryDataN] ",
                StrWHERE = string.Format("[{0}] = {1} AND [{2}] < {3} ",
                                        "lngMethodSN", stuSearchDataN.LngMethodSN,
                                        "lngTotalSN", stuSearchDataN.LngCurrentData),
                StrORDER = "[lngTotalSN] DESC "
            };
            stuData00.StrSELECT += " [lngTotalSN] , [lngMethodSN] , [lngDateSN] , [intSum] ,";
            for (int intL = 1; intL <= lstCurrentNums.Count; intL++)
            {
                stuData00.StrSELECT += string.Format(" [lngN{0}] , [sglAvg{0}] , [sglPA{0}] , [intMin{0}] , [intMax{0}] ,", intL);
                foreach (int Section in intSections)
                {
                    stuData00.StrSELECT += string.Format(" [sglAvg{0}{1}] , [sglPA{0}{1}] , [intMin{0}{1}] , [intMax{0}{1}] ,"
                                                                , intL, Section.ToString("D2"));
                }
            }
            stuData00.StrSELECT = stuData00.StrSELECT.TrimEnd(',');

            #region  Set Compare Option 
            //if (stuSearch00.boolFieldMode && stuSearch00.strCompares != "gen")
            //{
            //    string[] strCompare = stuSearch00.strCompares.Split('#');
            //    foreach (string strCompareOption in strCompare)
            //    {
            //        stuSearch00.strWHERE += string.Format(" {0} [{1}] = '{2}' "
            //                                            , stuSearch00.strCompareType
            //                                            , strCompareOption
            //                                            , dicCurrentData[strCompareOption]);
            //    }
            //}
            #endregion

            #endregion set Query string
            DataTable dtNextN = stuData00.GetSourceData(stuData00);

            dtNextN.TableName = "Top200";
            dsDataSet.Tables.Add(dtNextN);
            #endregion Loading Top200 Datas

            #region Caculate Next Data
            StuGLSearch stuSearchTemp = stuSearchDataN;
            DataSet dsTemp = new DataSet();

            #region query
            stuData00 = new StuGLData()
            {
                LottoType = stuSearchTemp.LottoType,
                DataBaseType = DatabaseType.DataN,
                StrSELECT = "*",
                StrFROM = string.Format("[{0}] ", "qryDataN"),
                StrWHERE = string.Format("[{0}] = {1} AND [{2}] < {3} ",
                                                   "lngMethodSN", stuSearchTemp.LngMethodSN,
                                                   "lngTotalSN", stuSearchTemp.LngCurrentData),
                StrORDER = string.Format("[lngTotalSN] DESC")
            };
            #endregion  query
            DataTable dtTemp = stuData00.GetSourceData(stuData00);
            dtTemp.TableName = "tblAll";
            dsTemp.Tables.Add(dtTemp);

            #region set dtTop
            foreach (int Section in intSections)
            {
                #region query
                stuData00 = new StuGLData()
                {
                    LottoType = stuSearchTemp.LottoType,
                    DataBaseType = DatabaseType.DataN,
                    StrSELECT = string.Format("Top {0} * ", Section - 1),
                    StrFROM = string.Format("[{0}] ", "qryDataN"),
                    StrWHERE = string.Format("[{0}] = {1} AND [{2}] < {3} ",
                                                       "lngMethodSN", stuSearchTemp.LngMethodSN,
                                                       "lngTotalSN", stuSearchTemp.LngCurrentData),
                    StrORDER = string.Format("[lngTotalSN] DESC")
                };
                #endregion query

                dtTemp = stuData00.GetSourceData(stuData00);

                dtTemp.TableName = string.Format("tblTop{0}", Section.ToString("D2"));
                dsTemp.Tables.Add(dtTemp);
            }
            #endregion

            #region Setup each Count Num Table
            for (int intL = 1; intL <= lstCurrentNums.Count; intL++)
            {
                DataTable tblNext = new DataTable()
                {
                    TableName = string.Format("lngN{0}", intL)
                };
                DataColumn dcColumn = new DataColumn(string.Format("lngN{0}", intL), typeof(float));
                tblNext.Columns.Add(dcColumn);
                dcColumn = new DataColumn(string.Format("sglAvg{0}", intL), typeof(float));
                tblNext.Columns.Add(dcColumn);
                dcColumn = new DataColumn(string.Format("sglPA{0}", intL), typeof(float));
                tblNext.Columns.Add(dcColumn);
                foreach (int Section in intSections)
                {
                    dcColumn = new DataColumn(string.Format("sglAvg{0}{1}", intL, Section.ToString("D2")), typeof(float));
                    tblNext.Columns.Add(dcColumn);
                    dcColumn = new DataColumn(string.Format("sglPA{0}{1}", intL, Section.ToString("D2")), typeof(float));
                    tblNext.Columns.Add(dcColumn);
                }
                dsDataSet.Tables.Add(tblNext);
            }
            #endregion Setup each Count Num Table

            #region Caculate each Count Num

            for (int intL = 1; intL <= lstCurrentNums.Count; intL++)
            {
                int intMin = int.Parse(dsTemp.Tables["tblAll"].Rows[0][string.Format("intMin{0}", intL)].ToString());
                int intMax = int.Parse(dsTemp.Tables["tblAll"].Rows[0][string.Format("intMax{0}", intL)].ToString());
                for (int intNum = intMin; intNum <= intMax; intNum++)
                {
                    string strNum = string.Format("lngN{0}", intL);
                    DataRow drRow = dsDataSet.Tables[strNum].NewRow();
                    drRow[strNum] = intNum;
                    double intRecordCount = (double)dsTemp.Tables["tblAll"].Rows.Count;
                    double dblTotal = double.Parse(dsTemp.Tables["tblAll"].Compute(string.Format("Sum([{0}])", strNum), string.Empty).ToString());
                    double dblAvg = Math.Round((dblTotal + (double)intNum) / (intRecordCount + 1d), 2);
                    double dblPA = Math.Round(double.Parse(((double)intNum - dblAvg).ToString()), 2);
                    drRow[string.Format("sglAvg{0}", intL)] = dblAvg;
                    drRow[string.Format("sglPA{0}", intL)] = dblPA;
                    foreach (int Section in intSections)
                    {
                        intRecordCount = (double)dsTemp.Tables[string.Format("tblTop{0}", Section.ToString("D2"))].Rows.Count;
                        dblTotal = double.Parse(dsTemp.Tables[string.Format("tblTop{0}", Section.ToString("D2"))].Compute(string.Format("Sum([{0}])", strNum), string.Empty).ToString());
                        dblAvg = Math.Round((dblTotal + (double)intNum) / (intRecordCount + 1d), 2);
                        dblPA = Math.Round(double.Parse(((double)intNum - dblAvg).ToString()), 2);
                        drRow[string.Format("sglAvg{0}{1}", intL, Section.ToString("D2"))] = dblAvg;
                        drRow[string.Format("sglPA{0}{1}", intL, Section.ToString("D2"))] = dblPA;
                    }
                    dsDataSet.Tables[string.Format("lngN{0}", intL)].Rows.Add(drRow);
                }
            }
            #endregion Caculate each Count Num

            #endregion Caculate Next Data

            dicArgument.Add("DataSet", dsDataSet);
            e.Result = dicArgument;
        }
        private void BwDataN_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 1)
            {
                this.tbStatusTextBlock.Text = "計算中...";
            }
        }

        private void BwDataN_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            BackgroundWorker bwMywork = (BackgroundWorker)sender;
            Dictionary<string, object> dicArgument = (Dictionary<string, object>)e.Result;
            DataSet dsDataSet = (DataSet)dicArgument["DataSet"];
            ShowDataN((StuGLSearch)dicArgument["stuSearch"], dsDataSet);
            tbStatusTextBlock.Text = "Ready";
            btnTableDataN.Content = "數字振盪表";
            btnTableDataN.IsEnabled = true;
            bwMywork.Dispose();
        }

        private void ShowDataN(StuGLSearch stuSearch00, DataSet dsDataSet)
        {
            Dictionary<string, string> dicCurrentData = new Dictionary<string, string>();
            dicCurrentData = new CGLSearch().GetCurrentData(stuSearch00);
            DataGridTextColumn dgtColumn;
            #region Show Window

            #region Setup Window
            Window wDataN = new Window()
            {
                Name = "tblDataN",
                Title = string.Format("數字振盪表 ({0}) ", stuSearch00.StrTitle),
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.SingleBorderWindow
            };
            #endregion Window             

            #region Setup StackPanel MAin
            StackPanel spMain = new StackPanel()
            {
                Orientation = Orientation.Vertical,
                VerticalAlignment = VerticalAlignment.Top,
                HorizontalAlignment = HorizontalAlignment.Left
            };
            #endregion StackPanel

            #region Set DataGrid dgCurrentData
            #region Convert dictionary dicCurrentData to table dtCurrentData
            DataTable dtCurrentData = new CGLFunc().CDicTOTable(dicCurrentData);
            #endregion
            DataGrid dgCurrentData = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                RowDetailsVisibilityMode = DataGridRowDetailsVisibilityMode.Visible,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch,
                CanUserAddRows = true,
                CanUserDeleteRows = false,
                CanUserResizeColumns = true,
                CanUserResizeRows = true,
                CanUserSortColumns = true,
                IsReadOnly = true,
                AutoGenerateColumns = false,
                ItemsSource = dtCurrentData.DefaultView
            };


            #region Set Columns of DataGrid dgCurrentData 

            if (dgCurrentData.Columns.Count == 0)
            {
                foreach (var KeyPair in dicCurrentData)
                {
                    dgtColumn = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(KeyPair.Key, 1),
                        Binding = new Binding(KeyPair.Key),
                        IsReadOnly = true
                    };
                    dgCurrentData.Columns.Add(dgtColumn);
                }
            }
            #endregion

            #endregion
            spMain.Children.Add(dgCurrentData);

            foreach (DataTable dtDataTable in dsDataSet.Tables)
            {

                #region Setup Expander 
                Expander expSection = new Expander()
                {
                    Header = dtDataTable.TableName,
                    IsExpanded = false
                };
                if (dtDataTable.TableName == "Top200") expSection.IsExpanded = true;
                expSection.HorizontalAlignment = HorizontalAlignment.Left;
                expSection.HorizontalContentAlignment = HorizontalAlignment.Stretch;
                expSection.VerticalAlignment = VerticalAlignment.Top;
                expSection.VerticalContentAlignment = VerticalAlignment.Top;
                #endregion Expan

                #region DataGrid DataN
                DataGrid dgDataN = new DataGrid()
                {
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    //dgDataN.RowDetailsVisibilityMode = DataGridRowDetailsVisibilityMode.Visible;
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    VerticalAlignment = VerticalAlignment.Top,
                    CanUserAddRows = false,
                    CanUserDeleteRows = false,
                    CanUserSortColumns = false,
                    CanUserResizeColumns = false,
                    CanUserResizeRows = false,
                    AutoGenerateColumns = false
                };
                #region Set Columns of DataGrid dgDataN
                List<int> lstCurrentNums = new CGLSearch().GetlstCurrentNums(stuSearch00);
                if (dgDataN.Columns.Count == 0)
                {
                    List<string> cName = new List<string>();
                    if (dtDataTable.TableName != "Top200")
                    {
                        int[] intSections = { 5, 10, 25, 50, 100 };
                        string strNumIndex = dtDataTable.TableName.Substring(4);
                        cName.Add(string.Format("sglAvg{0}", strNumIndex));
                        foreach (int Section in intSections)
                        {
                            cName.Add(string.Format("sglAvg{0}{1}", strNumIndex, Section.ToString("D2")));
                        }
                    }

                    foreach (DataColumn dcdgColumn in dtDataTable.Columns)
                    {
                        dgtColumn = new DataGridTextColumn
                        {
                            Header = new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1),
                            Binding = new Binding(dcdgColumn.ColumnName),
                            IsReadOnly = true
                        };
                        #region set hit numbers
                        if (dcdgColumn.ColumnName.Substring(0, 4) == "lngN")
                        {
                            Style Style00 = new Style()
                            {
                                TargetType = typeof(DataGridCell)
                            };
                            Setter setter00 = new Setter()
                            {
                                Property = BackgroundProperty,
                                Value = Brushes.LightSalmon
                            };
                            Style00.Setters.Add(setter00);
                            setter00 = new Setter()
                            {
                                Property = ForegroundProperty,
                                Value = Brushes.Black
                            };
                            Style00.Setters.Add(setter00);
                            if (dtDataTable.TableName != "Top200")
                            {
                                #region set Trigger
                                DataTrigger dtTrig00 = new DataTrigger()
                                {
                                    Binding = dgtColumn.Binding,
                                    Value = dsDataSet.Tables["Top200"].Rows[0][dcdgColumn.ColumnName].ToString()
                                };
                                setter00 = new Setter()
                                {
                                    Property = ForegroundProperty,
                                    Value = Brushes.Black
                                };
                                dtTrig00.Setters.Add(setter00);
                                setter00 = new Setter()
                                {
                                    Property = BackgroundProperty,
                                    Value = Brushes.OrangeRed
                                };
                                dtTrig00.Setters.Add(setter00);
                                Style00.Triggers.Add(dtTrig00);

                                DataTrigger dtTrig01 = new DataTrigger()
                                {
                                    Binding = dgtColumn.Binding,
                                    Value = (lstCurrentNums[int.Parse(dcdgColumn.ColumnName.Substring(4)) - 1]).ToString()
                                };
                                setter00 = new Setter()
                                {
                                    Property = ForegroundProperty,
                                    Value = Brushes.Black
                                };
                                dtTrig01.Setters.Add(setter00);
                                setter00 = new Setter()
                                {
                                    Property = BackgroundProperty,
                                    Value = Brushes.DeepSkyBlue
                                };
                                dtTrig01.Setters.Add(setter00);
                                Style00.Triggers.Add(dtTrig01);
                                #endregion set Trigger
                            }
                            dgtColumn.CellStyle = Style00;
                        }
                        #endregion set hit numbers

                        #region Setup Top n tablt
                        if (dtDataTable.TableName != "Top200" && cName.Contains(dcdgColumn.ColumnName))
                        {
                            Style Style00 = new Style()
                            {
                                TargetType = typeof(DataGridCell)
                            };
                            #region set Trigger
                            DataTrigger dtTrig00 = new DataTrigger()
                            {
                                Binding = dgtColumn.Binding,
                                Value = dsDataSet.Tables["Top200"].Rows[0][dcdgColumn.ColumnName].ToString()
                            };
                            Setter setter00 = new Setter()
                            {
                                Property = ForegroundProperty,
                                Value = Brushes.Black
                            };
                            dtTrig00.Setters.Add(setter00);
                            setter00 = new Setter()
                            {
                                Property = BackgroundProperty,
                                Value = Brushes.GreenYellow
                            };
                            dtTrig00.Setters.Add(setter00);
                            Style00.Triggers.Add(dtTrig00);
                            #endregion set Trigger
                            dgtColumn.CellStyle = Style00;
                        }
                        #endregion Setup Top n tablt

                        dgDataN.Columns.Add(dgtColumn);
                    }
                }
                #endregion
                dgDataN.ItemsSource = dtDataTable.DefaultView;
                dgDataN.Width = 1400;
                if (dtDataTable.Rows.Count > 20)
                {
                    dgDataN.Height = 300;
                }
                #endregion DataGrid
                expSection.Content = dgDataN;
                spMain.Children.Add(expSection);
            }
            #region ScrollViewer
            ScrollViewer svViewer = new ScrollViewer()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = spMain
            };
            #endregion ScrollViewer

            #region Show window
            wDataN.Content = svViewer;
            wDataN.Show();
            wDataN.WindowState = WindowState.Maximized;
            #endregion Show window

            #endregion
        }

        #endregion btnTableDataN

        /// <summary>
        /// 遺漏偏差表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnTableMiss_Click(object sender, RoutedEventArgs e) //遺漏偏差表 
        {
            #region 設定 stuSearchMiss
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSearchMiss = stuRibbonSearchOption;
            #endregion 設定 stuSearchMiss

            CGLDataSet Dataset00 = new CGLDataSet(stuSearchMiss.LottoType);
            Dictionary<string, string> dicCurrentData = new Dictionary<string, string>();
            dicCurrentData = new CGLSearch().GetCurrentData(stuSearchMiss);

            #region 檢查 tblActive是否有資料
            if (!new CGLSearch().HasActiveData(stuSearchMiss))
            {
                new CGLSearch().SearchActive(stuSearchMiss);
            }
            #endregion 檢查是否有資料

            #region StackPanel
            StackPanel spMain = new StackPanel()
            {
                CanVerticallyScroll = true,
                CanHorizontallyScroll = true
            };
            #endregion StackPanel

            dicCurrentData = new CGLSearch().GetCurrentData(stuSearchMiss);
            for (int intTrace = 5; intTrace <= 10; intTrace++)
            {
                #region 導入數值

                #region Query String
                StuGLData stuData00 = new StuGLData()
                {
                    LottoType = stuSearchMiss.LottoType,
                    DataBaseType = DatabaseType.Active,
                    StrSELECT = string.Format("TOP {0} ", intTrace.ToString()),
                    StrFROM = string.Format("[{0}] ", "qryActive"),
                    StrWHERE = string.Format("[{0}] = {1} AND [{2}] < {3} ",
                                            "lngMethodSN", stuSearchMiss.LngMethodSN,
                                            "lngTotalSN", stuSearchMiss.LngCurrentData),
                    StrORDER = "[lngTotalSN] DESC "
                };
                stuData00.StrSELECT += "[lngTotalSN] , [lngMethodSN] , [lngDateSN] , [strWeekend] , ";
                for (int i = 1; i <= Dataset00.CountNumber; i++)
                {
                    stuData00.StrSELECT += string.Format("[intAL{0}] ,", i.ToString());
                }
                if (stuSearchMiss.LottoType != TableType.LottoWeli && new CGLDataSet(stuSearchMiss.LottoType).SCountNumber > 0)
                {
                    stuData00.StrSELECT += "[intAS1] ,";
                }
                stuData00.StrSELECT = stuData00.StrSELECT.TrimEnd(',');

                #region  Set Compare Option 
                if (stuSearchMiss.BoolFieldMode && stuSearchMiss.StrCompares != "gen")
                {
                    string[] strCompare = stuSearchMiss.StrCompares.Split('#');
                    foreach (string strCompareOption in strCompare)
                    {
                        stuData00.StrWHERE += string.Format(" {0} [{1}] = '{2}' "
                                                            , stuSearchMiss.StrCompareType
                                                            , strCompareOption
                                                            , dicCurrentData[strCompareOption]);
                    }
                }
                #endregion

                #endregion Query String
                DataTable dtActive = stuData00.GetSourceData(stuData00);

                #endregion 導入數值
                Dictionary<string, int> dicALCount = new Dictionary<string, int>();
                for (int i = 0; i <= 5; i++) { dicALCount.Add(i.ToString(), 0); }
                #region 開始計算
                foreach (DataRow drRow in dtActive.Rows)
                {
                    for (int i = 1; i <= Dataset00.CountNumber; i++)
                    {
                        int intMissCount = int.Parse(drRow[string.Format("intAL{0}", i.ToString())].ToString());
                        if (intMissCount <= 5 && intTrace < 10)
                        {
                            dicALCount[intMissCount.ToString()]++;
                        }
                        if (intTrace == 10)
                        {
                            if (dicALCount.ContainsKey(intMissCount.ToString()))
                            {
                                dicALCount[intMissCount.ToString()]++;
                            }
                            else
                            {
                                dicALCount.Add(intMissCount.ToString(), 1);
                            }
                        }
                    }
                    if (Dataset00.DataType != TableType.LottoWeli && Dataset00.SCountNumber > 0)
                    {
                        int intMissCount = int.Parse(drRow["intAS1"].ToString());
                        if (intMissCount <= 5 && intTrace < 10)
                        {
                            dicALCount[intMissCount.ToString()]++;
                        }
                        if (intTrace == 10)
                        {
                            if (dicALCount.ContainsKey(intMissCount.ToString()))
                            {
                                dicALCount[intMissCount.ToString()]++;
                            }
                            else
                            {
                                dicALCount.Add(intMissCount.ToString(), 1);
                            }
                        }
                    }
                }
                #endregion 開始計算

                #region 按頻率小到大排序

                //var sortDictionary = from keys in dicALCount orderby keys.Value ascending select keys;
                //Dictionary<string, int> dicSortResult = new Dictionary<string, int>();
                //foreach (var KeyValuePair in sortDictionary)
                //{
                //    dicSortResult.Add(KeyValuePair.Key, KeyValuePair.Value);
                //}
                Dictionary<string, int> dicSortResult = dicALCount.OrderBy(x => x.Value).ToDictionary(x => x.Key, x => x.Value);
                #endregion 按頻率小到大排序

                #region 導入 qryMissAll 數值

                #region Query String
                stuData00 = new StuGLData()
                {
                    LottoType = stuSearchMiss.LottoType,
                    DataBaseType = DatabaseType.Miss,
                    StrSELECT = "TOP 1 ",
                    StrFROM = string.Format("[{0}] ", "qryMissAll"),
                    StrWHERE = string.Format("[{0}] = {1} AND [{2}] < {3} ",
                                            "lngMethodSN", stuSearchMiss.LngMethodSN,
                                            "lngTotalSN", stuSearchMiss.LngCurrentData),
                    StrORDER = "[lngTotalSN] DESC"
                };

                #region set select
                stuData00.StrSELECT += "[lngTotalSN] , [lngMethodSN] , [lngDateSN] , ";
                for (int i = 1; i <= Dataset00.CountNumber; i++)
                {
                    stuData00.StrSELECT += string.Format("[lngL{0}] ,", i.ToString());
                }
                if (Dataset00.DataType != TableType.LottoWeli && Dataset00.SCountNumber > 0)
                {
                    stuData00.StrSELECT += "[lngS1] ,";
                }
                for (int i = 1; i <= Dataset00.LottoNumbers; i++)
                {
                    stuData00.StrSELECT += string.Format("[lngM{0}] ,", i.ToString());
                }
                if (Dataset00.DataType == TableType.LottoWeli && Dataset00.SCountNumber > 0)
                {
                    for (int i = 1; i <= Dataset00.SCountNumber; i++)
                    {
                        stuData00.StrSELECT += string.Format("[lngMS{0}] ,", i.ToString());
                    }
                }
                stuData00.StrSELECT = stuData00.StrSELECT.TrimEnd(',');
                #endregion

                #region  Set Compare Option 
                if (stuSearchMiss.BoolFieldMode && stuSearchMiss.StrCompares != "gen")
                {
                    string[] strCompare = stuSearchMiss.StrCompares.Split('#');
                    foreach (string strCompareOption in strCompare)
                    {
                        stuData00.StrWHERE += string.Format(" {0} [{1}] = '{2}' "
                                                            , stuSearchMiss.StrCompareType
                                                            , strCompareOption
                                                            , dicCurrentData[strCompareOption]);
                    }
                }
                #endregion

                #endregion Query String
                DataTable dtMissAll = stuData00.GetSourceData(stuData00);

                #endregion 導入 qryMissAll 數值

                #region 建立 dtDataShow
                DataTable dtDataShow = new DataTable();
                #region 建立 Columns
                dtDataShow.Columns.Add("lngTotalSN", typeof(Int64));
                dtDataShow.Columns.Add("lngMethodSN", typeof(Int64));
                dtDataShow.Columns.Add("lngDateSN", typeof(Int64));
                dtDataShow.Columns.Add("intMiss", typeof(int));
                dtDataShow.Columns.Add("intMissCount", typeof(int));
                dtDataShow.Columns.Add("strNums", typeof(string));
                #endregion 建立 Columns
                #endregion 建立 dtDataShow

                #region 填入資料
                foreach (var KeyValuePair in dicSortResult)
                {
                    DataRow drRow = dtDataShow.NewRow();
                    drRow["lngTotalSN"] = dtMissAll.Rows[0]["lngTotalSN"];
                    drRow["lngMethodSN"] = dtMissAll.Rows[0]["lngMethodSN"];
                    drRow["lngDateSN"] = dtMissAll.Rows[0]["lngDateSN"];
                    drRow["intMiss"] = int.Parse(KeyValuePair.Key);
                    drRow["intMissCount"] = KeyValuePair.Value;
                    string strNums = "";
                    for (int intNum = 1; intNum <= Dataset00.LottoNumbers; intNum++)
                    {
                        if (dtMissAll.Rows[0][string.Format("lngM{0}", intNum.ToString())].ToString() == KeyValuePair.Key)
                        {
                            strNums += string.Format("{0}#", intNum.ToString("D2"));
                        }
                    }
                    drRow["strNums"] = strNums.TrimEnd('#');
                    dtDataShow.Rows.Add(drRow);
                }
                #endregion 填入資料

                #region Expander dgData

                #region DataGrid
                DataGrid dgMiss = new DataGrid()
                {
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Visible,
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    VerticalAlignment = VerticalAlignment.Top,
                    CanUserAddRows = false,
                    CanUserDeleteRows = false,
                    AutoGenerateColumns = false
                };
                #region Set Columns of DataGrid dgMiss
                DataGridTextColumn dgtColumnMiss;
                if (dgMiss.Columns.Count == 0)
                {
                    foreach (DataColumn dcdgColumn in dtDataShow.Columns)
                    {
                        dgtColumnMiss = new DataGridTextColumn()
                        {
                            Header = dcdgColumn.Caption,
                            Binding = new System.Windows.Data.Binding(dcdgColumn.ColumnName),
                            IsReadOnly = true
                        };
                        dgMiss.Columns.Add(dgtColumnMiss);
                    }
                }
                #endregion
                dgMiss.ItemsSource = dtDataShow.DefaultView;
                #endregion DataGrid

                Expander expMiss = new Expander()
                {
                    Header = string.Format(" {0}期 遺漏偏差表", intTrace.ToString()),
                    ExpandDirection = ExpandDirection.Down,
                    IsExpanded = true,
                    Background = Brushes.Aqua,
                    MaxHeight = 500,
                    Content = dgMiss
                };
                #endregion Expander
                spMain.Children.Add(expMiss);
            }

            #region ScrollViewer
            ScrollViewer svViewer = new ScrollViewer()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = spMain
            };
            #endregion ScrollViewer

            #region Window
            Window wMiss = new Window()
            {
                Name = "tblMiss",
                Title = string.Format("遺漏偏差表 ({0})", stuSearchMiss.StrTitle),
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.SingleBorderWindow,
                Content = svViewer
            };
            wMiss.Show();
            #endregion Window

        }

        #region btnTableMissAll 遺漏偏差總表
        private void BtnTableMissAll_Click(object sender, RoutedEventArgs e)
        {
            #region 設定 stuSearchMissAll
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSearchMissAll = stuRibbonSearchOption;
            #endregion 設定 stuSearchMissAll            

            #region Setting bwTablePercent
            bwBackgroundWorker00 = new BackgroundWorker()
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };
            bwBackgroundWorker00.DoWork += BwTableMissAll_DoWork;
            bwBackgroundWorker00.ProgressChanged += BwTableMissAll_ProgressChanged;
            bwBackgroundWorker00.RunWorkerCompleted += BwTableMissAll_RunWorkerCompleted;
            bwBackgroundWorker00.RunWorkerAsync(stuSearchMissAll);
            btnTableMissingTotal.Content = "計算中...";
            btnTableMissingTotal.IsEnabled = false;
            #endregion Setting bwTablePercent
        }
        private void BwTableMissAll_DoWork(object sender, DoWorkEventArgs e)
        {
            #region 設定
            StuGLSearch stuSearchMissAll = (StuGLSearch)e.Argument;
            BackgroundWorker bwMyWork = (BackgroundWorker)sender;
            bwMyWork.ReportProgress(1);
            Dictionary<string, object> dicMissAll = new Dictionary<string, object>
            {
                { "stuSearch", stuSearchMissAll },
                { "dicMissAll", new CGLMissAll().GetMissAlldic(stuSearchMissAll,3,"DESC") }
            };
            #endregion
            e.Result = dicMissAll;
        }
        private void BwTableMissAll_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 1)
            {
                this.tbStatusTextBlock.Text = "計算中...";
            }
        }
        private void BwTableMissAll_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            BackgroundWorker bwMywork = (BackgroundWorker)sender;
            Dictionary<string, object> dicArgument = (Dictionary<string, object>)e.Result;
            ShowTableMissAll(dicArgument);
            tbStatusTextBlock.Text = "Ready";
            btnTableMissingTotal.Content = "遺漏偏差總表";
            btnTableMissingTotal.IsEnabled = true;
            bwMywork.Dispose();
        }
        private void ShowTableMissAll(Dictionary<string, object> dicInput)
        {
            StuGLSearch stuSearch00 = (StuGLSearch)dicInput["stuSearch"];
            Dictionary<string, object> dicMissAll = (Dictionary<string, object>)dicInput["dicMissAll"];
            List<int> lstCurrentNums = new CGLSearch().GetlstCurrentNums(stuSearch00);
            #region spMain
            StackPanel spMain = new StackPanel()
            {
                Orientation = Orientation.Vertical,
                CanVerticallyScroll = true,
                CanHorizontallyScroll = true
            };
            #endregion

            #region Current Data
            spMain.Children.Add(new CGLSearch().GetCurrentDataSp(stuSearch00));
            #endregion

            #region ScrollViewer
            ScrollViewer svViewer = new ScrollViewer()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = spMain
            };
            #endregion ScrollViewer

            foreach (KeyValuePair<string, DataTable> KeyVal in (Dictionary<string, DataTable>)dicMissAll["MissAll"])
            {
                DataTable dtDataTable = KeyVal.Value;
                #region Convert 0 to -1 to -7
                for (int r = 0; r < dtDataTable.Rows.Count; r++)
                {
                    int intIndexofZero = -1;
                    for (int c = 1; c <= new CGLDataSet(stuSearch00.LottoType).LottoNumbers; c++)
                    {
                        string strColName = string.Format("lngM{0}", c.ToString());
                        if (int.Parse(dtDataTable.Rows[r][strColName].ToString()) == 0)
                        {
                            dtDataTable.Rows[r][strColName] = intIndexofZero;
                            intIndexofZero--;
                        }
                    }
                }
                #endregion
                #region Expander DataGrid
                dtDataTable.DefaultView.Sort = "lngTotalSN DESC";
                DataGrid dgMissAll = new DataGrid()
                {
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    RowDetailsVisibilityMode = DataGridRowDetailsVisibilityMode.Visible,
                    HorizontalAlignment = HorizontalAlignment.Left,
                    VerticalAlignment = VerticalAlignment.Top,
                    CanUserAddRows = false,
                    CanUserDeleteRows = false,
                    CanUserSortColumns = false,
                    CanUserResizeColumns = false,
                    CanUserResizeRows = false,
                    AutoGenerateColumns = false,
                    ItemsSource = dtDataTable.DefaultView,
                    FontSize = 12
                };
                #region Set Columns of DataGrid dgMissAll
                if (dgMissAll.Columns.Count == 0)
                {
                    DataGridTextColumn dgtColumnMissAll;
                    foreach (DataColumn dcdgColumn in dtDataTable.Columns)
                    {
                        dgtColumnMissAll = new DataGridTextColumn
                        {
                            Header = new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1),
                            Binding = new Binding(dcdgColumn.ColumnName),
                            IsReadOnly = true
                        };

                        #region change the lngL Background
                        if (dcdgColumn.ColumnName.Substring(0, 4) == "lngL")
                        {
                            Style Style00 = new Style()
                            {
                                TargetType = typeof(DataGridCell)
                            };
                            Setter setter00 = new Setter()
                            {
                                Property = BackgroundProperty,
                                Value = Brushes.LightYellow
                            };
                            Style00.Setters.Add(setter00);
                            setter00 = new Setter()
                            {
                                Property = ForegroundProperty,
                                Value = Brushes.Black
                            };
                            Style00.Setters.Add(setter00);
                            dgtColumnMissAll.CellStyle = Style00;
                        }
                        #endregion

                        if (dcdgColumn.ColumnName.Substring(0, 4) == "lngM" && dcdgColumn.ColumnName != "lngMethodSN")
                        {
                            #region Set Trigger
                            Style Style00 = new Style()
                            {
                                TargetType = typeof(DataGridCell)
                            };
                            for (int t = 1; t <= 7; t++)
                            {
                                DataTrigger dtTrig00 = new DataTrigger()
                                {
                                    Binding = dgtColumnMissAll.Binding,
                                    Value = string.Format("-{0}", t.ToString())
                                };
                                Setter setter00 = new Setter()
                                {
                                    Property = ForegroundProperty,
                                    Value = Brushes.Yellow
                                };
                                dtTrig00.Setters.Add(setter00);
                                setter00 = new Setter()
                                {
                                    Property = BackgroundProperty
                                };
                                switch (t)
                                {
                                    case 1:
                                        setter00.Value = Brushes.DarkGreen;
                                        break;
                                    case 2:
                                        setter00.Value = Brushes.DarkBlue;
                                        break;
                                    case 3:
                                        setter00.Value = Brushes.Red;
                                        break;
                                    case 4:
                                        setter00.Value = Brushes.DeepPink;
                                        break;
                                    case 5:
                                        setter00.Value = Brushes.DarkRed;
                                        break;
                                    case 6:
                                        setter00.Value = Brushes.Purple;
                                        break;
                                    case 7:
                                        setter00.Value = Brushes.Maroon;
                                        break;
                                }
                                dtTrig00.Setters.Add(setter00);
                                Style00.Triggers.Add(dtTrig00);
                            }

                            #endregion Set Trigger
                            if (int.Parse(dcdgColumn.ColumnName.Substring(4).ToString()) % 5 == 0)
                            {
                                Setter setter01 = new Setter()
                                {
                                    Property = BackgroundProperty,
                                    Value = Brushes.AliceBlue
                                };
                                Style00.Setters.Add(setter01);
                                setter01 = new Setter()
                                {
                                    Property = ForegroundProperty,
                                    Value = Brushes.Black
                                };
                                Style00.Setters.Add(setter01);
                            }
                            dgtColumnMissAll.CellStyle = Style00;
                            if (lstCurrentNums.Contains(int.Parse(dcdgColumn.ColumnName.Substring(4).ToString())))
                            {
                                Style Style01 = new Style()
                                {
                                    TargetType = typeof(DataGridColumnHeader)
                                };
                                Setter setter01 = new Setter()
                                {
                                    Property = BackgroundProperty,
                                };
                                int t = lstCurrentNums.IndexOf(int.Parse(dcdgColumn.ColumnName.Substring(4).ToString())) + 1;
                                switch (t)
                                {
                                    case 1:
                                        setter01.Value = Brushes.DarkGreen;
                                        break;
                                    case 2:
                                        setter01.Value = Brushes.DarkBlue;
                                        break;
                                    case 3:
                                        setter01.Value = Brushes.Red;
                                        break;
                                    case 4:
                                        setter01.Value = Brushes.DeepPink;
                                        break;
                                    case 5:
                                        setter01.Value = Brushes.DarkRed;
                                        break;
                                    case 6:
                                        setter01.Value = Brushes.Purple;
                                        break;
                                    case 7:
                                        setter01.Value = Brushes.Maroon;
                                        break;
                                }
                                Style01.Setters.Add(setter01);
                                setter01 = new Setter()
                                {
                                    Property = ForegroundProperty,
                                    Value = Brushes.White
                                };
                                Style01.Setters.Add(setter01);
                                dgtColumnMissAll.HeaderStyle = Style01;
                            }
                        }

                        dgMissAll.Columns.Add(dgtColumnMissAll);
                    }
                }
                #endregion
                dgMissAll.LoadingRow += DgMissAll_LoadingRow;

                Expander expdgMissAll = new Expander()
                {
                    Header = string.Format("遺漏偏差總表 {0} ({1})", KeyVal.Key, dtDataTable.Rows.Count),
                    ExpandDirection = ExpandDirection.Down,
                    IsExpanded = true,
                    Background = Brushes.Aqua,
                    MaxHeight = 700,
                    Content = dgMissAll
                    //expSumGraphic.Content = svViewerGraphic;
                };
                spMain.Children.Add(expdgMissAll);
                #endregion Expander DataGrid
            }

            #region Window
            Window wMissAll = new Window()
            {
                Name = "tblMissAll",
                Title = string.Format("遺漏偏差總表 ({0}) ", stuSearch00.StrTitle),
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.SingleBorderWindow,
                Content = svViewer
            };
            wMissAll.KeyDown += WMissAll_KeyDown;
            wMissAll.Show();
            #endregion Window

            #region Query String
            //StuGLData stuData00 = new StuGLData()
            //{
            //    LottoType = stuSearch00.LottoType,
            //    DataBaseType = DatabaseType.Miss,
            //    StrSELECT = "TOP 200 ",
            //    //stuSearch00.strSELECT = "";
            //    StrFROM = "[qryMissAll]",
            //    StrWHERE = string.Format("[{0}] = {1} AND [{2}] < {3} ",
            //                                    "lngMethodSN", stuSearch00.LngMethodSN,
            //                                    "lngTotalSN", stuSearch00.LngCurrentData),
            //    StrORDER = "[lngTotalSN] DESC"
            //};
            #region set select
            //stuData00.StrSELECT += "[lngTotalSN] , [lngMethodSN] , [lngDateSN] , ";
            //for (int i = 1; i <= new CGLDataSet(stuSearch00.LottoType).CountNumber; i++)
            //{
            //    stuData00.StrSELECT += string.Format("[lngL{0}] ,", i.ToString());
            //}
            //if (new CGLDataSet(stuSearch00.LottoType).DataType != TableType.LottoWeli && new CGLDataSet(stuSearch00.LottoType).SCountNumber > 0)
            //{
            //    stuData00.StrSELECT += " [lngS1] ,";
            //}
            //for (int i = 1; i <= new CGLDataSet(stuSearch00.LottoType).LottoNumbers; i++)
            //{
            //    stuData00.StrSELECT += string.Format(" [lngM{0}] ,", i.ToString());
            //}
            //if (new CGLDataSet(stuSearch00.LottoType).DataType == TableType.LottoWeli && new CGLDataSet(stuSearch00.LottoType).SCountNumber > 0)
            //{
            //    for (int i = 1; i <= new CGLDataSet(stuSearch00.LottoType).SCountNumber; i++)
            //    {
            //        stuData00.StrSELECT += string.Format(" [lngMS{0}] ,", i.ToString());
            //    }
            //}
            //stuData00.StrSELECT = stuData00.StrSELECT.TrimEnd(',');
            #endregion
            #region  Set Compare Option 
            ////if (stuSearch00.boolFieldMode && stuSearch00.strCompares != "gen")
            ////{
            ////    string[] strCompare = stuSearch00.strCompares.Split('#');
            ////    foreach (string strCompareOption in strCompare)
            ////    {
            ////        stuSearch00.strWHERE += string.Format(" {0} [{1}] = '{2}' "
            ////                                            , stuSearch00.strCompareType
            ////                                            , strCompareOption
            ////                                            , dicCurrentData[strCompareOption]);
            ////    }
            ////}
            #endregion
            #endregion Query String
            //DataTable dtDataTable = stuData00.GetSourceData(stuData00);
            //Dataset00.LottoNumbers
            #region Expander expMissAllGraphic 
            // for (int c = 1; c <= new CGLDataSet(stuSearch00.LottoType).LottoNumbers; c++)
            //  {
            //     string strColName = string.Format("lngM{0}", c.ToString());


            //#region Graphic MissAll

            //Chart chtGraphic = new Chart()
            //{
            //    Name = "gen",
            //    VerticalAlignment = VerticalAlignment.Stretch,
            //    HorizontalAlignment = HorizontalAlignment.Stretch,
            //    Height = 250,
            //    Width = 200 * 10
            //};
            //ColumnSeries csGraphic = new ColumnSeries()
            //{
            //    Title = strColName,
            //    ItemsSource = dtDataTable.DefaultView,
            //    IndependentValuePath = "lngDateSN",
            //    DependentValuePath = strColName,
            //    IsSelectionEnabled = true
            //};
            //chtGraphic.Series.Add(csGraphic);

            //#endregion Graphic

            //Expander expMissAllGraphic = new Expander()
            //{
            //    Header = strColName,
            //    ExpandDirection = ExpandDirection.Down,
            //    //expMissAllGraphic.IsExpanded = true;
            //    Background = Brushes.Aqua,
            //    MaxHeight = 700,
            //    Content = chtGraphic,
            //    IsExpanded = false
            //};
            //spMain.Children.Add(expMissAllGraphic);
            // }
            #endregion Expander Graphic

        }

        private void WMissAll_KeyDown(object sender, KeyEventArgs e)
        {
            Window wPercent = (Window)sender;
            ScrollViewer svViewer = (ScrollViewer)wPercent.Content;
            StackPanel spmain = (StackPanel)svViewer.Content;
            Expander expdgMissAll = (Expander)spmain.Children[1];
            DataGrid dgMissAll = (DataGrid)expdgMissAll.Content;
            if (e.Key == Key.P && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                PrintDialog pdPrint = new PrintDialog()
                {
                    CurrentPageEnabled = true,
                    SelectedPagesEnabled = true,
                    UserPageRangeEnabled = true
                };
                pdPrint.PrintTicket.PageOrientation = PageOrientation.Landscape;
                FixedDocument fdDocument = GetFixedDocument(dgMissAll, pdPrint);
                ShowPrintPreview(fdDocument);
                //                if (pdPrint.ShowDialog() == true)
                //                {
                // pdPrint.PrintTicket.PageOrientation = PageOrientation.Landscape;
                // pdPrint.PrintTicket.PageScalingFactor = 120;
                // DocumentPaginator dpPrint = ((IDocumentPaginatorSource)expdgMissAll.Content).DocumentPaginator;
                // pdPrint.PrintDocument(dpPrint, "Print");
                //pdPrint.PrintVisual(dvViewer, "Print");
                //}
            }
        }
        public static FixedDocument GetFixedDocument(FrameworkElement toPrint, PrintDialog printDialog)
        {
            PrintCapabilities capabilities = printDialog.PrintQueue.GetPrintCapabilities(printDialog.PrintTicket);
            Size pageSize = new Size(printDialog.PrintableAreaWidth, printDialog.PrintableAreaHeight);
            Size visibleSize = new Size(capabilities.PageImageableArea.ExtentWidth, capabilities.PageImageableArea.ExtentHeight);
            FixedDocument fixedDoc = new FixedDocument();
            //If the toPrint visual is not displayed on screen we neeed to measure and arrange it  
            toPrint.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            toPrint.Arrange(new Rect(new Point(0, 0), toPrint.DesiredSize));
            //  
            Size size = toPrint.DesiredSize;
            //Will assume for simplicity the control fits horizontally on the page  
            double yOffset = 0;
            while (yOffset < size.Height)
            {
                VisualBrush vb = new VisualBrush(toPrint)
                {
                    Stretch = Stretch.None,
                    AlignmentX = AlignmentX.Left,
                    AlignmentY = AlignmentY.Top,
                    ViewboxUnits = BrushMappingMode.Absolute,
                    TileMode = TileMode.None,
                    Viewbox = new Rect(0, yOffset, visibleSize.Width, visibleSize.Height)
                };
                PageContent pageContent = new PageContent();
                FixedPage page = new FixedPage();
                ((IAddChild)pageContent).AddChild(page);
                fixedDoc.Pages.Add(pageContent);
                page.Width = pageSize.Width;
                page.Height = pageSize.Height;
                Canvas canvas = new Canvas();
                FixedPage.SetLeft(canvas, capabilities.PageImageableArea.OriginWidth);
                FixedPage.SetTop(canvas, capabilities.PageImageableArea.OriginHeight);
                canvas.Width = visibleSize.Width;
                canvas.Height = visibleSize.Height;
                canvas.Background = vb;
                page.Children.Add(canvas);
                yOffset += visibleSize.Height;
            }
            return fixedDoc;
        }

        public static void ShowPrintPreview(FixedDocument fixedDoc)
        {
            Window wnd = new Window();
            DocumentViewer viewer = new DocumentViewer()
            {
                Document = fixedDoc
            };
            wnd.Content = viewer;
            wnd.ShowDialog();
        }

        private void DgMissAll_LoadingRow(object sender, DataGridRowEventArgs e)
        {

            DataGridRow dgrRow = e.Row;

            DataRowView drvGridRow = (DataRowView)dgrRow.Item;
            string strDateSN = drvGridRow.Row.ItemArray[2].ToString();
            int intYear = int.Parse(strDateSN.Substring(0, 4));
            int intMonth = int.Parse(strDateSN.Substring(4, 2));
            int intDay = int.Parse(strDateSN.Substring(6, 2));
            DateTime dtime = new DateTime(intYear, intMonth, intDay);
            if (dtime.DayOfWeek == DayOfWeek.Saturday)
            {
                dgrRow.Background = Brushes.MistyRose;
            }
            else
            {
                dgrRow.Background = Brushes.White;
            }
        }
        #endregion btnTableMissAll

        #region btnTableMissAll01 遺漏偏差總表01
        private void BtnTableMissAll01_Click(object sender, RoutedEventArgs e)
        {
            #region 設定 stuSearchMissAll01
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSearchMissAll01 = stuRibbonSearchOption;
            stuSearchMissAll01.IntMatchMin = 0;
            stuSearchMissAll01.IntDataOffset = 0;
            stuSearchMissAll01.IntSearchLimit = 0;
            stuSearchMissAll01.IntSearchOffset = 0;
            stuSearchMissAll01.StrCompares = "gen";
            stuSearchMissAll01.BoolFieldMode = false;
            stuSearchMissAll01 = new CGLSearch().GetMethodSN(stuSearchMissAll01);
            #endregion 設定 stuSearchMissAll

            #region Setting bwTablePercent
            bwBackgroundWorker00 = new BackgroundWorker()
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };
            bwBackgroundWorker00.DoWork += BwTableMissAll01_DoWork;
            bwBackgroundWorker00.ProgressChanged += BwTableMissAll01_ProgressChanged;
            bwBackgroundWorker00.RunWorkerCompleted += BwTableMissAll01_RunWorkerCompleted;
            #endregion Setting bwTablePercent
            bwBackgroundWorker00.RunWorkerAsync(stuSearchMissAll01);
            btnTableMissingTotal01.Content = "計算中...";
            btnTableMissingTotal01.IsEnabled = false;
        }

        private void BwTableMissAll01_DoWork(object sender, DoWorkEventArgs e)
        {
            #region 設定
            StuGLSearch stuSearchMissAll01 = (StuGLSearch)e.Argument;
            BackgroundWorker bwMyWork = (BackgroundWorker)sender;
            bwMyWork.ReportProgress(1);
            Dictionary<string, object> dicArgument = new Dictionary<string, object>
            {
                { "stuSearch", stuSearchMissAll01 }
            };
            #endregion
            #region 檢查 MissAll01 是否有資料 
            if (!new CGLMissAll().HasMissAll01Data(stuSearchMissAll01))
            {
                new CGLMissAll().SearchMissAll01(stuSearchMissAll01);
            }
            #endregion 檢查是否有資料
            e.Result = dicArgument;
        }

        private void BwTableMissAll01_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 1)
            {
                this.tbStatusTextBlock.Text = "計算中...";
            }
        }

        private void BwTableMissAll01_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            BackgroundWorker bwMywork = (BackgroundWorker)sender;
            Dictionary<string, object> dicArgument = (Dictionary<string, object>)e.Result;
            ShowTableMissAll01((StuGLSearch)dicArgument["stuSearch"]);
            tbStatusTextBlock.Text = "Ready";
            btnTableMissingTotal01.Content = "遺漏偏差總表01";
            btnTableMissingTotal01.IsEnabled = true;
            bwMywork.Dispose();
        }

        private void ShowTableMissAll01(StuGLSearch stuGLSearch)
        {


        }

        #endregion btnTableMissAll01

        #region btnTablePercent 百分比預測表
        private void BtnTablePercent_Click(object sender, RoutedEventArgs e)
        {
            #region 設定 stuSearchPercent
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSearchPercent = stuRibbonSearchOption;
            stuSearchPercent.IntMatchMin = 0;
            stuSearchPercent.IntDataOffset = 0;
            stuSearchPercent.IntSearchLimit = 0;
            stuSearchPercent.IntSearchOffset = 0;
            stuSearchPercent.StrCompares = "gen";
            stuSearchPercent.BoolFieldMode = false;
            stuSearchPercent = new CGLSearch().GetMethodSN(stuSearchPercent);
            #endregion 設定 stuSearchPercent
            #region Setting bwTablePercent
            bwBackgroundWorker00 = new BackgroundWorker()
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };
            bwBackgroundWorker00.DoWork += BwTablePercent_DoWork;
            bwBackgroundWorker00.ProgressChanged += BwTablePercent_ProgressChanged;
            bwBackgroundWorker00.RunWorkerCompleted += BwTablePercent_RunWorkerCompleted;
            #endregion Setting bwTablePercent
            bwBackgroundWorker00.RunWorkerAsync(stuSearchPercent);
            btnTablePercent.Content = "計算中...";
            btnTablePercent.IsEnabled = false;
        }
        private void BwTablePercent_DoWork(object sender, DoWorkEventArgs e)
        {
            #region 設定
            StuGLSearch stuSearchPercent = (StuGLSearch)e.Argument;
            BackgroundWorker bwMyWork = (BackgroundWorker)sender;
            bwMyWork.ReportProgress(1);
            Dictionary<string, object> dicArgument = new Dictionary<string, object>
            {
                { "stuSearch", stuSearchPercent }
            };
            #endregion
            #region Check Does Percent.Html exist ?
            string strCurrentDirectory = System.IO.Directory.GetCurrentDirectory();
            string strHtmlDirectory = System.IO.Path.Combine(strCurrentDirectory, "html");
            Dictionary<string, string> dicCurrentData = new CGLSearch().GetCurrentData(stuSearchPercent);
            string strFileName = string.Format("Percent_{0}{1}.html", new CGLDataSet(stuSearchPercent.LottoType).ID, dicCurrentData["lngDateSN"]);
            bool boolFileExist = new CGLFunc().FileExist(strHtmlDirectory, strFileName);
            #endregion
            if (stuSearchPercent.BoolRecalc || !boolFileExist)
            {
                Dictionary<string, object> dicCompares = new CGLSearch().GetTablePercent(stuSearchPercent);
                dicArgument.Add("Compares", dicCompares);
            }
            e.Result = dicArgument;
        }
        private void BwTablePercent_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            Dictionary<string, object> dicArgument = (Dictionary<string, object>)e.Result;
            StuGLSearch stuSearchPercent = (StuGLSearch)dicArgument["stuSearch"];

            if (dicArgument.ContainsKey("Compares"))
            {
                ShowTablePercent(stuSearchPercent, (Dictionary<string, object>)dicArgument["Compares"]);
                CreatPercentPage(stuSearchPercent, (Dictionary<string, object>)dicArgument["Compares"]);
            }
            else
            {
                string strCurrentDirectory = System.IO.Directory.GetCurrentDirectory();
                string strHtmlDirectory = System.IO.Path.Combine(strCurrentDirectory, "html");
                Dictionary<string, string> dicCurrentData = new CGLSearch().GetCurrentData(stuSearchPercent);
                string strFileName = string.Format("Percent_{0}{1}.html", new CGLDataSet(stuSearchPercent.LottoType).ID, dicCurrentData["lngDateSN"]);
                Process.Start(System.IO.Path.Combine(strHtmlDirectory, strFileName));
            }
            tbStatusTextBlock.Text = "Ready";
            btnTablePercent.Content = "百分比預測表";
            btnTablePercent.IsEnabled = true;
            BackgroundWorker bwMywork = (BackgroundWorker)sender;
            bwMywork.Dispose();
        }
        private void BwTablePercent_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 1)
            {
                this.tbStatusTextBlock.Text = "計算中...";
            }
        }
        private void DgHot_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            DataGridRow dgrRow = e.Row;
            DataRowView drvGridRow = (DataRowView)dgrRow.Item;
            if (int.Parse(drvGridRow.Row.ItemArray[5].ToString()) >= 4)
            {
                e.Row.Background = Brushes.LightSeaGreen;
            }
            if (int.Parse(drvGridRow.Row.ItemArray[1].ToString()) >= 3)
            {
                e.Row.Background = Brushes.LightPink;
            }
            if (int.Parse(drvGridRow.Row.ItemArray[1].ToString()) == 0 && int.Parse(drvGridRow.Row.ItemArray[5].ToString()) >= 4)
            {
                e.Row.Background = Brushes.LightGreen;
            }
        }
        private void ShowTablePercent(StuGLSearch stuSearch00, Dictionary<string, object> dicCpmpares)
        {
            #region 設定
            CGLDataSet Dataset00 = new CGLDataSet(stuSearch00.LottoType);
            Dictionary<string, object> dicTotal = (Dictionary<string, object>)dicCpmpares["dicTotal"];
            #region Set lstCurreny
            Dictionary<string, string> dicCurrentData = new CGLSearch().GetCurrentData(stuSearch00);
            List<int> lstCurrentData = new CGLSearch().GetlstCurrentNums(stuSearch00);
            #endregion Set lstCurreny

            #endregion

            #region Set Total MissAll
            Dictionary<string, int> dicTMA = (Dictionary<string, int>)dicTotal["dicTMA"];
            DataTable dtTMA = new CGLFunc().CDicTOTable(new CGLFunc().CDic_intTOstring(dicTMA));
            DataGrid dgTMA = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                RowDetailsVisibilityMode = DataGridRowDetailsVisibilityMode.Visible,
                CanUserAddRows = true,
                CanUserDeleteRows = false,
                CanUserResizeColumns = true,
                CanUserResizeRows = true,
                CanUserSortColumns = true,
                IsReadOnly = true,
                AutoGenerateColumns = false,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch,
                ItemsSource = dtTMA.DefaultView
            };

            #region Set Columns of DataGrid dgCurrentData 
            DataGridTextColumn dgtCol;
            if (dgTMA.Columns.Count == 0)
            {
                foreach (DataColumn dcCol in dtTMA.Columns)
                {
                    dgtCol = new DataGridTextColumn()
                    {
                        Header = dcCol.ColumnName,
                        Binding = new Binding(dcCol.ColumnName)
                    };
                    if (lstCurrentData.Contains(int.Parse(dcCol.ColumnName)))
                    {
                        Style styStyle00 = new Style();
                        Setter setter = new Setter()
                        {
                            Property = BackgroundProperty,
                            Value = Brushes.Red
                        };
                        styStyle00.Setters.Add(setter);
                        setter = new Setter()
                        {
                            Property = ForegroundProperty,
                            Value = Brushes.Yellow
                        };
                        styStyle00.Setters.Add(setter);
                        setter = new Setter()
                        {
                            Property = HorizontalAlignmentProperty,
                            Value = HorizontalAlignment.Center
                        };
                        styStyle00.Setters.Add(setter);
                        dgtCol.HeaderStyle = styStyle00;
                    }
                    dgtCol.IsReadOnly = true;
                    dgTMA.Columns.Add(dgtCol);
                }
            }
            #endregion
            #endregion

            #region Set dgT0105
            Dictionary<string, int> dicT0105 = (Dictionary<string, int>)dicTotal["dicT0105"];
            DataTable dtT0105 = new CGLFunc().CDicTOTable(new CGLFunc().CDic_intTOstring(dicT0105));

            DataGrid dgT0105 = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                RowDetailsVisibilityMode = DataGridRowDetailsVisibilityMode.Visible,
                CanUserAddRows = true,
                CanUserDeleteRows = false,
                CanUserResizeColumns = true,
                CanUserResizeRows = true,
                CanUserSortColumns = true,
                IsReadOnly = true,
                AutoGenerateColumns = false,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch,
                ItemsSource = dtT0105.DefaultView
            };

            #region Set Columns of DataGrid dgCurrentData 
            if (dgT0105.Columns.Count == 0)
            {
                foreach (DataColumn dcCol in dtT0105.Columns)
                {
                    dgtCol = new DataGridTextColumn()
                    {
                        Header = dcCol.ColumnName,
                        Binding = new Binding(dcCol.ColumnName)
                    };
                    if (lstCurrentData.Contains(int.Parse(dcCol.ColumnName)))
                    {
                        Style styStyle00 = new Style();
                        Setter setter = new Setter()
                        {
                            Property = BackgroundProperty,
                            Value = Brushes.Red
                        };
                        styStyle00.Setters.Add(setter);
                        setter = new Setter()
                        {
                            Property = ForegroundProperty,
                            Value = Brushes.Yellow
                        };
                        styStyle00.Setters.Add(setter);
                        setter = new Setter()
                        {
                            Property = HorizontalAlignmentProperty,
                            Value = HorizontalAlignment.Center
                        };
                        styStyle00.Setters.Add(setter);
                        dgtCol.HeaderStyle = styStyle00;
                    }
                    dgtCol.IsReadOnly = true;
                    dgT0105.Columns.Add(dgtCol);
                }
            }
            #endregion
            #endregion

            #region Set dgT0125
            Dictionary<string, int> dicT0125 = (Dictionary<string, int>)dicTotal["dicT0125"];
            DataTable dtT0125 = new CGLFunc().CDicTOTable(new CGLFunc().CDic_intTOstring(dicT0125));

            DataGrid dgT0125 = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                RowDetailsVisibilityMode = DataGridRowDetailsVisibilityMode.Visible,
                CanUserAddRows = true,
                CanUserDeleteRows = false,
                CanUserResizeColumns = true,
                CanUserResizeRows = true,
                CanUserSortColumns = true,
                IsReadOnly = true,
                AutoGenerateColumns = false,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch,
                ItemsSource = dtT0125.DefaultView
            };
            #region Set Columns of DataGrid dgCurrentData 
            if (dgT0125.Columns.Count == 0)
            {
                foreach (DataColumn dcCol in dtT0125.Columns)
                {
                    dgtCol = new DataGridTextColumn()
                    {
                        Header = dcCol.ColumnName,
                        Binding = new Binding(dcCol.ColumnName)
                    };
                    if (lstCurrentData.Contains(int.Parse(dcCol.ColumnName)))
                    {
                        Style styStyle00 = new Style();
                        Setter setter = new Setter()
                        {
                            Property = BackgroundProperty,
                            Value = Brushes.Red
                        };
                        styStyle00.Setters.Add(setter);
                        setter = new Setter()
                        {
                            Property = ForegroundProperty,
                            Value = Brushes.Yellow
                        };
                        styStyle00.Setters.Add(setter);
                        setter = new Setter()
                        {
                            Property = HorizontalAlignmentProperty,
                            Value = HorizontalAlignment.Center
                        };
                        styStyle00.Setters.Add(setter);
                        dgtCol.HeaderStyle = styStyle00;
                    }
                    dgtCol.IsReadOnly = true;
                    dgT0125.Columns.Add(dgtCol);
                }
            }
            #endregion
            #endregion

            #region Set dgAppear
            Dictionary<string, int> dicAppear = (Dictionary<string, int>)dicTotal["dicAppear"];
            DataTable dtAppear = new CGLFunc().CDicTOTable(new CGLFunc().CDic_intTOstring(dicAppear));

            DataGrid dgAppear = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                RowDetailsVisibilityMode = DataGridRowDetailsVisibilityMode.Visible,
                CanUserAddRows = true,
                CanUserDeleteRows = false,
                CanUserResizeColumns = true,
                CanUserResizeRows = true,
                CanUserSortColumns = true,
                IsReadOnly = true,
                AutoGenerateColumns = false,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                ItemsSource = dtAppear.DefaultView
            };



            #region Set Columns of DataGrid dgAppear 
            if (dgAppear.Columns.Count == 0)
            {
                foreach (DataColumn dcCol in dtAppear.Columns)
                {
                    dgtCol = new DataGridTextColumn()
                    {
                        Header = dcCol.ColumnName,
                        Binding = new Binding(dcCol.ColumnName)
                    };
                    if (lstCurrentData.Contains(int.Parse(dcCol.ColumnName)))
                    {
                        Style styStyle00 = new Style();
                        Setter setter = new Setter()
                        {
                            Property = BackgroundProperty,
                            Value = Brushes.Red
                        };
                        styStyle00.Setters.Add(setter);
                        setter = new Setter()
                        {
                            Property = ForegroundProperty,
                            Value = Brushes.Yellow
                        };
                        styStyle00.Setters.Add(setter);
                        setter = new Setter()
                        {
                            Property = HorizontalAlignmentProperty,
                            Value = HorizontalAlignment.Center
                        };
                        styStyle00.Setters.Add(setter);
                        dgtCol.HeaderStyle = styStyle00;
                    }
                    dgtCol.IsReadOnly = true;
                    dgAppear.Columns.Add(dgtCol);
                }
            }
            #endregion
            #endregion

            #region spMain
            StackPanel spMain = new StackPanel()
            {
                Orientation = Orientation.Vertical,
                CanVerticallyScroll = true,
                CanHorizontallyScroll = true
            };
            spMain.Children.Add(new CGLSearch().GetCurrentDataSp(stuSearch00));
            Label lblTMA = new Label()
            {
                Content = "Total Miss All :"
            };
            spMain.Children.Add(lblTMA);
            spMain.Children.Add(dgTMA);
            spMain.Children.Add(dgT0105);
            spMain.Children.Add(dgT0125);
            spMain.Children.Add(dgAppear);
            #endregion StackPanel

            #region Each Compare

            foreach (KeyValuePair<string, object> Items in dicCpmpares)
            {
                if (Items.Key != "dicTotal")
                {
                    Dictionary<string, object> dicEachCompare = (Dictionary<string, object>)Items.Value;
                    DataSet dsFreq = (DataSet)dicEachCompare["dsFreq"];
                    #region spOutLine
                    StackPanel spOutLine = new StackPanel()
                    {
                        Orientation = Orientation.Vertical,
                        CanVerticallyScroll = true,
                        CanHorizontallyScroll = true
                    };
                    #endregion StackPanel
                    #region Frequency Part

                    #region spFreq
                    StackPanel spFreq = new StackPanel()
                    {
                        Orientation = Orientation.Vertical,
                        CanVerticallyScroll = true,
                        CanHorizontallyScroll = true
                    };
                    #endregion StackPanel

                    foreach (DataTable dtFreq in dsFreq.Tables)
                    {
                        if (dtFreq.TableName != "DataAll" && dtFreq.TableName != "MissAll")
                        {
                            #region DataGrid dgFreq
                            DataGrid dgFreq = new DataGrid()
                            {
                                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                                HorizontalAlignment = HorizontalAlignment.Left,
                                VerticalAlignment = VerticalAlignment.Top,
                                CanUserAddRows = false,
                                CanUserDeleteRows = false,
                                CanUserResizeColumns = false,
                                CanUserSortColumns = false,
                                AutoGenerateColumns = false,
                                ItemsSource = dtFreq.DefaultView,
                                Name = string.Format("dg{0}", dtFreq.TableName)
                            };

                            #region Set Columns of DataGrid dgActive
                            DataGridTextColumn dgtColumnActive;
                            if (dgFreq.Columns.Count == 0)
                            {
                                foreach (DataColumn dcColumn in dtFreq.Columns)
                                {
                                    dgtColumnActive = new DataGridTextColumn
                                    {
                                        Header = new CGLFunc().ConvertFieldNameID(dcColumn.Caption, 1),
                                        Binding = new Binding(dcColumn.ColumnName),
                                        IsReadOnly = true
                                    };
                                    if (lstCurrentData.Contains(int.Parse(dcColumn.ColumnName.Substring(4))))
                                    {
                                        Style styStyle00 = new Style();
                                        Setter setter = new Setter()
                                        {
                                            Property = BackgroundProperty,
                                            Value = Brushes.Red
                                        };
                                        styStyle00.Setters.Add(setter);
                                        setter = new Setter()
                                        {
                                            Property = ForegroundProperty,
                                            Value = Brushes.Yellow
                                        };
                                        styStyle00.Setters.Add(setter);
                                        setter = new Setter()
                                        {
                                            Property = HorizontalAlignmentProperty,
                                            Value = HorizontalAlignment.Center
                                        };
                                        styStyle00.Setters.Add(setter);
                                        dgtColumnActive.HeaderStyle = styStyle00;
                                    }
                                    dgFreq.Columns.Add(dgtColumnActive);
                                }
                            }
                            #endregion
                            #endregion
                            spFreq.Children.Add(dgFreq);
                        }
                    }

                    #endregion Frequency Part
                    #region Hot Part

                    DataSet dsHot = (DataSet)dicEachCompare["dsHot"];

                    #region spHot
                    StackPanel spHot = new StackPanel()
                    {
                        Orientation = Orientation.Horizontal,
                        CanVerticallyScroll = true,
                        CanHorizontallyScroll = true
                    };
                    #endregion StackPanel

                    foreach (DataTable dtHot in dsHot.Tables)
                    {
                        #region DataGrid dgHot
                        DataGrid dgHot = new DataGrid()
                        {
                            HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                            HorizontalAlignment = HorizontalAlignment.Left,
                            VerticalAlignment = VerticalAlignment.Top,
                            CanUserAddRows = false,
                            CanUserDeleteRows = false,
                            CanUserResizeColumns = false,
                            CanUserSortColumns = false,
                            AutoGenerateColumns = false,
                            Name = string.Format("dg{0}", dtHot.TableName),
                            ItemsSource = dtHot.DefaultView
                        };
                        dtHot.DefaultView.Sort = "[0115] DESC , [0105] DESC , [Nums] ASC ";
                        dgHot.LoadingRow += DgHot_LoadingRow;

                        #region Set Columns of DataGrid dgActive
                        DataGridTextColumn dgtColumnActive;
                        if (dgHot.Columns.Count == 0)
                        {
                            int intColumnCount = 0;
                            foreach (DataColumn dcdgColumn in dtHot.Columns)
                            {
                                dgtColumnActive = new DataGridTextColumn
                                {
                                    Header = new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1),
                                    Binding = new Binding(dcdgColumn.ColumnName),
                                    IsReadOnly = true
                                };
                                if (intColumnCount == 0)
                                {
                                    #region Set Trigger
                                    Style Style00 = new Style()
                                    {
                                        TargetType = typeof(DataGridCell)
                                    };
                                    foreach (int intNum in lstCurrentData)
                                    {
                                        DataTrigger dtTrig00 = new DataTrigger()
                                        {
                                            Binding = dgtColumnActive.Binding,
                                            Value = string.Format("{0}", intNum.ToString("D2"))
                                        };
                                        Setter setter00 = new Setter()
                                        {
                                            Property = ForegroundProperty,
                                            Value = Brushes.Yellow
                                        };
                                        dtTrig00.Setters.Add(setter00);
                                        setter00 = new Setter()
                                        {
                                            Property = BackgroundProperty,
                                            Value = Brushes.Red
                                        };
                                        dtTrig00.Setters.Add(setter00);
                                        Style00.Triggers.Add(dtTrig00);
                                    }
                                    #endregion Set Trigger
                                    dgtColumnActive.CellStyle = Style00;

                                }
                                dgHot.Columns.Add(dgtColumnActive);
                                intColumnCount++;
                            }
                        }
                        #endregion
                        #endregion
                        spHot.Children.Add(dgHot);
                    }

                    #endregion Hot Part
                    spOutLine.Children.Add(spFreq);
                    spOutLine.Children.Add(spHot);
                    #region Expander expOutLine
                    Expander expOutLine = new Expander()
                    {
                        Background = Brushes.Gray,
                        IsExpanded = true
                    };
                    #region Set Title 
                    string strTitle = string.Format("{0}: ", Dataset00.LottoDescription);
                    StuGLSearch testdate = new StuGLSearch();
                    testdate = stuSearch00;
                    testdate.LngCurrentData = testdate.LngDataStart;
                    dicCurrentData = new Dictionary<string, string>();
                    dicCurrentData = new CGLSearch().GetCurrentData(testdate);
                    strTitle += dicCurrentData["lngDateSN"] + " => ";
                    testdate.LngCurrentData = testdate.LngDataEnd;
                    dicCurrentData = new CGLSearch().GetCurrentData(testdate);
                    strTitle += dicCurrentData["lngDateSN"];

                    if (Items.Key != "gen")
                    {
                        if (Items.Key.Length > 0)
                        {
                            string[] strCompare = Items.Key.Split('#');
                            strTitle += " 相同 {";
                            foreach (string strCompareOption in strCompare)
                            {
                                strTitle += " " + new CGLFunc().ConvertFieldNameID(strCompareOption, 1) + " ,";
                            }
                            strTitle = strTitle.TrimEnd(',') + " }";
                        }
                    }
                    strTitle += string.Format("({0:00}期)", dsFreq.Tables["DataAll"].Rows.Count);
                    #endregion
                    if (dsFreq.Tables["DataAll"].Rows.Count >= 20)
                    {
                        //expOutLine.IsExpanded = true;
                        expOutLine.Background = Brushes.Aqua;
                    }
                    expOutLine.Header = strTitle;
                    expOutLine.ExpandDirection = ExpandDirection.Down;
                    expOutLine.Content = spOutLine;
                    #endregion Expander expOutLine
                    spMain.Children.Add(expOutLine);
                }
            }
            #endregion Each Compare

            #region Scroll
            ScrollViewer svViewer = new ScrollViewer()
            {
                Name = "svViewer",
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = spMain
            };
            #endregion

            #region Window
            Window wPercent = new Window()
            {
                Name = "tblPercent",
                Title = string.Format("百分比預測表({0})", dicCurrentData["lngDateSN"]),
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.SingleBorderWindow,

                Content = svViewer,
                Width = Dataset00.LottoNumbers * 30
            };
            if (wPercent.Width > 1280) wPercent.Width = 1280;
            wPercent.KeyDown += WPercent_KeyDown;
            wPercent.Show();
            #endregion Window

            #region Change the color of DataGrid's row

            //foreach (var KeyValue in dicCompares)
            //{
            //    Dictionary<string, object> dicDataGrid = (Dictionary<string, object>)KeyValue.Value;
            //    int intdgCount = 0;
            //    foreach (var Var in dicDataGrid)
            //    {
            //        if (intdgCount >= 4 && intdgCount <= 8)
            //        {
            //            DataGrid dgHot00 = (DataGrid)dicDataGrid[Var.Key];
            //            if (dgHot00.Items.Count > 0)
            //            {
            //                foreach (DataRowView drvGridRow in dgHot00.ItemsSource)
            //                {
            //                    DataGridRow row = new DataGridRow();
            //                    row = (DataGridRow)dgHot00.ItemContainerGenerator.ContainerFromItem(drvGridRow);
            //                    if (int.Parse(drvGridRow.Row.ItemArray[5].ToString()) >= 4)
            //                    {
            //                          row.Background = Brushes.LightSeaGreen;
            //                    }
            //                    if (int.Parse(drvGridRow.Row.ItemArray[1].ToString()) >= 3)
            //                    {
            //                        row.Background = Brushes.LightPink;
            //                    }
            //                    if (int.Parse(drvGridRow.Row.ItemArray[1].ToString()) == 0 && int.Parse(drvGridRow.Row.ItemArray[5].ToString()) >= 4)
            //                    {
            //                        row.Background = Brushes.LightGreen;
            //                    }
            //                }
            //            }
            //        }
            //        intdgCount++;
            //    }
            //}

            #endregion Change the color of DataGrid's row

        }
        private void CreatPercentPage(StuGLSearch stuSearch00, Dictionary<string, object> dicCpmpares)
        {
            #region Setting
            CGLDataSet DataSet00 = new CGLDataSet(stuSearch00.LottoType);
            Dictionary<string, object> dicTotal = (Dictionary<string, object>)dicCpmpares["dicTotal"];
            Dictionary<string, int> dicCurrentDataNums = new CGLSearch().GetCurrentDataNums(stuSearch00);
            #region Get File Name
            string strCurrentDirectory = System.IO.Directory.GetCurrentDirectory();
            string strHtmlDirectory = System.IO.Path.Combine(strCurrentDirectory, "html");
            Dictionary<string, string> dicCurrentData = new CGLSearch().GetCurrentData(stuSearch00);
            string strFileName = string.Format("Percent_{0}{1}.html", DataSet00.ID, dicCurrentData["lngDateSN"]);
            #endregion Get File Name
            #region Setting Title
            string strPageTitle = string.Format("Percent{0}_{1}", DataSet00.LottoDescription, dicCurrentData["lngDateSN"]);
            #endregion Setting Title
            #endregion Setting

            StringBuilder stringbuilder = new StringBuilder();
            stringbuilder.AppendLine("<!DOCTYPE html>");
            stringbuilder.AppendLine("<html lang = 'zh-tw' xmlns = 'http://www.w3.org/1999/xhtml' >");
            stringbuilder.AppendLine("<head>");
            stringbuilder.AppendLine(string.Format("<title>{0}</title>", strPageTitle));
            stringbuilder.AppendLine("<meta charset = 'utf-8' />");
            stringbuilder.AppendLine("<link rel = 'stylesheet' href = '../css/glmain.css' />");
            stringbuilder.AppendLine("</head>");
            stringbuilder.AppendLine("<body>");
            #region Set Current Data
            stringbuilder.AppendLine(new CGLSearch().GetCurrentDataHtml(stuSearch00));
            #endregion Set Set Current Data
            stringbuilder.AppendLine("<details class='PercentAll' open>");
            stringbuilder.AppendLine(string.Format("<summary>{0}</summary>", "PercentAll"));
            #region Set Total MissAll
            Dictionary<string, int> dicTMA = (Dictionary<string, int>)dicTotal["dicTMA"];
            stringbuilder.AppendLine("<table>");
            stringbuilder.AppendLine(string.Format("<caption>{0}</caption>", "Total Miss All"));
            stringbuilder.AppendLine("<thead>");
            stringbuilder.AppendLine("<tr>");
            foreach (var keypair in dicTMA)
            {
                if (dicCurrentDataNums.ContainsValue(int.Parse(keypair.Key)))
                {
                    stringbuilder.AppendLine(string.Format("<th class='{0}'>", dicCurrentDataNums.FirstOrDefault(x => x.Value == (int.Parse(keypair.Key))).Key));
                }
                else
                {
                    stringbuilder.AppendLine("<th>");
                }
                stringbuilder.AppendLine(keypair.Key);
                stringbuilder.AppendLine("</th>");
            }
            stringbuilder.AppendLine("</tr>");
            stringbuilder.AppendLine("</thead>");
            stringbuilder.AppendLine("<tbody>");
            stringbuilder.AppendLine("<tr>");
            foreach (var keypair in dicTMA)
            {
                stringbuilder.AppendLine("<td>");
                stringbuilder.AppendLine(keypair.Value.ToString());
                stringbuilder.AppendLine("</td>");
            }
            stringbuilder.AppendLine("</tr>");
            stringbuilder.AppendLine("</tbody>");
            stringbuilder.AppendLine("</table>");
            #endregion Set Total MissAll
            #region Set dicT0105
            Dictionary<string, int> dicT0105 = (Dictionary<string, int>)dicTotal["dicT0105"];
            stringbuilder.AppendLine("<table>");
            stringbuilder.AppendLine(string.Format("<caption>{0}</caption>", "Total 01-05"));
            stringbuilder.AppendLine("<thead>");
            stringbuilder.AppendLine("<tr>");
            foreach (var keypair in dicT0105)
            {
                if (dicCurrentDataNums.ContainsValue(int.Parse(keypair.Key)))
                {
                    stringbuilder.AppendLine(string.Format("<th class='{0}'>", dicCurrentDataNums.FirstOrDefault(x => x.Value == (int.Parse(keypair.Key))).Key));
                }
                else
                {
                    stringbuilder.AppendLine("<th>");
                }
                stringbuilder.AppendLine(keypair.Key);
                stringbuilder.AppendLine("</th>");
            }
            stringbuilder.AppendLine("</tr>");
            stringbuilder.AppendLine("</thead>");
            stringbuilder.AppendLine("<tbody>");
            stringbuilder.AppendLine("<tr>");
            foreach (var keypair in dicT0105)
            {
                stringbuilder.AppendLine("<td>");
                stringbuilder.AppendLine(keypair.Value.ToString());
                stringbuilder.AppendLine("</td>");
            }
            stringbuilder.AppendLine("</tr>");
            stringbuilder.AppendLine("</tbody>");
            stringbuilder.AppendLine("</table>");
            #endregion Set dicT0105
            #region Set dicT0125
            Dictionary<string, int> dicT0125 = (Dictionary<string, int>)dicTotal["dicT0125"];
            stringbuilder.AppendLine("<table>");
            stringbuilder.AppendLine(string.Format("<caption>{0}</caption>", "Total 01-25"));
            stringbuilder.AppendLine("<thead>");
            stringbuilder.AppendLine("<tr>");
            foreach (var keypair in dicT0125)
            {
                if (dicCurrentDataNums.ContainsValue(int.Parse(keypair.Key)))
                {
                    stringbuilder.AppendLine(string.Format("<th class='{0}'>", dicCurrentDataNums.FirstOrDefault(x => x.Value == (int.Parse(keypair.Key))).Key));
                }
                else
                {
                    stringbuilder.AppendLine("<th>");
                }
                stringbuilder.AppendLine(keypair.Key);
                stringbuilder.AppendLine("</th>");
            }
            stringbuilder.AppendLine("</tr>");
            stringbuilder.AppendLine("</thead>");
            stringbuilder.AppendLine("<tbody>");
            stringbuilder.AppendLine("<tr>");
            foreach (var keypair in dicT0125)
            {
                stringbuilder.AppendLine("<td>");
                stringbuilder.AppendLine(keypair.Value.ToString());
                stringbuilder.AppendLine("</td>");
            }
            stringbuilder.AppendLine("</tr>");
            stringbuilder.AppendLine("</tbody>");
            stringbuilder.AppendLine("</table>");
            #endregion Set dicT0125
            #region Set dicAppear
            Dictionary<string, int> dicAppear = (Dictionary<string, int>)dicTotal["dicAppear"];
            stringbuilder.AppendLine("<table>");
            stringbuilder.AppendLine(string.Format("<caption>{0}</caption>", "Total Freq 1~6 of 0125"));
            stringbuilder.AppendLine("<thead>");
            stringbuilder.AppendLine("<tr>");
            foreach (var keypair in dicAppear)
            {
                if (dicCurrentDataNums.ContainsValue(int.Parse(keypair.Key)))
                {
                    stringbuilder.AppendLine(string.Format("<th class='{0}'>", dicCurrentDataNums.FirstOrDefault(x => x.Value == (int.Parse(keypair.Key))).Key));
                }
                else
                {
                    stringbuilder.AppendLine("<th>");
                }
                stringbuilder.AppendLine(keypair.Key);
                stringbuilder.AppendLine("</th>");
            }
            stringbuilder.AppendLine("</tr>");
            stringbuilder.AppendLine("</thead>");
            stringbuilder.AppendLine("<tbody>");
            stringbuilder.AppendLine("<tr>");
            foreach (var keypair in dicAppear)
            {
                stringbuilder.AppendLine("<td>");
                stringbuilder.AppendLine(keypair.Value.ToString());
                stringbuilder.AppendLine("</td>");
            }
            stringbuilder.AppendLine("</tr>");
            stringbuilder.AppendLine("</tbody>");
            stringbuilder.AppendLine("</table>");
            #endregion Set dicAppear
            stringbuilder.AppendLine("</details>");

            #region Each Compare

            foreach (KeyValuePair<string, object> Items in dicCpmpares)
            {
                if (Items.Key != "dicTotal")
                {
                    Dictionary<string, object> dicEachCompare = (Dictionary<string, object>)Items.Value;
                    DataSet dsFreq = (DataSet)dicEachCompare["dsFreq"];
                    if (dsFreq.Tables["DataAll"].Rows.Count >= 20)
                    {
                        stringbuilder.AppendLine("<details class='PercentOK' open>");
                    }
                    else
                    {
                        stringbuilder.AppendLine("<details class='PercentNon'>");
                    }

                    #region Set Title 
                    string strTitle = string.Format("{0}: ", DataSet00.LottoDescription);
                    StuGLSearch testdate = new StuGLSearch();
                    testdate = stuSearch00;
                    testdate.LngCurrentData = testdate.LngDataStart;
                    dicCurrentData = new Dictionary<string, string>();
                    dicCurrentData = new CGLSearch().GetCurrentData(testdate);
                    strTitle += dicCurrentData["lngDateSN"] + " => ";
                    testdate.LngCurrentData = testdate.LngDataEnd;
                    dicCurrentData = new CGLSearch().GetCurrentData(testdate);
                    strTitle += dicCurrentData["lngDateSN"];
                    if (Items.Key != "gen")
                    {
                        if (Items.Key.Length > 0)
                        {
                            string[] strCompare = Items.Key.Split('#');
                            strTitle += " 相同 {";
                            foreach (string strCompareOption in strCompare)
                            {
                                strTitle += " " + new CGLFunc().ConvertFieldNameID(strCompareOption, 1) + " ,";
                            }
                            strTitle = strTitle.TrimEnd(',') + " }";
                        }
                    }
                    strTitle += string.Format("({0:00}期)", dsFreq.Tables["DataAll"].Rows.Count);
                    stringbuilder.AppendLine(string.Format("<summary class='Title'>{0}</summary>", strTitle));
                    #endregion Set Title 

                    #region Frequency Part
                    foreach (DataTable dtFreq in dsFreq.Tables)
                    {
                        if (dtFreq.TableName != "DataAll" && dtFreq.TableName != "MissAll")
                        {
                            #region Set dtFreq

                            stringbuilder.AppendLine("<table>");
                            stringbuilder.AppendLine(string.Format("<caption>{0}</caption>", dtFreq.TableName));
                            stringbuilder.AppendLine("<thead>");
                            stringbuilder.AppendLine("<tr>");
                            foreach (DataColumn dcColumn in dtFreq.Columns)
                            {
                                int intNums = int.Parse(dcColumn.ColumnName.Substring(4).ToString());
                                if (dicCurrentDataNums.ContainsValue(intNums))
                                {
                                    stringbuilder.AppendLine(string.Format("<th class='{0}'>", dicCurrentDataNums.FirstOrDefault(x => x.Value == intNums).Key));
                                }
                                else
                                {
                                    stringbuilder.AppendLine("<th>");
                                }
                                stringbuilder.AppendLine(string.Format("{0:D2}", intNums));
                                stringbuilder.AppendLine("</th>");
                            }
                            stringbuilder.AppendLine("</tr>");
                            stringbuilder.AppendLine("</thead>");
                            stringbuilder.AppendLine("<tbody>");
                            stringbuilder.AppendLine("<tr>");
                            foreach (DataRow drRow in dtFreq.Rows)
                            {
                                for (int intcolumn = 0; intcolumn < dtFreq.Columns.Count; intcolumn++)
                                {
                                    stringbuilder.AppendLine("<td>");
                                    stringbuilder.AppendLine(drRow[intcolumn].ToString());
                                    stringbuilder.AppendLine("</td>");
                                }
                            }
                            stringbuilder.AppendLine("</tr>");
                            stringbuilder.AppendLine("</tbody>");
                            stringbuilder.AppendLine("</table>");
                            #endregion Set dtFreq
                        }
                    }
                    #endregion Frequency Part

                    #region Hot Part
                    DataSet dsHot = (DataSet)dicEachCompare["dsHot"];
                    stringbuilder.AppendLine("<table class='hot00'>");
                    stringbuilder.AppendLine("<tr>");
                    foreach (DataTable dtHot in dsHot.Tables)
                    {
                        DataView dvShow = dtHot.DefaultView;
                        dvShow.Sort = "[0115] DESC , [0105] DESC , [Nums] ASC ";
                        DataTable dtShow = dvShow.ToTable();
                        stringbuilder.AppendLine("<td class='hot'>");
                        stringbuilder.AppendLine("<table>");
                        stringbuilder.AppendLine("<thead>");
                        stringbuilder.AppendLine("<tr>");
                        foreach (DataColumn dcdgColumn in dtShow.Columns)
                        {
                            stringbuilder.AppendLine("<th>");
                            stringbuilder.AppendLine(string.Format("{0}", new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1)));
                            stringbuilder.AppendLine("</th>");
                        }
                        stringbuilder.AppendLine("</tr>");
                        stringbuilder.AppendLine("</thead>");
                        foreach (DataRow drRow in dtShow.Rows)
                        {
                            string strclass = "";
                            if (int.Parse(drRow[5].ToString()) >= 4)
                            {
                                strclass = "LightSeaGreen";
                            }
                            if (int.Parse(drRow[1].ToString()) >= 3)
                            {
                                strclass = "LightPink";
                            }
                            if (int.Parse(drRow[1].ToString()) == 0 && int.Parse(drRow[5].ToString()) >= 4)
                            {
                                strclass = "LightGreen";
                            }

                            stringbuilder.AppendLine(string.Format("<tr class='{0}'>", strclass));
                            for (int intColumn = 0; intColumn < dtShow.Columns.Count; intColumn++)
                            {
                                if (intColumn == 0)
                                {
                                    int intNums = int.Parse(drRow[intColumn].ToString());
                                    if (dicCurrentDataNums.ContainsValue(intNums))
                                    {
                                        stringbuilder.AppendLine(string.Format("<td class='{0}'>", dicCurrentDataNums.FirstOrDefault(x => x.Value == intNums).Key));
                                    }
                                    else
                                    {
                                        stringbuilder.AppendLine("<td>");
                                    }
                                    stringbuilder.AppendLine(string.Format("{0}", drRow[intColumn]));
                                    stringbuilder.AppendLine("</td>");
                                }
                                else
                                {
                                    stringbuilder.AppendLine("<td>");
                                    stringbuilder.AppendLine(string.Format("{0}", drRow[intColumn]));
                                    stringbuilder.AppendLine("</td>");
                                }
                            }
                            stringbuilder.AppendLine("</tr>");
                        }
                        stringbuilder.AppendLine("</table>");
                        stringbuilder.AppendLine("</td>");
                    }
                    stringbuilder.AppendLine("</tr>");
                    stringbuilder.AppendLine("</table>");
                    #endregion Hot Part

                    stringbuilder.AppendLine("</details>");
                }
            }
            #endregion Each Compare

            stringbuilder.AppendLine("</body>");
            stringbuilder.AppendLine("</html>");

            System.IO.File.WriteAllText(System.IO.Path.Combine(strHtmlDirectory, strFileName), stringbuilder.ToString(), Encoding.UTF8);

        }
        private void WPercent_KeyDown(object sender, KeyEventArgs e)
        {
            Window wPercent = (Window)sender;
            ScrollViewer svViewer = (ScrollViewer)wPercent.Content;
            StackPanel spmain = (StackPanel)svViewer.Content;
            if (e.Key == Key.P && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                PrintDialog pdPrint = new PrintDialog()
                {
                    CurrentPageEnabled = true,
                    SelectedPagesEnabled = true,
                    UserPageRangeEnabled = true
                };
                pdPrint.PrintTicket.PageOrientation = PageOrientation.Landscape;
                FixedDocument fdDocument = GetFixedDocument(spmain, pdPrint);
                ShowPrintPreview(fdDocument);
                //                if (pdPrint.ShowDialog() == true)
                //                {
                // pdPrint.PrintTicket.PageOrientation = PageOrientation.Landscape;
                // pdPrint.PrintTicket.PageScalingFactor = 120;
                // DocumentPaginator dpPrint = ((IDocumentPaginatorSource)expdgMissAll.Content).DocumentPaginator;
                // pdPrint.PrintDocument(dpPrint, "Print");
                //pdPrint.PrintVisual(dvViewer, "Print");
                //}
            }

        }
        #endregion btnTablePercent

        #region btnTablePercent01 百分比預測表01
        private void BtnTablePercent01_Click(object sender, RoutedEventArgs e)
        {
            #region 設定 stuSearchPercent
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSearchPercent01 = stuRibbonSearchOption;
            stuSearchPercent01.BoolFieldMode = false;
            stuSearchPercent01.IntMatchMin = 0;
            stuSearchPercent01.IntDataOffset = 0;
            stuSearchPercent01.IntSearchLimit = 0;
            stuSearchPercent01.IntSearchOffset = 0;
            stuSearchPercent01 = new CGLSearch().GetMethodSN(stuSearchPercent01);
            #endregion 設定 stuSearchPercent

            #region Setting bwTablePercent
            bwBackgroundWorker00 = new BackgroundWorker()
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };
            bwBackgroundWorker00.DoWork += BwTablePercent01_DoWork;
            bwBackgroundWorker00.ProgressChanged += BwTablePercent_ProgressChanged;
            bwBackgroundWorker00.RunWorkerCompleted += BwTablePercent01_RunWorkerCompleted;
            #endregion Setting bwTablePercent

            bwBackgroundWorker00.RunWorkerAsync(stuSearchPercent01);
            btnTablePercent01.Content = "計算中...";
            btnTablePercent01.IsEnabled = false;

        }

        private void BwTablePercent01_DoWork(object sender, DoWorkEventArgs e)
        {
            #region 設定
            StuGLSearch stuSearchPercent01 = (StuGLSearch)e.Argument;
            BackgroundWorker bwMyWork = (BackgroundWorker)sender;
            bwMyWork.ReportProgress(1);
            Dictionary<string, object> dicArgument = new Dictionary<string, object>
            {
                { "stuSearch", stuSearchPercent01 }
            };
            #endregion
            #region Check tblPercent hsa Data
            Dictionary<string, object> dicCompares01 = new Dictionary<string, object>();
            if (!new CGLSearch().HasPercentAllData(stuSearchPercent01))
            {
                new CGLSearch().SearchTablePercentAll(stuSearchPercent01);
            }

            new CGLSearch().GetTablePercentAll(stuSearchPercent01);

            #endregion
            dicArgument.Add("Compares", dicCompares01);
            e.Result = dicArgument;
        }

        private void BwTablePercent01_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Dictionary<string, object> dicArgument = (Dictionary<string, object>)e.Result;
            //ShowTablePercent((stuGLSearch)dicArgument["stuSearch"], (Dictionary<string, object>)dicArgument["Compares"]);
            tbStatusTextBlock.Text = "Ready";
            btnTablePercent01.Content = "百分比預測表01";
            btnTablePercent01.IsEnabled = true;
            BackgroundWorker bwMywork = (BackgroundWorker)sender;
            bwMywork.Dispose();
        }

        #endregion btnTablePercent01

        #region btnLastNum 末期資料
        private void BtnLastNum_Click(object sender, RoutedEventArgs e)
        {
            #region 設定 stuSearchLast
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSearchLast = stuRibbonSearchOption;
            #endregion 設定 stuSearchsmart

            #region Setting bwBackgroundWorker00
            bwBackgroundWorker00 = new BackgroundWorker()
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };
            bwBackgroundWorker00.DoWork += BwLastNum_DoWork;
            bwBackgroundWorker00.ProgressChanged += BwLastNum_ProgressChanged;
            bwBackgroundWorker00.RunWorkerCompleted += BwLastNum_RunWorkerCompleted;
            #endregion Setting bwTablePercent
            bwBackgroundWorker00.RunWorkerAsync(stuSearchLast);
            btnLastNum.Content = "計算中...";
            btnLastNum.IsEnabled = false;
        }
        private void BwLastNum_DoWork(object sender, DoWorkEventArgs e)
        {
            StuGLSearch stuSearchLast = (StuGLSearch)e.Argument;
            BackgroundWorker bwMyWork = (BackgroundWorker)sender;
            bwMyWork.ReportProgress(0);
            Dictionary<string, object> dicArgument = new Dictionary<string, object>
            {
                { "stuSearch", stuSearchLast }
            };
            DataSet dsDataSet00 = new CGLSearch().SearchLastNum(stuSearchLast);
            dicArgument.Add("Freq", dsDataSet00);
            e.Result = dicArgument;
        }
        private void BwLastNum_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 0)
            {
                tbStatusTextBlock.Text = "計算中...";
            }

        }
        private void BwLastNum_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Dictionary<string, object> dicArgument = (Dictionary<string, object>)e.Result;
            StuGLSearch stuSearchLast = (StuGLSearch)dicArgument["stuSearch"];
            DataSet dsDataSet00 = (DataSet)dicArgument["Freq"];
            ShowLastNum(stuSearchLast, dsDataSet00);


            tbStatusTextBlock.Text = "Ready";
            btnLastNum.Content = "末期資料";
            btnLastNum.IsEnabled = true;
            BackgroundWorker bwMywork = (BackgroundWorker)sender;
            bwMywork.Dispose();
        }
        private void DgLastSum_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            DataGridRow dgrRow = e.Row;
            DataRowView drvGridRow = (DataRowView)dgrRow.Item;
            if (int.Parse(drvGridRow.Row.ItemArray[1].ToString()) == 0 && int.Parse(drvGridRow.Row.ItemArray[2].ToString()) == 0)
            {
                e.Row.Background = Brushes.LightYellow;
            }
        }
        private void ShowLastNum(StuGLSearch stuSearch00, DataSet dsDataSet00)
        {
            CGLDataSet Dataset00 = new CGLDataSet(stuSearch00.LottoType);
            Dictionary<string, string> dicCurrentData = new CGLSearch().GetCurrentData(stuSearch00);
            #region Set lstCurreny
            List<int> lstCurrentData = new List<int>();
            for (int intCount = 1; intCount <= Dataset00.CountNumber; intCount++)
            {
                lstCurrentData.Add(int.Parse(dicCurrentData[string.Format("lngL{0}", intCount)].ToString()));
            }
            if (Dataset00.DataType != TableType.LottoWeli && Dataset00.SCountNumber > 0)
            {
                lstCurrentData.Add(int.Parse(dicCurrentData[string.Format("lngS1")].ToString()));
            }
            #endregion Set lstCurreny

            #region 顯示
            #region Set DataGrid dgCurrentData
            DataGrid dgCurrentData = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                RowDetailsVisibilityMode = DataGridRowDetailsVisibilityMode.Visible,
                CanUserAddRows = true,
                CanUserDeleteRows = false,
                CanUserResizeColumns = true,
                CanUserResizeRows = true,
                CanUserSortColumns = true,
                IsReadOnly = true,
                AutoGenerateColumns = false,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch
            };
            #region Convert dictionary dicCurrentData to table dtCurrentData
            DataTable dtCurrentData = new CGLFunc().CDicTOTable(dicCurrentData);
            #endregion

            dgCurrentData.ItemsSource = dtCurrentData.DefaultView;

            #region Set Columns of DataGrid dgCurrentData 
            DataGridTextColumn dgtColumn;
            if (dgCurrentData.Columns.Count == 0)
            {
                foreach (var KeyPair in dicCurrentData)
                {
                    dgtColumn = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(KeyPair.Key, 1),
                        Binding = new Binding(KeyPair.Key),
                        IsReadOnly = true
                    };
                    dgCurrentData.Columns.Add(dgtColumn);
                }
            }
            #endregion

            #endregion

            #region DataGrid dgLast
            DataTable dtShow = dsDataSet00.Tables["Show"];
            DataGrid dgLast = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                AutoGenerateColumns = false,
                ItemsSource = dtShow.DefaultView
            };

            #region Set Columns of DataGrid dgActive
            DataGridTextColumn dgtColumnLast;
            if (dgLast.Columns.Count == 0)
            {
                foreach (DataColumn dcdgColumn in dtShow.Columns)
                {
                    dgtColumnLast = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1),
                        Binding = new Binding(dcdgColumn.ColumnName),
                        IsReadOnly = true
                    };
                    if (dcdgColumn.ColumnName == "lngDateSN")
                    {
                        Style Style00 = new Style();
                        Setter setter00 = new Setter()
                        {
                            Property = BackgroundProperty,
                            Value = Brushes.LightBlue
                        };
                        Style00.Setters.Add(setter00);
                        dgtColumnLast.CellStyle = Style00;
                    }
                    if (dcdgColumn.ColumnName == "intSum")
                    {
                        Style Style00 = new Style();
                        Setter setter00 = new Setter()
                        {
                            Property = BackgroundProperty,
                            Value = Brushes.LightGreen
                        };
                        Style00.Setters.Add(setter00);
                        dgtColumnLast.CellStyle = Style00;
                    }
                    if (dcdgColumn.ColumnName.Substring(0, 4) == "lngL")
                    {
                        #region Set Trigger
                        Style Style00 = new Style()
                        {
                            TargetType = typeof(DataGridCell)
                        };
                        foreach (int intNum in lstCurrentData)
                        {
                            if (intNum > 0)
                            {
                                DataTrigger dtTrig00 = new DataTrigger()
                                {
                                    Binding = dgtColumnLast.Binding,
                                    Value = intNum.ToString()
                                };
                                Setter setter00 = new Setter()
                                {
                                    Property = ForegroundProperty,
                                    Value = Brushes.Yellow
                                };
                                dtTrig00.Setters.Add(setter00);
                                setter00 = new Setter()
                                {
                                    Property = BackgroundProperty,
                                    Value = Brushes.Red
                                };
                                dtTrig00.Setters.Add(setter00);
                                Style00.Triggers.Add(dtTrig00);
                            }
                        }
                        #endregion Set Trigger
                        dgtColumnLast.CellStyle = Style00;
                    }
                    dgLast.Columns.Add(dgtColumnLast);
                }
            }
            #endregion
            #endregion DataGrid

            #region DataGrid dgUpper
            DataTable dtUpper = dsDataSet00.Tables["Upper"];
            DataGrid dgUpper = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                AutoGenerateColumns = false,
                ItemsSource = dtUpper.DefaultView
            };

            #region Set Columns of DataGrid dgActive
            DataGridTextColumn dgtColumnUpper;
            if (dgUpper.Columns.Count == 0)
            {
                foreach (DataColumn dcdgColumn in dtUpper.Columns)
                {
                    dgtColumnUpper = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1),
                        Binding = new Binding(dcdgColumn.ColumnName),
                        IsReadOnly = true
                    };
                    if (dcdgColumn.ColumnName == "lngDateSN")
                    {
                        Style StyleUp = new Style();
                        Setter setter01 = new Setter()
                        {
                            Property = BackgroundProperty,
                            Value = Brushes.LightBlue
                        };
                        StyleUp.Setters.Add(setter01);
                        dgtColumnUpper.CellStyle = StyleUp;
                    }
                    if (dcdgColumn.ColumnName.Substring(0, 4) == "lngL")
                    {
                        Style Style00 = new Style();
                        #region Set Trigger
                        foreach (int intNum in lstCurrentData)
                        {
                            Style00.TargetType = typeof(DataGridCell);
                            DataTrigger dtTrig00 = new DataTrigger()
                            {
                                Binding = dgtColumnUpper.Binding,
                                Value = intNum
                            };
                            Setter setter00 = new Setter()
                            {
                                Property = ForegroundProperty,
                                Value = Brushes.Yellow
                            };
                            dtTrig00.Setters.Add(setter00);
                            setter00 = new Setter()
                            {
                                Property = BackgroundProperty,
                                Value = Brushes.Red
                            };
                            dtTrig00.Setters.Add(setter00);
                            Style00.Triggers.Add(dtTrig00);
                        }
                        #endregion Set Trigger
                        dgtColumnUpper.CellStyle = Style00;
                    }
                    dgUpper.Columns.Add(dgtColumnUpper);
                }
            }
            #endregion
            #endregion DataGrid

            #region DataGrid dgLower
            DataTable dtLower = dsDataSet00.Tables["Lower"];
            DataGrid dgLower = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                AutoGenerateColumns = false,
                ItemsSource = dtLower.DefaultView
            };


            #region Set Columns of DataGrid dgActive
            DataGridTextColumn dgtColumnLower;
            if (dgLower.Columns.Count == 0)
            {
                foreach (DataColumn dcdgColumn in dtLower.Columns)
                {
                    dgtColumnLower = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1),
                        Binding = new System.Windows.Data.Binding(dcdgColumn.ColumnName),
                        IsReadOnly = true
                    };
                    if (dcdgColumn.ColumnName == "lngDateSN")
                    {
                        Style StyleDown = new Style();
                        Setter setter00 = new Setter()
                        {
                            Property = BackgroundProperty,
                            Value = Brushes.LightBlue
                        };
                        StyleDown.Setters.Add(setter00);
                        dgtColumnLower.CellStyle = StyleDown;
                    }
                    if (dcdgColumn.ColumnName.Substring(0, 4) == "lngL")
                    {
                        Style Style00 = new Style();
                        #region Set Trigger
                        foreach (int intNum in lstCurrentData)
                        {
                            Style00.TargetType = typeof(DataGridCell);
                            DataTrigger dtTrig00 = new DataTrigger()
                            {
                                Binding = dgtColumnLower.Binding,
                                Value = intNum
                            };
                            Setter setter00 = new Setter()
                            {
                                Property = ForegroundProperty,
                                Value = Brushes.Yellow
                            };
                            dtTrig00.Setters.Add(setter00);
                            setter00 = new Setter()
                            {
                                Property = BackgroundProperty,
                                Value = Brushes.Red
                            };
                            dtTrig00.Setters.Add(setter00);
                            Style00.Triggers.Add(dtTrig00);
                        }
                        #endregion Set Trigger
                        dgtColumnLower.CellStyle = Style00;
                    }
                    dgLower.Columns.Add(dgtColumnLower);
                }
            }
            #endregion
            #endregion DataGrid

            #region DataGrid dgLastSum
            DataTable dtLastSum = dsDataSet00.Tables["LastSum"];
            DataGrid dgLastSum = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                CanUserResizeColumns = false,
                CanUserSortColumns = false,
                CanUserResizeRows = false,
                AutoGenerateColumns = false,
                ItemsSource = dtLastSum.DefaultView
            };


            #region Set Columns of DataGrid dgCurrentData 
            DataGridTextColumn dgtColumn00;
            if (dgLastSum.Columns.Count == 0)
            {
                int intColumnCount = 0;
                foreach (DataColumn dcColumn in dtLastSum.Columns)
                {
                    dgtColumn00 = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(dcColumn.ColumnName, 1),
                        Binding = new Binding(dcColumn.ColumnName),
                        IsReadOnly = true
                    };
                    if (intColumnCount == 0)
                    {
                        Style Style00 = new Style();
                        #region Set Trigger
                        foreach (int intNum in lstCurrentData)
                        {
                            Style00.TargetType = typeof(DataGridCell);
                            DataTrigger dtTrig00 = new DataTrigger()
                            {
                                Binding = new Binding(dcColumn.ColumnName),
                                Value = string.Format("{0:00}", intNum)
                            };
                            Setter setter00 = new Setter()
                            {
                                Property = ForegroundProperty,
                                Value = Brushes.Yellow
                            };
                            dtTrig00.Setters.Add(setter00);
                            setter00 = new Setter()
                            {
                                Property = BackgroundProperty,
                                Value = Brushes.Red
                            };
                            dtTrig00.Setters.Add(setter00);
                            Style00.Triggers.Add(dtTrig00);
                            setter00 = new Setter()
                            {
                                Property = BackgroundProperty,
                                Value = Brushes.LightBlue
                            };
                            Style00.Setters.Add(setter00);
                        }
                        #endregion Set Trigger
                        dgtColumn00.CellStyle = Style00;
                        dgtColumn00.Header = "號碼";
                    }
                    dgLastSum.Columns.Add(dgtColumn00);
                    intColumnCount++;
                }
            }
            #endregion
            dgLastSum.LoadingRow += DgLastSum_LoadingRow;
            dtLastSum.DefaultView.Sort = "[Upper] DESC , [Lower] DESC , [Nums] ASC";

            #endregion DataGrid

            #region DataGrid dgLast2
            DataTable dtLast2 = dsDataSet00.Tables["Last2"];
            DataGrid dgLast2 = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                CanUserResizeColumns = false,
                CanUserSortColumns = false,
                CanUserResizeRows = false,
                AutoGenerateColumns = false,
                ItemsSource = dtLast2.DefaultView
            };

            #region Set Columns of DataGrid dgCurrentData 
            DataGridTextColumn dgtColumn2;
            if (dgLast2.Columns.Count == 0)
            {
                int intColumnCount = 0;
                foreach (DataColumn dcColumn in dtLast2.Columns)
                {
                    dgtColumn2 = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(dcColumn.ColumnName, 1),
                        Binding = new Binding(dcColumn.ColumnName),
                        IsReadOnly = true
                    };
                    if (intColumnCount == 0)
                    {
                        Style Style00 = new Style();
                        #region Set Trigger
                        foreach (int intNum in lstCurrentData)
                        {
                            Style00.TargetType = typeof(DataGridCell);
                            DataTrigger dtTrig00 = new DataTrigger()
                            {
                                Binding = new Binding(dcColumn.ColumnName),
                                Value = string.Format("{0:00}", intNum)
                            };
                            Setter setter00 = new Setter()
                            {
                                Property = ForegroundProperty,
                                Value = Brushes.Yellow
                            };
                            dtTrig00.Setters.Add(setter00);
                            setter00 = new Setter()
                            {
                                Property = BackgroundProperty,
                                Value = Brushes.Red
                            };
                            dtTrig00.Setters.Add(setter00);
                            Style00.Triggers.Add(dtTrig00);
                            setter00 = new Setter()
                            {
                                Property = BackgroundProperty,
                                Value = Brushes.LightBlue
                            };
                            Style00.Setters.Add(setter00);
                        }
                        #endregion Set Trigger
                        dgtColumn2.CellStyle = Style00;
                        dgtColumn2.Header = "號碼";
                    }
                    dgLast2.Columns.Add(dgtColumn2);
                    intColumnCount++;
                }
            }
            #endregion
            dgLast2.LoadingRow += DgLastSum_LoadingRow;
            dtLast2.DefaultView.Sort = "[Upper] DESC , [Lower] DESC , [Nums] ASC";

            #endregion DataGrid

            #region StackPanel spDataGrid
            StackPanel spDataGrid = new StackPanel()
            {
                Orientation = Orientation.Horizontal
            };
            spDataGrid.Children.Add(dgLastSum);
            spDataGrid.Children.Add(dgLast2);
            spDataGrid.Children.Add(dgLast);
            spDataGrid.Children.Add(dgUpper);
            spDataGrid.Children.Add(dgLower);
            #endregion

            #region StackPanel spMain
            StackPanel spMain = new StackPanel()
            {
                CanVerticallyScroll = true,
                CanHorizontallyScroll = true,
                Orientation = Orientation.Vertical
            };
            spMain.Children.Add(dgCurrentData);
            spMain.Children.Add(spDataGrid);
            #endregion StackPanel

            #region ScrollViewer
            ScrollViewer svViewer = new ScrollViewer()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = spMain
            };
            #endregion ScrollViewer


            #region Window
            dtLastSum.DefaultView.RowFilter = "NOT ([Upper] = 0 AND [Lower] = 0) ";
            int intLastSumCount = dtLastSum.DefaultView.Count;
            dtLastSum.DefaultView.RowFilter = "";
            Window wActive = new Window()
            {
                Name = "tblLast",
                Title = string.Format("末期資料 ({0}) (共 {1} 個 {2}:{3} #{4})",
                                           dicCurrentData["lngDateSN"],
                                           intLastSumCount,
                                           dtUpper.DefaultView.Count,
                                           dtLower.DefaultView.Count,
                                           Dataset00.LottoNumbers - intLastSumCount),
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.SingleBorderWindow,
                Content = svViewer
            };
            wActive.Show();
            #endregion Window
            #endregion 顯示

        }
        #endregion btnLastNum

        #region btnDoubleZero 雙零資料
        private void BtnDoubleZero_Click(object sender, RoutedEventArgs e)
        {
            #region 設定 stuSearchDoubleZero
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSearchDoubleZero = stuRibbonSearchOption;
            #endregion 設定 stuSearchsmart

            #region Setting bwBackgroundWorker00
            bwBackgroundWorker00 = new BackgroundWorker()
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };
            bwBackgroundWorker00.DoWork += BwDoubleZero_DoWork;
            bwBackgroundWorker00.ProgressChanged += BwDoubleZero_ProgressChanged;
            bwBackgroundWorker00.RunWorkerCompleted += BwDoubleZero_RunWorkerCompleted;
            #endregion Setting bwTablePercent
            bwBackgroundWorker00.RunWorkerAsync(stuSearchDoubleZero);
            btnDoubleZero.Content = "計算中...";
            btnDoubleZero.IsEnabled = false;
        }
        private void BwDoubleZero_DoWork(object sender, DoWorkEventArgs e)
        {
            StuGLSearch stuSearchDoubleZero = (StuGLSearch)e.Argument;
            BackgroundWorker bwMyWork = (BackgroundWorker)sender;
            bwMyWork.ReportProgress(0);
            Dictionary<string, object> dicArgument = new Dictionary<string, object>
            {
                { "stuSearch", stuSearchDoubleZero }
            };
            DataSet dsDataSet00 = new CGLSearch().SearchDoubleZero(stuSearchDoubleZero);
            dicArgument.Add("Freq", dsDataSet00);
            e.Result = dicArgument;
        }
        private void BwDoubleZero_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 0)
            {
                tbStatusTextBlock.Text = "計算中...";
            }
        }
        private void BwDoubleZero_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Dictionary<string, object> dicArgument = (Dictionary<string, object>)e.Result;
            BackgroundWorker bwMywork = (BackgroundWorker)sender;
            StuGLSearch stuSearchDoubleZero = (StuGLSearch)dicArgument["stuSearch"];
            DataSet dsDataSet00 = (DataSet)dicArgument["Freq"];
            ShowDoubleZero(stuSearchDoubleZero, dsDataSet00);
            tbStatusTextBlock.Text = "Ready";
            btnDoubleZero.Content = "雙零資料";
            btnDoubleZero.IsEnabled = true;
            bwMywork.Dispose();
        }
        private void ShowDoubleZero(StuGLSearch stuSearch00, DataSet dsDataSet00)
        {
            #region Seting
            CGLDataSet Dataset00 = new CGLDataSet(stuSearch00.LottoType);
            Dictionary<string, string> dicCurrentData = new CGLSearch().GetCurrentData(stuSearch00);
            #region lstCurrentNum
            List<int> lstCurrentNum = new List<int>();
            for (int intCount = 1; intCount <= Dataset00.CountNumber; intCount++)
            {
                lstCurrentNum.Add(int.Parse(dicCurrentData[string.Format("lngL{0}", intCount)].ToString()));
            }
            if (Dataset00.DataType != TableType.LottoWeli && Dataset00.SCountNumber > 0)
            {
                lstCurrentNum.Add(int.Parse(dicCurrentData[string.Format("lngS1")].ToString()));
            }
            #endregion
            DataTable dtShow = dsDataSet00.Tables["Show"];
            DataTable dtDZCount = dsDataSet00.Tables["DZCount"];
            //DataTable dtResult = dsDataSet00.Tables["Result"];
            #endregion Seting

            #region 顯示

            #region dgShow
            DataGrid dgShow = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                AutoGenerateColumns = false,
                ItemsSource = dtShow.DefaultView
            };

            #region Set Columns of DataGrid dgActive
            DataGridTextColumn dgtColumnActive;
            if (dgShow.Columns.Count == 0)
            {
                foreach (DataColumn dcdgColumn in dtShow.Columns)
                {
                    dgtColumnActive = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1),
                        Binding = new Binding(dcdgColumn.ColumnName),
                        IsReadOnly = true
                    };
                    dgShow.Columns.Add(dgtColumnActive);
                }
            }
            #endregion
            #endregion dgShow

            #region dgDZCount
            DataGrid dgDZCount = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                CanUserResizeColumns = false,
                CanUserSortColumns = false,
                AutoGenerateColumns = false,
                ItemsSource = dtDZCount.DefaultView
            };

            #region Set Columns of DataGrid dgActive
            DataGridTextColumn dgtDZ;
            if (dgDZCount.Columns.Count == 0)
            {
                foreach (DataColumn dcdgColumn in dtDZCount.Columns)
                {
                    dgtDZ = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1),
                        Binding = new Binding(dcdgColumn.ColumnName)
                    };

                    if (dcdgColumn.ColumnName == "strNum")
                    {
                        #region Set Trigger
                        Style Style00 = new Style();
                        foreach (int intNum in lstCurrentNum)
                        {
                            Style00.TargetType = typeof(DataGridCell);
                            DataTrigger dtTrig00 = new DataTrigger()
                            {
                                Binding = dgtDZ.Binding,
                                Value = string.Format("{0:00}", intNum)
                            };
                            Setter setter00 = new Setter()
                            {
                                Property = ForegroundProperty,
                                Value = Brushes.Yellow
                            };
                            dtTrig00.Setters.Add(setter00);
                            setter00 = new Setter()
                            {
                                Property = BackgroundProperty,
                                Value = Brushes.Red
                            };
                            dtTrig00.Setters.Add(setter00);
                            Style00.Triggers.Add(dtTrig00);
                            setter00 = new Setter()
                            {
                                Property = BackgroundProperty,
                                Value = Brushes.LightBlue
                            };
                            Style00.Setters.Add(setter00);
                        }
                        #endregion Set Trigger
                        dgtDZ.CellStyle = Style00;
                    }
                    dgtDZ.IsReadOnly = true;
                    dgDZCount.Columns.Add(dgtDZ);
                }
            }
            #endregion
            dtDZCount.DefaultView.Sort = "[strNum] ASC ";
            #endregion dgDZCount

            #region dgCurrentData

            #region Set DataGrid dgCurrentData
            DataGrid dgCurrentData = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                RowDetailsVisibilityMode = DataGridRowDetailsVisibilityMode.Visible,
                CanUserAddRows = true,
                CanUserDeleteRows = false,
                CanUserResizeColumns = true,
                CanUserResizeRows = true,
                CanUserSortColumns = true,
                IsReadOnly = true,
                AutoGenerateColumns = false,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch
            };
            #endregion

            #region Convert dictionary dicCurrentData to table dtCurrentData
            DataTable dtCurrentData = new CGLFunc().CDicTOTable(dicCurrentData);
            #endregion

            dgCurrentData.ItemsSource = dtCurrentData.DefaultView;

            #region Set Columns of DataGrid dgCurrentData 
            DataGridTextColumn dgtColumn;
            if (dgCurrentData.Columns.Count == 0)
            {
                foreach (var KeyPair in dicCurrentData)
                {
                    dgtColumn = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(KeyPair.Key, 1),
                        Binding = new Binding(KeyPair.Key),
                        IsReadOnly = true
                    };
                    dgCurrentData.Columns.Add(dgtColumn);
                }
            }
            #endregion

            #endregion

            #region spDataGrids
            StackPanel spDataGrids = new StackPanel()
            {
                Orientation = Orientation.Horizontal,
                CanVerticallyScroll = true,
                CanHorizontallyScroll = true
            };
            //spDataGrids.Children.Add(dgDoubleZero);
            spDataGrids.Children.Add(dgShow);
            spDataGrids.Children.Add(dgDZCount);
            #endregion StackPanel

            #region spMain
            StackPanel spMain = new StackPanel()
            {
                Orientation = Orientation.Vertical,
                CanVerticallyScroll = true,
                CanHorizontallyScroll = true
            };
            spMain.Children.Add(dgCurrentData);
            spMain.Children.Add(spDataGrids);
            #endregion StackPanel

            #region ScrollViewer
            ScrollViewer svViewer = new ScrollViewer()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = spMain
            };
            #endregion ScrollViewer

            #region Window
            Window wActive = new Window()
            {
                Name = "tblDoubleZero",
                Title = string.Format("雙零資料 ({0}) ", dicCurrentData["lngDateSN"]),
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.SingleBorderWindow,
                Content = svViewer
            };
            wActive.Show();
            #endregion Window

            #region Change the color of DataGrid's row

            //foreach (KeyValuePair<string, int> drvGridRow in dgDoubleZero.ItemsSource)
            //{
            //    var row = (DataGridRow)dgDoubleZero.ItemContainerGenerator.ContainerFromItem(drvGridRow);
            //    if (lstCurrentNum.Contains(int.Parse(drvGridRow.Key)))
            //    {
            //        row.Background = Brushes.LightPink;
            //    }
            //}
            #endregion Change the color of DataGrid's row

            #endregion 顯示
        }

        #endregion btnDoubleZero

        #region btnMatch 週期同步效應
        private void BtnMatch_Click(object sender, RoutedEventArgs e)
        {
            #region 設定 stuSearchMatch
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSearchMatch = stuRibbonSearchOption;
            #endregion 設定 stuSearchsmart

            #region Setting bwBackgroundWorker00
            bwBackgroundWorker00 = new BackgroundWorker()
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };
            bwBackgroundWorker00.DoWork += BwMatch_DoWork;
            bwBackgroundWorker00.ProgressChanged += BwMatch_ProgressChanged;
            bwBackgroundWorker00.RunWorkerCompleted += BwMatch_RunWorkerCompleted;
            #endregion Setting bwTablePercent
            bwBackgroundWorker00.RunWorkerAsync(stuSearchMatch);
            btnMatch.Content = "計算中...";
            btnMatch.IsEnabled = false;
        }
        private void BwMatch_DoWork(object sender, DoWorkEventArgs e)
        {
            StuGLSearch stuSearchMatch = (StuGLSearch)e.Argument;
            BackgroundWorker bwMyWork = (BackgroundWorker)sender;
            bwMyWork.ReportProgress(0);
            Dictionary<string, object> dicArgument = new Dictionary<string, object>
            {
                { "stuSearch", stuSearchMatch }
            };
            DataSet dsDataSet00 = new CGLSearch().SearchMatch(stuSearchMatch);
            dicArgument.Add("Freq", dsDataSet00);
            e.Result = dicArgument;
        }
        private void BwMatch_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 0)
            {
                tbStatusTextBlock.Text = "計算中...";
            }
        }

        private void BwMatch_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Dictionary<string, object> dicArgument = (Dictionary<string, object>)e.Result;
            StuGLSearch stuSearchMatch = (StuGLSearch)dicArgument["stuSearch"];
            DataSet dsDataSet00 = (DataSet)dicArgument["Freq"];
            ShowMatch(stuSearchMatch, dsDataSet00);
            tbStatusTextBlock.Text = "Ready";
            btnMatch.Content = "同步效應";
            btnMatch.IsEnabled = true;
            BackgroundWorker bwMywork = (BackgroundWorker)sender;
            bwMywork.Dispose();
        }
        private void ShowMatch(StuGLSearch stuSearch00, DataSet dsDataSet00)
        {
            CGLDataSet Dataset00 = new CGLDataSet(stuSearch00.LottoType);
            Dictionary<string, string> dicCurrentData = new CGLSearch().GetCurrentData(stuSearch00);
            #region Set lstCurreny
            List<int> lstCurrentData = new List<int>();
            for (int intCount = 1; intCount <= Dataset00.CountNumber; intCount++)
            {
                lstCurrentData.Add(int.Parse(dicCurrentData[string.Format("lngL{0}", intCount)].ToString()));
            }
            if (Dataset00.DataType != TableType.LottoWeli && Dataset00.SCountNumber > 0)
            {
                lstCurrentData.Add(int.Parse(dicCurrentData[string.Format("lngS1")].ToString()));
            }
            #endregion Set lstCurreny

            #region 顯示
            #region Set DataGrid dgCurrentData
            DataGrid dgCurrentData = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                RowDetailsVisibilityMode = DataGridRowDetailsVisibilityMode.Visible,
                CanUserAddRows = true,
                CanUserDeleteRows = false,
                CanUserResizeColumns = true,
                CanUserResizeRows = true,
                CanUserSortColumns = true,
                IsReadOnly = true,
                AutoGenerateColumns = false,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch
            };
            #region Convert dictionary dicCurrentData to table dtCurrentData
            DataTable dtCurrentData = new CGLFunc().CDicTOTable(dicCurrentData);
            #endregion

            dgCurrentData.ItemsSource = dtCurrentData.DefaultView;

            #region Set Columns of DataGrid dgCurrentData 
            DataGridTextColumn dgtColumn;
            if (dgCurrentData.Columns.Count == 0)
            {
                foreach (var KeyPair in dicCurrentData)
                {
                    dgtColumn = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(KeyPair.Key, 1),
                        Binding = new Binding(KeyPair.Key),
                        IsReadOnly = true
                    };
                    dgCurrentData.Columns.Add(dgtColumn);
                }
            }
            #endregion

            #endregion

            #region DataGrid dgMatch
            DataTable dtMatch = dsDataSet00.Tables["dtMatch"];
            DataGrid dgMatch = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                CanUserResizeColumns = false,
                CanUserSortColumns = false,
                CanUserResizeRows = false,
                AutoGenerateColumns = false,
                ItemsSource = dtMatch.DefaultView
            };

            #region Set Columns of DataGrid dgCurrentData 
            DataGridTextColumn dgtColumn00;
            if (dgMatch.Columns.Count == 0)
            {
                int intColumnCount = 0;
                foreach (DataColumn dcColumn in dtMatch.Columns)
                {
                    dgtColumn00 = new DataGridTextColumn()
                    {
                        Binding = new Binding(dcColumn.ColumnName),
                        Header = dcColumn.Caption,
                        IsReadOnly = true
                    };
                    if (intColumnCount >= 3)
                    {
                        Style Style00 = new Style();
                        #region Set Trigger
                        foreach (int intNum in lstCurrentData)
                        {
                            Style00.TargetType = typeof(DataGridCell);
                            DataTrigger dtTrig00 = new DataTrigger()
                            {
                                Binding = new Binding(dcColumn.ColumnName),
                                Value = string.Format("{0:00}", intNum)
                            };
                            Setter setter00 = new Setter()
                            {
                                Property = ForegroundProperty,
                                Value = Brushes.Yellow
                            };
                            dtTrig00.Setters.Add(setter00);
                            setter00 = new Setter()
                            {
                                Property = BackgroundProperty,
                                Value = Brushes.Red
                            };
                            dtTrig00.Setters.Add(setter00);
                            Style00.Triggers.Add(dtTrig00);
                            setter00 = new Setter()
                            {
                                Property = BackgroundProperty,
                                Value = Brushes.LightBlue
                            };
                            Style00.Setters.Add(setter00);
                        }
                        #endregion Set Trigger
                        dgtColumn00.CellStyle = Style00;
                    }
                    dgMatch.Columns.Add(dgtColumn00);
                    intColumnCount++;
                }
            }
            #endregion
            dtMatch.DefaultView.Sort = "[intMatch] DESC , [lngDateSN] DESC";

            #endregion DataGrid

            #region DataGrid dgMatchCount
            DataTable dtMarchCount = dsDataSet00.Tables["dtMarchCount"];
            DataGrid dgMatchCount = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                CanUserResizeColumns = false,
                CanUserSortColumns = false,
                CanUserResizeRows = false,
                AutoGenerateColumns = false,
                ItemsSource = dtMarchCount.DefaultView
            };

            #region Set Columns of DataGrid dgCurrentData 
            DataGridTextColumn dgtColumnMatch;
            if (dgMatchCount.Columns.Count == 0)
            {
                int intColumnCount = 0;
                foreach (DataColumn dcColumn in dtMarchCount.Columns)
                {
                    dgtColumnMatch = new DataGridTextColumn()
                    {
                        Binding = new Binding(dcColumn.ColumnName),
                        Header = dcColumn.Caption,
                        IsReadOnly = true
                    };
                    if (intColumnCount == 0)
                    {
                        Style Style00 = new Style();
                        #region Set Trigger
                        foreach (int intNum in lstCurrentData)
                        {
                            Style00.TargetType = typeof(DataGridCell);
                            DataTrigger dtTrig00 = new DataTrigger()
                            {
                                Binding = new Binding(dcColumn.ColumnName),
                                Value = string.Format("{0:00}", intNum)
                            };
                            Setter setter00 = new Setter()
                            {
                                Property = ForegroundProperty,
                                Value = Brushes.Yellow
                            };
                            dtTrig00.Setters.Add(setter00);
                            setter00 = new Setter()
                            {
                                Property = BackgroundProperty,
                                Value = Brushes.Red
                            };
                            dtTrig00.Setters.Add(setter00);
                            Style00.Triggers.Add(dtTrig00);
                            setter00 = new Setter()
                            {
                                Property = BackgroundProperty,
                                Value = Brushes.LightBlue
                            };
                            Style00.Setters.Add(setter00);
                        }
                        #endregion Set Trigger
                        dgtColumnMatch.CellStyle = Style00;
                    }
                    dgMatchCount.Columns.Add(dgtColumnMatch);
                    intColumnCount++;
                }
            }
            #endregion
            dtMarchCount.DefaultView.Sort = "[Count] DESC , [Nums] ASC ";

            #endregion DataGrid

            #region StackPanel spDataGrid
            StackPanel spDataGrid = new StackPanel()
            {
                Orientation = Orientation.Horizontal
            };
            spDataGrid.Children.Add(dgMatch);
            spDataGrid.Children.Add(dgMatchCount);

            #endregion

            #region StackPanel spMain
            StackPanel spMain = new StackPanel()
            {
                CanVerticallyScroll = true,
                CanHorizontallyScroll = true,
                Orientation = Orientation.Vertical
            };
            spMain.Children.Add(dgCurrentData);
            spMain.Children.Add(spDataGrid);
            #endregion StackPanel

            #region ScrollViewer
            ScrollViewer svViewer = new ScrollViewer()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = spMain
            };
            #endregion ScrollViewer


            #region Window
            Window wActive = new Window()
            {
                Name = "tblLast",
                Title = string.Format("同步效應 ({0}) ", stuSearch00.StrTitle),
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.SingleBorderWindow,
                Content = svViewer
            };
            wActive.Show();
            #endregion Window
            #endregion 顯示
        }
        #endregion btnMatch

        #region 聰明組合
        private void BtnSmartRun_Click(object sender, RoutedEventArgs e) //計算組合 
        {
            if (tbTestNum.Text.Length == 0) return;
            #region 設定


            #region 設定 stuSearchsmart
            StuGLSearch stuSearchsmart = ButtonClick(stuRibbonSearchOption);
            stuSearchsmart.IntMatchMin = 0;
            //stuSearchsmart = new CGLSearch().InitSearch(stuSearchsmart);
            stuSearchsmart = new CGLSearch().GetMethodSN(stuSearchsmart);
            #endregion 設定 stuSearchsmart

            CGLDataSet Dataset00 = new CGLDataSet(stuSearchsmart.LottoType);
            int intSeparateNum = Dataset00.LottoNumbers / 2;

            #region 設定 dtCombination
            DataTable dtCombination = new DataTable();
            #region Set Column
            DataColumn dcColumn = new DataColumn()
            {
                ColumnName = "ComSN",
                Caption = "序號",
                DataType = typeof(int)
            };
            dtCombination.Columns.Add(dcColumn);

            for (int i = 1; i <= Dataset00.CountNumber; i++)
            {
                dcColumn = new DataColumn()
                {
                    ColumnName = string.Format("lngM{0}", i.ToString()),
                    Caption = string.Format("號{0}", i.ToString()),
                    DataType = typeof(int)
                };
                dtCombination.Columns.Add(dcColumn);
            }

            dcColumn = new DataColumn()
            {
                ColumnName = "intSum",
                Caption = "和數值",
                DataType = typeof(int)
            };
            dtCombination.Columns.Add(dcColumn);

            dcColumn = new DataColumn()
            {
                ColumnName = "boolSum",
                Caption = "[和數]",
                DataType = typeof(string),
                DefaultValue = " "
            };
            dtCombination.Columns.Add(dcColumn);

            dcColumn = new DataColumn()
            {
                ColumnName = "intOdd",
                Caption = "奇數",
                DataType = typeof(int)
            };
            dtCombination.Columns.Add(dcColumn);

            dcColumn = new DataColumn()
            {
                ColumnName = "intEven",
                Caption = "偶數",
                DataType = typeof(int)
            };
            dtCombination.Columns.Add(dcColumn);

            dcColumn = new DataColumn()
            {
                ColumnName = "boolOddEven",
                Caption = "[奇偶]",
                DataType = typeof(string),
                DefaultValue = " "
            };
            dtCombination.Columns.Add(dcColumn);

            dcColumn = new DataColumn()
            {
                ColumnName = "intHigh",
                Caption = "大數",
                DataType = typeof(int)
            };
            dtCombination.Columns.Add(dcColumn);

            dcColumn = new DataColumn()
            {
                ColumnName = "intLow",
                Caption = "小數",
                DataType = typeof(int)
            };
            dtCombination.Columns.Add(dcColumn);

            dcColumn = new DataColumn()
            {
                ColumnName = "boolHighLow",
                Caption = "[大小]",
                DataType = typeof(string),
                DefaultValue = " "
            };
            dtCombination.Columns.Add(dcColumn);

            dcColumn = new DataColumn()
            {
                ColumnName = "intHot",
                Caption = "熱門",
                DataType = typeof(int)
            };
            dtCombination.Columns.Add(dcColumn);

            dcColumn = new DataColumn()
            {
                ColumnName = "intCool",
                Caption = "冷門",
                DataType = typeof(int)
            };
            dtCombination.Columns.Add(dcColumn);

            dcColumn = new DataColumn()
            {
                ColumnName = "boolHotCool",
                Caption = "[熱冷]",
                DataType = typeof(string),
                DefaultValue = " "
            };
            dtCombination.Columns.Add(dcColumn);

            dcColumn = new DataColumn()
            {
                ColumnName = "intLSum",
                Caption = "遺漏和",
                DataType = typeof(int)
            };
            dtCombination.Columns.Add(dcColumn);

            dcColumn = new DataColumn()
            {
                ColumnName = "intLAvg",
                Caption = "遺漏平均",
                DataType = typeof(float)
            };
            dtCombination.Columns.Add(dcColumn);

            #endregion Set Column
            #endregion 設定 DataTable

            #endregion 設定

            #region 計算

            #region 導入 tblMissAll
            DataTable dtMissall = new DataTable();
            #region Query String
            StuGLData stuData00 = new StuGLData()
            {
                LottoType = stuSearchsmart.LottoType,
                DataBaseType = DatabaseType.Miss,
                StrSELECT = " * ",
                StrFROM = "[qryMissAll]",
                StrWHERE = string.Format("[{0}] = {1}  ;",
                                         "lngTotalSN", stuSearchsmart.LngCurrentData - 1),
                StrORDER = ""
            };
            #endregion Query String
            dtMissall = stuData00.GetSourceData(stuData00);

            #endregion 導入 tblMissAll

            #region 處理組合陣列
            string[] strTest = tbTestNum.Text.Split(',');

            string strCom;
            strCom = new CGLFunc().Combination(strTest, Dataset00.CountNumber).TrimEnd(';');
            string[] strComArray = strCom.Split(';');
            #endregion 處理組合陣列

            #region 開始計算
            int intCount = 1;
            foreach (string strNums00 in strComArray)
            {
                int intSum = 0, intOdd = 0, intEven = 0, intHigh = 0, intLow = 0, intHot = 0, intCool = 0, intLSum = 0;
                float flLAvg = 0;
                DataRow drRow = dtCombination.NewRow();
                drRow["ComSN"] = intCount;
                int intNum = 1;
                string[] strNumArray = strNums00.Split('#');
                foreach (string strNum in strNumArray)
                {
                    int intTestNum = int.Parse(strNum);
                    drRow[string.Format("lngM{0}", intNum)] = intTestNum;
                    intSum += intTestNum;
                    if (intTestNum % 2 == 1) { intOdd++; } else { intEven++; }
                    if (intTestNum > intSeparateNum) { intHigh++; } else { intLow++; }
                    int intLost = int.Parse(dtMissall.Rows[0][string.Format("lngM{0}", intTestNum.ToString())].ToString());
                    if (intLost < 10) { intHot++; } else { intCool++; }
                    intLSum += intLost;
                    intNum++;
                }
                drRow["intSum"] = intSum;
                drRow["intOdd"] = intOdd;
                drRow["intEven"] = intEven;
                drRow["intHigh"] = intHigh;
                drRow["intLow"] = intLow;
                drRow["intHot"] = intHot;
                drRow["intCool"] = intCool;
                drRow["intLSum"] = intLSum;
                flLAvg = float.Parse(intLSum.ToString()) / float.Parse(Dataset00.CountNumber.ToString());
                drRow["intLAvg"] = flLAvg;
                int inttbOdds = int.Parse(tbOdds.Text);
                int inttbEvens = Dataset00.CountNumber - inttbOdds;
                int inttbHigh = int.Parse(tbHigh.Text);
                int inttbLow = Dataset00.CountNumber - inttbHigh;
                int inttbHot = int.Parse(tbHot.Text);
                int inttbCool = Dataset00.CountNumber - inttbHot;
                if ((intSum >= int.Parse(tbSumMin.Text)) && (intSum <= int.Parse(tbSumMax.Text))) { drRow["boolSum"] = "☻"; }
                if ((intOdd == inttbOdds) && (intEven <= inttbEvens)) { drRow["boolOddEven"] = "☻"; }
                if ((intHigh == inttbHigh) && (intLow <= inttbLow)) { drRow["boolHighLow"] = "☻"; }
                if ((intHot == inttbHot) && (intCool <= inttbCool)) { drRow["boolHotCool"] = "☻"; }
                dtCombination.Rows.Add(drRow);
                intCount++;
            }
            #endregion 開始計算

            #endregion 計算

            #region 顯示

            #region DataGrid
            DataGrid dgCombination = new DataGrid()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Top,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                AutoGenerateColumns = false,
                ItemsSource = dtCombination.DefaultView
            };
            #region Set Columns of DataGrid dgMissAll
            DataGridTextColumn dgtColumnMissAll;
            if (dgCombination.Columns.Count == 0)
            {
                foreach (DataColumn dcdgColumn in dtCombination.Columns)
                {
                    dgtColumnMissAll = new DataGridTextColumn
                    {
                        Header = new CGLFunc().ConvertFieldNameID(dcdgColumn.Caption, 1),
                        Binding = new System.Windows.Data.Binding(dcdgColumn.ColumnName),
                        IsReadOnly = true
                    };
                    dgCombination.Columns.Add(dgtColumnMissAll);
                }
            }
            #endregion
            dtCombination.DefaultView.Sort = "[boolSum] DESC , [boolOddEven] DESC , [boolHighLow] DESC , [boolHotCool] DESC ";
            #endregion DataGrid

            #region Expander 
            Expander expCombination = new Expander()
            {
                Header = "100 期 奇數偶數資料",
                ExpandDirection = ExpandDirection.Down,
                IsExpanded = true,
                Background = Brushes.Aqua,
                MaxHeight = 500,
                Content = dgCombination
            };
            #endregion Expander 

            #region ScrollViewer
            ScrollViewer svViewer = new ScrollViewer()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                Content = expCombination
            };
            #endregion ScrollViewer

            #region Window
            Window wMissAll = new Window()
            {
                Name = "Combination",
                Title = "組合表",
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.SingleBorderWindow,
                Content = svViewer
            };
            wMissAll.Show();
            #endregion Window

            #endregion 顯示

        }
        private void BtnSmartTest_Click(object sender, RoutedEventArgs e)
        {
            if (tbTestNum.Text.Length == 0)
            {
                return;
            }
            #region 設定 stuGLSearch
            stuRibbonSearchOption = ButtonClick(stuRibbonSearchOption);
            StuGLSearch stuSmartTest = stuRibbonSearchOption;
            #endregion 設定 stuGLSearch
            List<string> lstSmartTest = new CGLFunc().GetSmartSet(stuSmartTest, tbTestNum.Text, ChoseAndHit.Chose4Hit3);



        }
        #endregion 聰明組合
    }
}

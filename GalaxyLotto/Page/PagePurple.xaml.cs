using GalaxyLotto.ClassLibrary;
using System;
using System.Collections.Generic;
using System.Data;
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
using System.Xml;

namespace GalaxyLotto
{
    /// <summary>
    /// CreatPurple.xaml 的互動邏輯
    /// </summary>
    public partial class CreatPurple : Page
    {
        private StuPurple stuOptionPurple;
        public CreatPurple()
        {
            InitializeComponent();
            InitializeDate();
        }

        private void InitializeDate()
        {
            DataSet daOptions = new DataSet("Options");
            daOptions = SetOptionsFromXML();
            SetCombobox(ref cmbWYear, daOptions.Tables["cmbWYear"]);
            SetCombobox(ref cmbWMonth, daOptions.Tables["cmbWMonth"]);
            SetCombobox(ref cmbWDay, daOptions.Tables["cmbWDay"]);
            SetCombobox(ref cmbWHour, daOptions.Tables["cmbWHour"]);
            stuOptionPurple.IntGender = Gender.Male_plus;
        }

        private DataSet SetOptionsFromXML()
        {
            DataSet dsOptions = new DataSet("Options");
            //string strXmlFileName;
            //strXmlFileName = string.Format("{0}\\{1}", Environment.CurrentDirectory, "OptionSearch.xml");
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.LoadXml(Properties.Resources.OptionPurple);
            //xmldoc.Load(@strXmlFileName);
            foreach (XmlNode node in xmldoc.DocumentElement)
            {
                if (node.HasChildNodes)
                {
                    string TableName = "";
                    DataColumn DC;
                    DataRow DR;
                    foreach (XmlNode child in node.ChildNodes)
                    {
                        if (child.Name == "name")
                        {
                            TableName = child.InnerText;
                            dsOptions.Tables.Add(TableName);
                            if (dsOptions.Tables[TableName].Columns.Count == 0)
                            {
                                DC = new DataColumn()
                                {
                                    Caption = "id",
                                    ColumnName = "id",
                                    DataType = typeof(string)
                                };
                                dsOptions.Tables[TableName].Columns.Add(DC);
                                DC = new DataColumn()
                                {
                                    Caption = "description",
                                    ColumnName = "description",
                                    DataType = typeof(string)
                                };
                                dsOptions.Tables[TableName].Columns.Add(DC);
                            }
                        }
                        else
                        {
                            if (child.HasChildNodes)
                            {
                                Dictionary<string, string> DicDR = new Dictionary<string, string>();
                                foreach (XmlNode child1 in child.ChildNodes)
                                {
                                    DicDR.Add(child1.Name, child1.InnerText);
                                    if (child1.Name == "description")
                                    {
                                        DR = dsOptions.Tables[TableName].NewRow();
                                        DR["id"] = DicDR["id"];
                                        DR["description"] = DicDR["description"];
                                        dsOptions.Tables[TableName].Rows.Add(DR);
                                        DicDR.Clear();
                                    }
                                }
                            }
                        }
                    }
                    // MessageBox.Show(xmlString);
                }
            }
            return dsOptions;
        }
        private void SetCombobox(ref System.Windows.Controls.ComboBox comboBox, DataTable dataTable)
        {
            if (comboBox.ItemsSource != null)
            {
                return;
            }
            comboBox.ItemsSource = dataTable.DefaultView;
            comboBox.SelectedValuePath = "id";
            comboBox.DisplayMemberPath = "description";
            comboBox.SelectedIndex = 0;
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            RadioButton rb = (RadioButton)sender;
            if (rb.Content != null)
            {
                switch (Convert.ToInt16(rb.Tag))
                {
                    case 1:
                        stuOptionPurple.IntGender = Gender.Male_plus;
                        break;
                    case 2:
                        stuOptionPurple.IntGender = Gender.Femal_minuse;
                        break;
                }
            }
        }

        private void CmbWYear_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox cmb = (ComboBox)sender;
            stuOptionPurple.StrWYear = cmb.SelectedValue.ToString();
        }

        private void CmbWMonth_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox cmb = (ComboBox)sender;
            stuOptionPurple.StrWMonth = cmb.SelectedValue.ToString();
        }

        private void CmbWDay_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox cmb = (ComboBox)sender;
            stuOptionPurple.StrWDay = cmb.SelectedValue.ToString();
        }

        private void CmbWHour_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox cmb = (ComboBox)sender;
            stuOptionPurple.StrWHour = cmb.SelectedValue.ToString();
        }

        private void BtnCreate_Click(object sender, RoutedEventArgs e)
        {
            stuOptionPurple.Init();

            if (sp00.Children.Count > 0) sp00.Children.Clear();
            StackPanel sp000;
            Label lblsp00;
            #region 顯示四柱 干支          
            foreach (var KeyValue in stuOptionPurple.Dic12LocationShow["00"])
            {
                sp000 = new StackPanel()
                {
                    Orientation = Orientation.Horizontal
                };
                lblsp00 = new Label()
                {
                    Name = KeyValue.Key,
                    Width = 380,
                    FontSize = 11,
                    Content = KeyValue.Value
                };
                sp000.Children.Add(lblsp00);
                sp00.Children.Add(sp000);
            }
            #endregion 顯示四柱 干支

            #region 顯示十二宮
            Dictionary<string, WrapPanel> dic12StackPanel = new Dictionary<string, WrapPanel>();
            if (wp01.Children.Count > 0) wp01.Children.Clear(); dic12StackPanel.Add("01", wp01);
            if (wp02.Children.Count > 0) wp02.Children.Clear(); dic12StackPanel.Add("02", wp02);
            if (wp03.Children.Count > 0) wp03.Children.Clear(); dic12StackPanel.Add("03", wp03);
            if (wp04.Children.Count > 0) wp04.Children.Clear(); dic12StackPanel.Add("04", wp04);
            if (wp05.Children.Count > 0) wp05.Children.Clear(); dic12StackPanel.Add("05", wp05);
            if (wp06.Children.Count > 0) wp06.Children.Clear(); dic12StackPanel.Add("06", wp06);
            if (wp07.Children.Count > 0) wp07.Children.Clear(); dic12StackPanel.Add("07", wp07);
            if (wp08.Children.Count > 0) wp08.Children.Clear(); dic12StackPanel.Add("08", wp08);
            if (wp09.Children.Count > 0) wp09.Children.Clear(); dic12StackPanel.Add("09", wp09);
            if (wp10.Children.Count > 0) wp10.Children.Clear(); dic12StackPanel.Add("10", wp10);
            if (wp11.Children.Count > 0) wp11.Children.Clear(); dic12StackPanel.Add("11", wp11);
            if (wp12.Children.Count > 0) wp12.Children.Clear(); dic12StackPanel.Add("12", wp12);
            for (int i = 1; i <= 12; i++)
            {
                string strLocation = string.Format("{0}", i.ToString("D2"));
                int intColor = 0;
                foreach (var KeyValue in stuOptionPurple.Dic12LocationShow[strLocation])
                {
                    intColor += 13;
                    Label lblLabel = new Label()
                    {
                        Name = KeyValue.Key,
                        FontSize = 11,
                        Width = 90
                    };
                    if (stuOptionPurple.DicStar.ContainsKey(KeyValue.Key))
                    {
                        lblLabel.Background = stuOptionPurple.DicStar[KeyValue.Key].BrushColor;
                    }
                    else
                    {
                        lblLabel.Background = Brushes.White;
                    }
                    lblLabel.Padding = new Thickness(0);
                    lblLabel.Content = KeyValue.Value;
                    dic12StackPanel[strLocation].Children.Add(lblLabel);
                }
            }
            #endregion 顯示十二宮           

        }
        
        private void CreatPurpleXAML_Loaded(object sender, RoutedEventArgs e)
        {
            foreach (Window wintest in System.Windows.Application.Current.Windows)
            {
                if (wintest.Name == "NotifyIcon")
                {
                    WinNotifyIcon WinNotifyIcon00 = (WinNotifyIcon)wintest;
                    WinNotifyIcon00.notifyIcon00.ContextMenu.MenuItems["notifyIconCreatePurple"].Checked = true;
                }
            }
        }

        private void ImgClose_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            Image image = (Image)sender;
            image.Source = new BitmapImage(new Uri("pack://siteoforigin:,,,/Resources/btnclose01.png"));
        }
        private void ImgClose_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            Image image = (Image)sender;
            image.Source = new BitmapImage(new Uri("pack://siteoforigin:,,,/Resources/btnclose00.png"));
        }
    }
}

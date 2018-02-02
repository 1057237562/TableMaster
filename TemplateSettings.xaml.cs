using MahApps.Metro.Controls;
using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
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
using static TableMaster.SheetChooser;

namespace TableMaster
{
    /// <summary>
    /// Interaction logic for TemplateSettings.xaml
    /// </summary>
    public partial class TemplateSettings : MetroWindow
    {
        public IWorkbook nworkbook;
        public System.Windows.Controls.DataGrid nDataGrid;
        public List<Pos> headerPos = new List<Pos>();
        public Pos outputPos;
        public Timer UITimer;
        public bool menu_state = false;
        public static String tmp_fp = Environment.CurrentDirectory + @"\config.xml";

        public TemplateSettings()
        {
            InitializeComponent();
            XMLOperator.InitializeContainer(tmp_fp, this);
        }

        private void Browser_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Filter = "Excel(*.xls)|*.xls|Excel(*.xlsx)|*.xlsx";
            if (dialog.ShowDialog() == false)
            {
                return;
            }

            Path.Text = dialog.FileName;
        }

        private void Path_TextChanged(object sender, TextChangedEventArgs events)
        {
            if (File.Exists(@Path.Text))
            {
                TabControl1.Items.Clear();
                // Reinitialize Variable
                headerPos = new List<Pos>();
                FileStream fs = new FileStream(Path.Text, FileMode.Open, FileAccess.Read);
                HSSFWorkbook book = new HSSFWorkbook(fs);
                nworkbook = book;
                for (int i = 0; i < book.NumberOfSheets; i++)
                {
                    MetroTabItem item = new MetroTabItem();
                    System.Windows.Controls.DataGrid data = new System.Windows.Controls.DataGrid();
                    data.SelectedCellsChanged += new SelectedCellsChangedEventHandler(DataGrid_SelectedCellsChanged);

                    data.SelectionUnit = DataGridSelectionUnit.Cell;

                    data.AutoGenerateColumns = false;

                    ISheet nsheet = book.GetSheetAt(i);

                    List<String[]> list = new List<String[]>();
                    /* DataTable sheetdata = new DataTable();*/
                    for (int r = nsheet.FirstRowNum; r < nsheet.LastRowNum; r++)
                    {
                        if (nsheet.GetRow(r) == null)
                        {
                            continue;
                        }
                        bool done = false;
                        for (int j = nsheet.GetRow(r).FirstCellNum; j < nsheet.GetRow(r).LastCellNum; j++)
                        {
                            if (nsheet.GetRow(r).GetCell(j).ToString() != "")
                            {
                                headerPos.Add(new Pos(j, r));
                                done = true;
                                break;
                            }
                        }
                        if (done)
                        {
                            break;
                        }
                    }

                    for (int j = nsheet.GetRow(headerPos[i].Row).FirstCellNum + headerPos[i].Column; j < nsheet.GetRow(headerPos[i].Row).LastCellNum; j++)
                    {
                        nsheet.GetRow(headerPos[i].Row).GetCell(j).SetCellType(CellType.String);
                        DataGridColumn column = new DataGridTextColumn
                        {
                            Header = nsheet.GetRow(headerPos[i].Row).GetCell(j).StringCellValue,
                            Binding = new System.Windows.Data.Binding($"[{j.ToString()}]")
                        };
                        data.Columns.Add(column);
                    }

                    for (int j = headerPos[i].Row + 1; j < nsheet.LastRowNum; j++)
                    {
                        List<String> row = new List<String>();
                        IRow nr = nsheet.GetRow(j);
                        if (nr == null)
                        {
                            continue;
                        }
                        for (int w = nr.FirstCellNum + headerPos[i].Column; w < nr.LastCellNum; w++)
                        {
                            ICell cell = nr.GetCell(w);
                            if (cell == null)
                            {
                                continue;
                            }
                            switch (cell.CellType)
                            {
                                case CellType.Formula:
                                    try
                                    {
                                        if (cell.NumericCellValue == 0)
                                        {
                                            HSSFFormulaEvaluator e = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                                            e.EvaluateInCell(cell);
                                            row.Add(cell.ToString());
                                        }
                                        else
                                        {
                                            row.Add(cell.NumericCellValue.ToString());
                                        }
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            row.Add(cell.ToString());
                                        }
                                        catch
                                        {
                                            row.Add("");
                                        }
                                    }
                                    break;
                                default:
                                    row.Add(cell.ToString());
                                    break;
                            }

                            /*String cellvalue = cell.StringCellValue;
                            if (cellvalue != "")
                            {
                                row.Add(cellvalue);
                            }*/
                            /*DataRow row = sheetdata.NewRow();
                            row[nsheet.GetRow(0).GetCell(w).StringCellValue] = nsheet.GetRow(j).GetCell(w).StringCellValue;
                            sheetdata.Rows.Add(row);*/
                        }
                        list.Add(row.ToArray());
                    }

                    data.ItemsSource = list;
                    item.Content = data;
                    item.Header = book.GetSheetName(i);
                    TabControl1.Items.Add(item);
                }

                TabControl1.SelectedIndex = 0;

                //Generating Settings
                ISheet sheet = nworkbook.GetSheetAt(0);
                //Finding Head

                Pos startingarea = new Pos(0, 0);
                for(int i = sheet.FirstRowNum; i < sheet.LastRowNum; i++)
                {
                    if(sheet.GetRow(i) == null)
                    {
                        continue;
                    }
                    bool done = false;
                    for(int j = sheet.GetRow(i).FirstCellNum; j < sheet.GetRow(i).LastCellNum; j++)
                    {
                        ICell cell = sheet.GetRow(i).GetCell(j);
                        if(cell != null && cell.ToString() != "")
                        {
                            startingarea.Column = j;
                            startingarea.Row = i;
                            done = true;
                            break;
                        }
                    }
                    if (done)
                    {
                        break;
                    }
                }

                Pos startPoint = new Pos(0,0);
                for (int i = startingarea.Row; i < sheet.LastRowNum; i++)
                {
                    if (sheet.GetRow(i) == null || sheet.GetRow(i-1) == null)
                    {
                        continue;
                    }
                    bool done = false;
                    for (int j = startingarea.Column; j < sheet.GetRow(i).LastCellNum; j++)
                    {
                        ICell cell = sheet.GetRow(i).GetCell(j);
                        ICell lcell = sheet.GetRow(i - 1).GetCell(j);
                        if(lcell == null)
                        {
                            break;
                        }
                        if (cell.ToString() == "" && lcell.ToString() != "")
                        {
                            startPoint.Column = j;
                            startPoint.Row = i - 1;
                            done = true;
                            break;
                        }
                    }
                    if (done)
                    {
                        break;
                    }
                }
                //Ending
                Pos endPoint = FindEnding(startPoint, sheet);
                for (int i = 0; i < endPoint.Column - startPoint.Column + 1; i++)
                {
                    /*String content = sheet.GetRow(startPoint.Row).GetCell(startPoint.Column + i).ToString();
                    String[] contents = content.Split("\r\n".ToCharArray());
                    if (contents.Length > 1)
                    {
                        foreach (String c in contents)
                        {
                            ListBoxItem item = new ListBoxItem();
                            item.Content = c;
                            ItemList.Items.Add(item);
                        }
                    }
                    else
                    {
                        ListBoxItem item = new ListBoxItem();
                        item.Content = content;
                        ItemList.Items.Add(item);
                    }*/
                    ListBoxItem item = new ListBoxItem();
                    item.Content = sheet.GetRow(startPoint.Row).GetCell(startPoint.Column + i).ToString();
                    ItemList.Items.Add(item);
                }

                Output.Text = startPoint.Column + "," + (startPoint.Row + 1);
                outputPos = Pos.FromString(Output.Text);
                outputPos.SheetID = 0;
                XMLOperator.AddIntoXML(tmp_fp, "Path", Path.Text);
            }
        }
        
        private void DataGrid_SelectedCellsChanged(object sender,SelectedCellsChangedEventArgs e)
        {
            nDataGrid = (System.Windows.Controls.DataGrid)sender;
        }

        private void More_Click(object sender, RoutedEventArgs e)
        {
            if (UITimer == null)
            {
                UITimer = new Timer();
                UITimer.Tick += new EventHandler(UITimer_Tick);
                UITimer.Interval = 10;
                UITimer.Enabled = true;
                UITimer.Start();
                menu_state = !menu_state;
            }
        }

        private void UITimer_Tick(object sender,EventArgs e)
        {
            if (!menu_state)
            {
                Height -= (Height - 415) / 8;
                if(Height - 415 <= 1)
                {
                    Height = 415;
                    UITimer.Stop();
                    UITimer.Enabled = false;
                    UITimer = null;
                    More.Content = "More ▽";
                }
            }
            else
            {
                Height += (640 - Height) / 8;
                if (640 - Height <= 1)
                {
                    Height = 640;
                    UITimer.Stop();
                    UITimer.Enabled = false;
                    UITimer = null;
                    More.Content = "More △";
                }
            }
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            ListBoxItem item = new ListBoxItem();
            item.Content = Item.Text;
            ItemList.Items.Add(item);
            Item.Text = "";
        }

        private void Remove_Click(object sender, RoutedEventArgs e)
        {
            ItemList.Items.Remove(ItemList.SelectedItem);
        }

        private void Up_Click(object sender, RoutedEventArgs e)
        {
            int index = ItemList.Items.IndexOf(ItemList.SelectedItem);
            if (index > 0)
            {
                ListBoxItem item = (ListBoxItem)ItemList.SelectedItem;
                ItemList.Items.Remove(ItemList.SelectedItem);
                ItemList.Items.Insert(index - 1, item);
            }
        }

        private void Down_Click(object sender, RoutedEventArgs e)
        {
            int index = ItemList.Items.IndexOf(ItemList.SelectedItem);
            if (index > 0)
            {
                ListBoxItem item = (ListBoxItem)ItemList.SelectedItem;
                ItemList.Items.Remove(ItemList.SelectedItem);
                ItemList.Items.Insert(index + 1, item);
            }
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            var _cells = nDataGrid.SelectedCells;
            String pos = "";
            if (_cells.Any())
            {
                int rowIndex = nDataGrid.Items.IndexOf(_cells.First().Item) + headerPos[TabControl1.SelectedIndex].Row;
                int columnIndex = _cells.First().Column.DisplayIndex + headerPos[TabControl1.SelectedIndex].Column;
                pos = columnIndex + "," + (rowIndex + 1);
                // Because of title,the row Index must be added by one
            }
            else
            {
                return;
            }
            ISheet sheet = nworkbook.GetSheetAt(TabControl1.SelectedIndex);
            Pos startPoint = Pos.FromString(pos);
            Pos endPoint = FindEnding(startPoint, sheet);
            for(int i = 0; i<endPoint.Column - startPoint.Column + 1; i++)
            {
                ListBoxItem item = new ListBoxItem();
                item.Content = sheet.GetRow(startPoint.Row).GetCell(startPoint.Column + i).ToString();
                ItemList.Items.Add(item);
            }
            
            Output.Text = startPoint.Column + "," + (startPoint.Row + 1); // TODO Auto Generation
            outputPos = Pos.FromString(Output.Text);
            outputPos.SheetID = TabControl1.SelectedIndex;
        }

        private Pos FindEnding(Pos startPoint, ISheet sheet)
        {
            Pos endPoint = new Pos(startPoint.Column, startPoint.Row);
            bool stop = false;
            for (int j = startPoint.Column; j < sheet.GetRow(startPoint.Row).LastCellNum; j++)
            {
                ICell cell = sheet.GetRow(startPoint.Row).GetCell(j);
                if (cell.ToString() == "" || cell == null)
                {
                    endPoint.Column = j;
                    stop = true;
                    break;
                }
            }
            if (!stop)
            {
                endPoint.Column = sheet.GetRow(startPoint.Row).LastCellNum - 1;
            }
            return endPoint;
        }

        private void Yes_Click(object sender, RoutedEventArgs e)
        {
            List<String> il = new List<String>();
            foreach(ListBoxItem i in ItemList.Items)
            {
                il.Add(i.Content.ToString());
            }
            MainWindow.M.itemlist = il;
            MainWindow.M.outputPos = outputPos;
            MainWindow.M.modelPath = Path.Text;
            Close();
        }
    }
}

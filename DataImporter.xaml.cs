using MahApps.Metro.Controls;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using static TableMaster.SheetChooser;

namespace TableMaster
{
    /// <summary>
    /// Interaction logic for DataImporter.xaml
    /// </summary>
    public partial class DataImporter : MetroWindow
    {
        public IWorkbook nworkbook;
        public DataGrid nDataGrid;
        public List<Pos> headerPos = new List<Pos>();
        public bool menu_state = false;
        public TextBox SelectedText;

        public DataImporter()
        {
            InitializeComponent();
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
                //Reinitialize Variable
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
            }
        }

        private void DataGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            DataGrid data = (DataGrid)sender;
            var _cells = data.SelectedCells;
            if (_cells.Any())
            {
                if (SelectedText != null)
                {
                    int rowIndex = data.Items.IndexOf(_cells.First().Item) + headerPos[TabControl1.SelectedIndex].Row;
                    int columnIndex = _cells.First().Column.DisplayIndex + headerPos[TabControl1.SelectedIndex].Column;
                    SelectedText.Text = columnIndex + "," + (rowIndex + 1);
                    // Because of title,the row Index must be added by one
                }
            }

            if (SelectedText == From)
            {
                ISheet sheet = nworkbook.GetSheetAt(TabControl1.SelectedIndex);
                Pos startPoint = Pos.FromString(From.Text);
                Pos endPoint = FindEnding(startPoint, sheet);
                To.Text = endPoint.toString();
            }
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            Pos startPos = Pos.FromString(From.Text);
            Pos endPos = Pos.FromString(To.Text);
            ISheet sheet = nworkbook.GetSheetAt(TabControl1.SelectedIndex);
            for (int i = startPos.Row; i <= endPos.Row; i++)
            {
                for(int j = startPos.Column; j<=endPos.Column; j++)
                {
                    ListBoxItem item = new ListBoxItem();
                    item.Content = sheet.GetRow(i).GetCell(j).ToString();
                    MainWindow.M.NameList.Items.Add(item);
                }
            }
            XMLOperator.AddIntoXML(MainWindow.ini_fp, MainWindow.M.NameList);
            Close();
        }

        private Pos FindEnding(Pos startPoint, ISheet sheet)
        {
            Pos endPoint = new Pos(startPoint.Column, startPoint.Row);
            bool stop = false;
            for (int i = startPoint.Row; i < sheet.LastRowNum; i++)
            {
                if(sheet.GetRow(i) == null)
                {
                    continue;
                }
                ICell cell = sheet.GetRow(i).GetCell(startPoint.Column);
                if (cell.ToString() == "" || cell == null)
                {
                    endPoint.Row = i - 1;
                    stop = true;
                    break;
                }
            }
            if (!stop)
            {
                endPoint.Row = sheet.LastRowNum - 1;
            }
            return endPoint;
        }

        private void From_GotFocus(object sender, RoutedEventArgs e)
        {
            SelectedText = (TextBox)sender;
        }

        private void To_GotFocus(object sender, RoutedEventArgs e)
        {
            SelectedText = (TextBox)sender;
        }
    }
}

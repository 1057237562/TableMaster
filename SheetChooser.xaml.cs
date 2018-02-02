using MahApps.Metro.Controls;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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

namespace TableMaster
{

    /// <summary>
    /// Interaction logic for SheetChooser.xaml
    /// </summary>
    public partial class SheetChooser : MetroWindow
    {
        public TextBox SelectedTextBox = null;
        public IWorkbook nworkbook = null;
        public List<Pos> headerPos = new List<Pos>();

        public struct Pos
        {
            public int Column;
            public int Row;
            public int SheetID;
            public Pos(int c,int r)
            {
                Column = c;
                Row = r;
                SheetID = 0;
            }

            public String toString()
            {
                return Column + "," + Row;
            }

            public static Pos FromString(String str)
            {
                try
                {
                    return new Pos(Convert.ToInt16(str.Split(",".ToCharArray())[0].Trim()), Convert.ToInt16(str.Split(",".ToCharArray())[1].Trim()));
                }
                catch {
                    try
                    {
                        System.Console.WriteLine(str.Split(",".ToCharArray())[0] + "|" + str.Split(",".ToCharArray())[1]);
                    }
                    catch {
                        System.Console.WriteLine(str);
                    }
                    return new Pos(0, 0);
                }
            }
        }

        public SheetChooser()
        {
            InitializeComponent();
            
            FileStream fs = new FileStream(MainWindow.addfile, FileMode.Open, FileAccess.Read);
            HSSFWorkbook book = new HSSFWorkbook(fs);
            nworkbook = book;
            for (int i = 0; i<book.NumberOfSheets; i++)
            {
                MetroTabItem item = new MetroTabItem();
                DataGrid data = new DataGrid();
                data.SelectedCellsChanged += new SelectedCellsChangedEventHandler(DataGrid_SelectedCellsChanged);

                data.SelectionUnit = DataGridSelectionUnit.Cell;

                data.AutoGenerateColumns = false;

                ISheet nsheet = book.GetSheetAt(i);

                List<String[]> list = new List<String[]>();
                /* DataTable sheetdata = new DataTable();*/

                //Find Start
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

                for (int j = nsheet.GetRow(headerPos[i].Row).FirstCellNum + headerPos[i].Column; j<nsheet.GetRow(headerPos[i].Row).LastCellNum; j++)
                {
                    nsheet.GetRow(headerPos[i].Row).GetCell(j).SetCellType(CellType.String);
                    DataGridColumn column = new DataGridTextColumn
                    {
                        Header = nsheet.GetRow(headerPos[i].Row).GetCell(j).StringCellValue,
                        Binding = new Binding($"[{j.ToString()}]")
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
                                    if(cell.NumericCellValue == 0)
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
        }

        private void TabControl1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SheetId.Text = TabControl1.SelectedIndex.ToString();

            //Generate Range
            ISheet sheet = nworkbook.GetSheetAt(TabControl1.SelectedIndex);
            Pos startPoint = new Pos(0,0);
            for(int i = headerPos[TabControl1.SelectedIndex].Row + 1; i<sheet.LastRowNum; i++)
            {
                if(sheet.GetRow(i) == null)
                {
                    continue;
                }
                bool done = false;
                for(int j = sheet.GetRow(i).FirstCellNum; j<sheet.GetRow(i).LastCellNum; j++)
                {
                    if(sheet.GetRow(i).GetCell(j) != null)
                    {
                        startPoint = new Pos(j,i);
                        done = true;
                        break;
                    }
                }
                if (done)
                {
                    break;
                }
            }

            Pos endPoint = FindEnding(startPoint, sheet);

            From.Text = startPoint.toString();
            To.Text = endPoint.toString();
        }

        private void DataGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            DataGrid data = (DataGrid)sender;
            var _cells = data.SelectedCells;
            if (_cells.Any())
            {
                if (SelectedTextBox != null) {
                    int rowIndex = data.Items.IndexOf(_cells.First().Item) + headerPos[TabControl1.SelectedIndex].Row;
                    int columnIndex = _cells.First().Column.DisplayIndex + headerPos[TabControl1.SelectedIndex].Column;
                    SelectedTextBox.Text = columnIndex + "," + (rowIndex + 1);
                    // Because of title,the row Index must be added by one
                }
            }

            if(SelectedTextBox == From)
            {
                ISheet sheet = nworkbook.GetSheetAt(TabControl1.SelectedIndex);
                Pos startPoint = Pos.FromString(From.Text);
                Pos endPoint = FindEnding(startPoint, sheet);
                To.Text = endPoint.toString();
            }
        }

        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            SelectedTextBox = (TextBox)sender;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            AdvancedFlyout.IsOpen = true;
        }

        private Pos FindEnding(Pos startPoint,ISheet sheet)
        {
            Pos endPoint = new Pos(0, 0);
            for (int i = (startPoint.Row + 1); i < sheet.LastRowNum; i++)
            {
                if(sheet.GetRow(i) == null)
                {
                    continue;
                }
                bool done = false;
                for (int j = startPoint.Column; j < sheet.GetRow(i).LastCellNum; j++)
                {
                    ICell cell = sheet.GetRow(i).GetCell(j);
                    if(cell == null)
                    {
                        continue;
                    }
                    if (cell.ToString() == "")
                    {
                        if (j == startPoint.Column)
                        {
                            endPoint.Row = i - 1;
                            done = true;
                            break;
                        }
                        bool detected = false;
                        for (int h = 0; h < Convert.ToInt16(Detect.Text); h++)
                        {
                            ICell icell = sheet.GetRow(i).GetCell(j + h);
                            if(icell == null)
                            {
                                break;
                            }
                            if (icell.ToString() != "")
                            {
                                detected = true;
                            }
                        }
                        if (detected)
                        {
                            continue;
                        }
                        endPoint.Column = j;
                        break;
                    }
                }
                if (done)
                {
                    break;
                }
            }
            return endPoint;
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            ListBoxItem item = new ListBoxItem();
            item.Content = MainWindow.addfile + @"|" + SheetId.Text + @"|From" + From.Text + "To" + To.Text;
            MainWindow.M.FileList.Items.Add(item);
            MainWindow.M.FileName.Text = "";
            XMLOperator.AddIntoXML(MainWindow.ini_fp,MainWindow.M.FileList);
            Close();
        }
    }
}

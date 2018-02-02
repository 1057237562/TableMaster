using MahApps.Metro.Controls;
using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using static TableMaster.SheetChooser;

namespace TableMaster
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>


    public partial class MainWindow : MetroWindow
    {
        public static String addfile = "";
        public static MainWindow M;

        public List<String> itemlist;
        public Pos outputPos;
        public String modelPath;

        public static String ini_fp = Environment.CurrentDirectory + @"\ini.xml";

        public MainWindow()
        {
            InitializeComponent();
            M = this;
            XMLOperator.InitializeContainer(ini_fp,this);
            Dispatcher.UnhandledException += new System.Windows.Threading.DispatcherUnhandledExceptionEventHandler(ExceptionHanlder);
        }

        private void ExceptionHanlder(object sender,System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            System.Windows.MessageBox.Show("An error has occurred. Check if there are some blank that isn't be filled." + Environment.NewLine + "Here are the Exception descriptions : " + e.Exception.Message);
            e.Handled = true;
        }

        private void Help_Click(object sender, RoutedEventArgs e)
        {
            HelpFlyout.IsOpen = true;
            AdvancedFlyout.IsOpen = false;
        }

        private void Advance_Click(object sender, RoutedEventArgs e)
        {
            HelpFlyout.IsOpen = false;
            AdvancedFlyout.IsOpen = true;
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            foreach (ListBoxItem name in NameList.Items)
            {
                FileStream fs = new FileStream(modelPath, FileMode.Open, FileAccess.Read);
                HSSFWorkbook book = new HSSFWorkbook(fs);
                ISheet msheet = book.GetSheetAt(outputPos.SheetID);
                for (int s = 0; s < FileList.Items.Count; s++){
                    ListBoxItem source = (ListBoxItem)FileList.Items[s];

                    String[] i_arg = ((String)source.Content).Split(@"|".ToCharArray()); // This arg is in the format like FilePath:SheetID:From...To...
                    String[] i_range = Regex.Split(i_arg[2].Replace("From", "").Trim(),"To");
                    //System.Console.WriteLine(i_arg[2].Replace("From", "") + ":" + i_range[0] + "To" + i_range[1]);
                    FileStream i_fs = new FileStream(i_arg[0], FileMode.Open, FileAccess.Read);
                    HSSFWorkbook i_book = new HSSFWorkbook(i_fs);
                    ISheet nsheet = i_book.GetSheetAt(Convert.ToInt16(i_arg[1]));
                    Pos startPos = Pos.FromString(i_range[0]);
                    Pos endPos = Pos.FromString(i_range[1]);

                    List<String> d_list = new List<String>();
                    if(nsheet.GetRow(startPos.Row - 1) == null)
                    {
                        continue;
                    }
                    for (int d = startPos.Column; d<=endPos.Column; d++)
                    {
                        ICell cell = nsheet.GetRow(startPos.Row - 1).GetCell(d);
                        //System.Console.WriteLine(startPos.Row - 1+","+d);
                        if (cell == null)
                        {
                            continue;
                        }
                        d_list.Add(cell.ToString().Trim());
                    }

                    int key_cell = FuzzyIndexOf(d_list,KeyWord.Text.Trim());
                    int fin_row = startPos.Row;
                    for(int n_index = startPos.Row; n_index<=endPos.Row; n_index++)
                    {
                        ICell cell = nsheet.GetRow(n_index).GetCell(key_cell);
                        if(cell == null)
                        {
                            continue;
                        }
                        if (cell.ToString().Trim() == (String)name.Content)
                        {
                            fin_row = n_index;
                            break;
                        }
                    }

                    for (int i = 0; i<itemlist.Count; i++)
                    {
                        String item = itemlist[i];
                        ICell output = msheet.GetRow(outputPos.Row + s).GetCell(outputPos.Column + i);
                        if(output == null)
                        {
                            continue;
                        }
                        //System.Console.WriteLine(fin_row + "," + d_list.IndexOf(item.Trim()));
                        String value = GetPracticalCellData(nsheet.GetRow(fin_row).GetCell(FuzzyIndexOf(d_list, item.Trim())));
                        if(value.Replace(@"[^\d]*","") == value && value.Replace(" ","").Replace("  ","") != "")
                        {
                            try
                            {
                                output.SetCellValue(Convert.ToDouble(value.Trim()));
                            }
                            catch {
                                output.SetCellValue(value);
                            }
                        }
                        else
                        {
                            output.SetCellValue(value);
                        }
                    }
                }
                //Output the Workbook
                FileStream outs = new FileStream(Path.Text + @"\"+name.Content + @".xls", FileMode.Create);
                outs.Flush();
                book.Write(outs);
                outs.Close();
                fs.Close();
            }
        }

        private String GetPracticalCellData(ICell cell)
        {
            if(cell == null)
            {
                System.Console.WriteLine(true);
                return "";
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
                            return cell.ToString();
                        }
                        else
                        {
                            return cell.NumericCellValue.ToString();
                        }
                    }
                    catch
                    {
                        try
                        {
                            return cell.ToString();
                        }
                        catch
                        {
                            return "";
                        }
                    }
                default:
                    return cell.ToString();
            }
        }

        private void Select_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog sflg = new Microsoft.Win32.OpenFileDialog();
            sflg.Filter = "Excel(*.xls)|*.xls|Excel(*.xlsx)|*.xlsx";
            if (sflg.ShowDialog() == false)
            {
                return;
            }

            FileName.Text = sflg.FileName;
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            addfile = FileName.Text;

            SheetChooser tb = new SheetChooser();
            tb.ShowDialog();
        }

        private void Remove_Click(object sender, RoutedEventArgs e)
        {
            FileList.Items.Remove(FileList.SelectedItem);
        }

        private void Up_Click(object sender, RoutedEventArgs e)
        {
            int index = FileList.Items.IndexOf(FileList.SelectedItem);
            if (index > 0)
            {
                ListBoxItem item = (ListBoxItem)FileList.SelectedItem;
                FileList.Items.Remove(FileList.SelectedItem);
                FileList.Items.Insert(index - 1, item);
            }
        }

        private void Down_Click(object sender, RoutedEventArgs e)
        {
            int index = FileList.Items.IndexOf(FileList.SelectedItem);
            if (index > 0)
            {
                ListBoxItem item = (ListBoxItem)FileList.SelectedItem;
                FileList.Items.Remove(FileList.SelectedItem);
                FileList.Items.Insert(index + 1, item);
            }
        }

        private void TemplateSetting_Click(object sender, RoutedEventArgs e)
        {
            TemplateSettings tp = new TemplateSettings();
            tp.ShowDialog();
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            DataImporter dp = new DataImporter();
            dp.ShowDialog();
        }

        private void Add1_Click(object sender, RoutedEventArgs e)
        {
            ListBoxItem item = new ListBoxItem();
            item.Content = Name.Text;
            NameList.Items.Add(item);
            Name.Text = "";
            XMLOperator.AddIntoXML(ini_fp, NameList);
        }

        private void Remove1_Click(object sender, RoutedEventArgs e)
        {
            //if (NameList.SelectedIndex == -1) { return; }
            //This two takes different effects.
            NameList.Items.Remove(NameList.SelectedItem);
            XMLOperator.AddIntoXML(ini_fp, NameList);
        }

        private void BroswerPath_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog f_dialog = new FolderBrowserDialog();
            if(f_dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Path.Text = f_dialog.SelectedPath;
            }
        }

        private int FuzzyIndexOf(List<String> list,String str)
        {
            foreach(String s in list)
            {
                if(s.Replace(str,"") != s)
                {
                    return list.IndexOf(s);
                }
            }
            return -1;
        }

        private void Path_TextChanged(object sender, TextChangedEventArgs e)
        {
            XMLOperator.AddIntoXML(ini_fp,"Path",Path.Text);
        }
    }
}

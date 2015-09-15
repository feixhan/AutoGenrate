using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CppRipper;
using System.IO;
using System.Data.OleDb;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;

namespace AutoGenrate
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void openBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog sourceFileDia = new OpenFileDialog();
            //sourceFile.InitialDirectory = ".";
            sourceFileDia.Filter = "c源文件|*.c";
            sourceFileDia.RestoreDirectory = false;

            if (DialogResult.OK == sourceFileDia.ShowDialog())
            {
                sourceFileTexBox.Text = sourceFileDia.FileName;

                funcListBox.Items.Clear();
                parseFile();
            }
        }

        private void parseFile()
        {
            string fileName = sourceFileTexBox.Text;

            if (File.Exists(fileName))
            {
                CppStructuralOutput output = new CppStructuralOutput();
                CppFileParser parser = new CppFileParser(output, fileName);

                List<FunctionDefine> funList = output.getFunctionList();
                
                foreach (FunctionDefine item in funList)
                {
                    funcListBox.Items.Add(item.FunctionName);
                    
                }

                if (!parser.Message.Equals("Successfully parsed file"))
                {
                    MessageBox.Show(parser.Message);
                }

                for (int i = 0; i < funcListBox.Items.Count; i++)
                {
                    funcListBox.SelectedIndex = i;
                }   
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            //funcListBox.Items.Add("hello");
            funcListBox.SelectionMode = SelectionMode.MultiExtended;
        }

        private void genBtn_Click(object sender, EventArgs e)
        {
            string fileName;
            fileName = System.IO.Path.GetDirectoryName(sourceFileTexBox.Text) + "\\";
            fileName += "test_" + System.IO.Path.GetFileName(sourceFileTexBox.Text);

            SaveFileDialog saveFileDia = new SaveFileDialog();
            saveFileDia.Filter = "c源文件|*.c";
            saveFileDia.FileName = fileName;

            DialogResult result = saveFileDia.ShowDialog();
            if (DialogResult.OK == result)
            {
                fileName = saveFileDia.FileName;
                genrateFile(fileName);
            }
        }

        private void genrateFile(string fileName)
        {
            if (File.Exists(fileName))
            {
                MessageBox.Show("目标文件已经存在！","错误");
                return;
            }

            List<FunctionDefine> funList = new List<FunctionDefine>();;

            for (int i = 0; i < funcListBox.SelectedItems.Count; i++)
            {
                FunctionDefine funcItem = new FunctionDefine();
                funcItem.FunctionName = funcListBox.SelectedItems[i].ToString();
                funcItem.FunctionName = "test_" + funcItem.FunctionName;
                funList.Add(funcItem);
            }
            try
            {
                TestFile.genNewFile(fileName, funList);
            }catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
            MessageBox.Show("生成成功！","成功");
        }

        private void aboutRightMenu_Click(object sender, EventArgs e)
        {
            Version ApplicationVersion = new Version(Application.ProductVersion);
            int ss = ApplicationVersion.Major;

            MessageBox.Show("版本:ver" + ss+".0\n联系 hanfei@pset.suntec.net", "关于");
        }

        private void funcListBox_SizeChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string fileName;
            //fileName = System.IO.Path.GetDirectoryName(sourceFileTexBox.Text) + "\\";
            fileName = "list_" + System.IO.Path.GetFileName(sourceFileTexBox.Text);
            fileName += ".xls";

            SaveFileDialog saveFileDia = new SaveFileDialog();
            saveFileDia.Filter = "Excel文件|*.xls";
            saveFileDia.FileName = fileName;

            DialogResult result = saveFileDia.ShowDialog();
            if (DialogResult.OK == result)
            {
                fileName = saveFileDia.FileName;
                if (File.Exists(fileName))
                {
                    MessageBox.Show("目标文件已经存在！", "错误");
                    return;
                }

                try
                {
                    IWorkbook wb = new HSSFWorkbook();
                    
                    ISheet tb = wb.CreateSheet(System.IO.Path.GetFileName(sourceFileTexBox.Text));
                    tb.DisplayGridlines = false;

                    tb.CreateRow(5).CreateCell(29).SetCellValue("備考");
                    tb.GetRow(5).CreateCell(3).SetCellValue("No.");
                    tb.GetRow(5).CreateCell(5).SetCellValue("関数名");
                    tb.GetRow(5).CreateCell(17).SetCellValue("テスト方法");

                    ICellStyle bg = (HSSFCellStyle)wb.CreateCellStyle();
                    IFont ft=wb.CreateFont();
                    ft.Color = NPOI.HSSF.Util.HSSFColor.White.Index;
                    bg.SetFont(ft);
                    HSSFWorkbook wob = new HSSFWorkbook();
                    HSSFPalette pa=wob.GetCustomPalette();
                    NPOI.HSSF.Util.HSSFColor XlColour = pa.FindSimilarColor(23, 55, 93);
                    bg.FillForegroundColor = XlColour.Indexed; //15000;
                    bg.FillPattern = FillPattern.SolidForeground;

                    tb.CreateRow(3);
                    for (int i = 0; i < 50; i++)
                    {
                        tb.GetRow(3).CreateCell(i).CellStyle = bg;
                    }
                    tb.GetRow(3).GetCell(1).SetCellValue(System.IO.Path.GetFileName(sourceFileTexBox.Text));

                    ICellStyle Border3 = (HSSFCellStyle)wb.CreateCellStyle();
                    Border3.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
                    Border3.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
                    Border3.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

                    tb.GetRow(5).GetCell(3).CellStyle = Border3;

                    ICellStyle Border2 = (HSSFCellStyle)wb.CreateCellStyle();
                    Border2.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
                    Border2.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

                    ICellStyle Border4 = (HSSFCellStyle)wb.CreateCellStyle();
                    Border4.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
                    Border4.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    Border4.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

                    ICellStyle Border5 = (HSSFCellStyle)wb.CreateCellStyle();
                    Border5.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
                    Border5.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
                    Border5.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

                    ICellStyle Border6 = (HSSFCellStyle)wb.CreateCellStyle();
                    Border6.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    Border6.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
                    Border6.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

                     ICellStyle Border7 = (HSSFCellStyle)wb.CreateCellStyle();
                    Border7.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    Border7.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    Border7.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

                    ICellStyle Border8 = (HSSFCellStyle)wb.CreateCellStyle();
                    Border8.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    Border8.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
                    Border8.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

                    ICellStyle Border9 = (HSSFCellStyle)wb.CreateCellStyle();
                    Border9.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    Border9.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
                    Border9.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;

                    ICellStyle Border10 = (HSSFCellStyle)wb.CreateCellStyle();
                    Border10.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    Border10.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;

                    ICellStyle Border11 = (HSSFCellStyle)wb.CreateCellStyle();
                    Border11.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    Border11.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    Border11.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;

                    ICellStyle Border12 = (HSSFCellStyle)wb.CreateCellStyle();
                    Border12.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    Border12.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
                    Border12.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
                
                    tb.GetRow(5).CreateCell(4).CellStyle = Border2;
                    tb.GetRow(5).GetCell(5).CellStyle = Border4;

                    for (int i = 6; i < 17; i++)
                    {
                        tb.GetRow(5).CreateCell(i).CellStyle = Border2;
                    }
                    tb.GetRow(5).GetCell(17).CellStyle = Border4;

                    for (int i = 18; i < 29; i++)
                    {
                        tb.GetRow(5).CreateCell(i).CellStyle = Border2;
                    }
                    tb.GetRow(5).GetCell(29).CellStyle = Border4;

                    for (int i = 30; i < 41; i++)
                    {
                        tb.GetRow(5).CreateCell(i).CellStyle = Border2;
                    }
                    tb.GetRow(5).CreateCell(41).CellStyle = Border5;

                    ICellStyle Border1 = (HSSFCellStyle)wb.CreateCellStyle();
                    Border1.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    Border1.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

                    int lastline=6;
                    for (int i = 0; i < funcListBox.SelectedItems.Count; i++)
                    {                    
                        tb.CreateRow(6 + i).CreateCell(5).SetCellValue(funcListBox.SelectedItems[i].ToString());
                        tb.GetRow(6 + i).CreateCell(3).SetCellValue(" " + (i + 1).ToString());

                       tb.GetRow(6 + i).GetCell(3).CellStyle = Border6;
          
                        tb.GetRow(6 + i).CreateCell(4).CellStyle = Border1;
                        tb.GetRow(6 + i).GetCell(5).CellStyle = Border7;
                       
                        for (int j = 6; j < 17; j++)
                        {
                            tb.GetRow(6 + i).CreateCell(j).CellStyle = Border1;
                        }
                        tb.GetRow(6 + i).CreateCell(17).CellStyle = Border7;
                         for (int j = 18; j < 29; j++)
                        {
                            tb.GetRow(6 + i).CreateCell(j).CellStyle = Border1;
                        }
                        tb.GetRow(6 + i).CreateCell(29).CellStyle = Border7;
                        for (int j = 30; j < 41; j++)
                        {
                            tb.GetRow(6 + i).CreateCell(j).CellStyle = Border1;
                        }
                        tb.GetRow(6 + i).CreateCell(41).CellStyle = Border8;

                        lastline += 1;
                    }

                    tb.CreateRow(lastline).CreateCell(3).CellStyle = Border9;
                    tb.GetRow(lastline).CreateCell(4).CellStyle = Border10;
                    tb.GetRow(lastline).CreateCell(5).CellStyle = Border11;
                    for (int j = 6; j < 17; j++)
                    {
                        tb.GetRow(lastline).CreateCell(j).CellStyle = Border10;
                    }
                    tb.GetRow(lastline).CreateCell(17).CellStyle = Border11;
                    for (int j = 18; j < 29; j++)
                    {
                        tb.GetRow(lastline).CreateCell(j).CellStyle = Border10;
                    }
                    tb.GetRow(lastline).CreateCell(29).CellStyle = Border11;
                    for (int j = 30; j < 41; j++)
                    {
                        tb.GetRow(lastline).CreateCell(j).CellStyle = Border10;
                    }
                    tb.GetRow(lastline).CreateCell(41).CellStyle = Border12;

                    for (int i = 0; i < 50; i++)
                    {
                        tb.SetColumnWidth(i,1024);                       
                    }
                   

                    using (FileStream fs = File.OpenWrite(fileName))
                    {
                        wb.Write(fs);
                     }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                        return;
                    }
                
                MessageBox.Show("生成成功！", "成功");
            }
        }
    }
}

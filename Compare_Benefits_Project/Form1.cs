using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Compare_Benefits_Project
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string fileExcel;

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files|*.xls;*.xlsx;*.csv";
            ofd.Title = "Open File Excel";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                fileExcel = ofd.FileName;
                textBox1.Text = ofd.SafeFileName;
            }
            else return;

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Worksheet xlWorkSheet2;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();



            //open workbook
            MessageBox.Show(ofd.FileName + " opened successfully.");
            xlWorkBook = xlApp.Workbooks.Open(fileExcel, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            xlApp.Visible = true;
            //***need try exception handler for when sheet 2 is not there

            //foreach (object item in xlWorkSheet)
            //{

            //}
            //add unique ID in column A sheet1
            Excel.Range oRng = xlWorkSheet.Range["A1"];
            oRng.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight,
                    Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            oRng = xlWorkSheet.Range["A1"];
            oRng.Value2 = "Uniq_ID";
            //Excel.Range rgcon2 = xlWorkSheet.Range["A2:A {0}", rowCount];
            //maybe a forEach cell in column add formula


            
            var rCnt = 1;
            var cCnt = 1;

            Excel.Range range = xlWorkSheet.UsedRange;

            rCnt = range.Rows.Count;
            cCnt = range.Columns.Count;


            //**********research more about cell range, would like to be able to replace A33 with actual rowCount "xlWorkSheet.Range["A2:A {0}", rowCount];"
            Excel.Range concat = xlWorkSheet.Range["A2:A33"];
            

            //int rowCount;
            //Excel.Range rgcon2 = xlWorkSheet.Range(xlWorkSheet.Cells[iRow, column]).Formula = string.Format("=SUM(G1,G{0})", iRow)
            //(((Excel.Range))yourWorkSheet.Cells[rowCount, column]).Formula = string.Format("=SUM(G1,G{0})", rowCount);
            //REVISIT*********final may not be these exact cells**************************************
            concat.Formula = string.Format("=CONCATENATE(C2,D2,H2,J2)");

            //conditional formating to find matches

            Excel.Range condFormatRange = xlWorkSheet.get_Range("C1", "C2");
            Excel.FormatConditions fcs = condFormatRange.FormatConditions;
            object Formula1 = "=IF($C$1 = $C$2)";
            //Excel.FormatCondition fc = (Excel.FormatCondition)fcs.Add
            //(Excel.XlFormatConditionType.xlExpression, Type.Missing, "=IF($C$1 = $C$2");
            Excel.FormatCondition fc = (Excel.FormatCondition)fcs.Add(Excel.XlFormatConditionType.xlCellValue, Formula1);
            Excel.Interior interior = fc.Interior;
            interior.Color = ColorTranslator.ToOle(Color.Red);
            interior = null;
            fc = null;
            fcs = null;
            condFormatRange = null;

            //add unique ID in column A sheet2
            //replace with a method- pass worksheet in then add uniqID column and concatenate formula
            Excel.Range oRng2 = xlWorkSheet2.Range["A1"];
            oRng2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight,
                    Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            oRng2 = xlWorkSheet2.Range["A1"];
            oRng2.Value2 = "Uniq_ID";
            Excel.Range newRange = xlWorkSheet2.get_Range("A2","A33");
            newRange.Formula = string.Format("=CONCATENATE(C2,D2,H2,J2)");
            newRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            //xlWorkSheet2.Cells[1, 2] = "Red";

            //xlWorkBook.Close(true, misValue, misValue);
            //xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files|*.xls;*.xlsx;*.csv";
            ofd.Title = "Open File Excel";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                fileExcel = ofd.FileName;
                textBox2.Text = ofd.SafeFileName;
            }
            else return;

            //call method to add uniqID and concatenate
            uniqConcat();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //add code to combine selected excel workbooks into COMPARE_WB 

        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void uniqConcat()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Worksheet xlWorkSheet2;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();

            //open workbook
            
            xlWorkBook = xlApp.Workbooks.Open(fileExcel, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            xlApp.Visible = true;
            //add unique ID in column A sheet1
            Excel.Range oRng = xlWorkSheet.Range["A1"];
            oRng.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight,
                    Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            oRng = xlWorkSheet.Range["A1"];
            oRng.Value2 = "Uniq_ID";

            Excel.Range concat = xlWorkSheet.Range["A2:A33"];
            concat.Formula = string.Format("=CONCATENATE(D2,E2,I2,K2)");
        }
    }
}

using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

namespace BessolitsynExcelToolAdd_in
{
    public partial class RibbonBessolitsyn
    {
        Excel.Worksheet ActiveSheet;
        List<string> Range1, Range2;

        // List as result of R1-R2
        List<string> R1_R2 = new List<string>();

        // List as result of R2-R1
        List<string> R2_R1 = new List<string>();

        //Общий диапазон
        List<string> intersectionOf_R1_R2 = new List<string>();

        private void RibbonBessolitsyn_Load(object sender, RibbonUIEventArgs e)
        {
            //var Window = Globals.ThisAddIn.Application.ActiveWindow;
            //Excel.Workbook aWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //ActiveSheet = aWorkBook.ActiveSheet;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Range2 = RangeToList(Globals.ThisAddIn.Application.Selection as Excel.Range);

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Range1 = RangeToList(Globals.ThisAddIn.Application.Selection as Excel.Range);

        }
        //Range processing
        private void button4_Click(object sender, RibbonControlEventArgs e)
        {

            // List as result of R1-R2
            List<string> R1_R2 = new List<string>();
            R1_R2 = Range1;

            // List as result of R2-R1
            List<string> R2_R1 = new List<string>();

            //Общий диапазон
            List<string> intersectionOf_R1_R2 = new List<string>();

            foreach (var item in Range2)
            {

                try
                {
                    if (R1_R2.Remove(item))
                    {
                        intersectionOf_R1_R2.Add(item);
                    }
                    else {
                        R2_R1.Add(item);

                    }

                }
                catch (Exception Ex)
                {
                    MessageBox.Show(Ex.ToString());
                }
            }

            

            TextWriter w1 = new StreamWriter("R1_R2.txt");
            foreach (String s in R1_R2)
                w1.WriteLine(s);

            w1.Close();

            w1= new StreamWriter("R2_R1.txt");
            foreach (String s in R2_R1)
                w1.WriteLine(s);
            w1.Close();

            w1 = new StreamWriter("intersectionOf_R1_R2");
            foreach (String s in R2_R1)
                w1.WriteLine(s);

            w1.Close();


        }

        private List<string> RangeToList(Excel.Range range)
        {
            List<string> result = new List<string>();
            foreach (Excel.Range row in range.Rows)
            {
                var cell = (Excel.Range)row.Cells[1, 1];
                result.Add(Convert.ToString(cell.Value2));
            }
            return result;
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            Openfile("intersectionOf_R1_R2");
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            Openfile("R2_R1.txt");
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Openfile("R1_R2.txt");
        }

        public void Openfile(string file)
        {
            Process myProcess = new Process();
            Process.Start(@"c:\Program Files\Notepad++\notepad++.exe", "\""+file+"\"");
        }
    }
}

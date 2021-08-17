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
        List<KeyValuePair<string,string>> RangeWithIds;
        List<KeyValuePair<string, string>> RangeAB;
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
            R1_R2 = new List<string>();
            R1_R2 = Range1;

            // List as result of R2-R1
            R2_R1 = new List<string>();

            //Общий диапазон
            intersectionOf_R1_R2 = new List<string>();

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




            string path = "R1_R2.txt";
            FileInfo fileInf = new FileInfo(path);
            if (fileInf.Exists)
            {
                fileInf.Delete();
            }
            path = "R2_R1.txt";
            fileInf = new FileInfo(path);
            if (fileInf.Exists)
            {
                fileInf.Delete();
            }

            path = "intersectionOf_R1_R2";
            fileInf = new FileInfo(path);
            if (fileInf.Exists)
            {
                fileInf.Delete();
            }


            StreamWriter w1 = new StreamWriter("R1_R2.txt");
            foreach (String s in R1_R2)
                w1.WriteLine(s);
            w1.Close();

            w1= new StreamWriter("R2_R1.txt");
            foreach (String s in R2_R1)
                w1.WriteLine(s);
            w1.Close();

            w1 = new StreamWriter("intersectionOf_R1_R2");
            foreach (String s in intersectionOf_R1_R2)
                w1.WriteLine(s);
            w1.Close();


        }

        private List<string> RangeToList(Excel.Range range)
        {
            List<string> result = new List<string>();
            foreach (Excel.Range row in range.Rows)
            {
                var cell = (Excel.Range)row.Cells[1, 1];
                
                result.Add(Convert.ToString(cell.Value2)?.ToLower());
            }
            return result;
        }
        private List<KeyValuePair<string, string>> RangeWithIDsToList(Excel.Range range)
        {
            List<KeyValuePair<string, string>> result = new List<KeyValuePair<string, string>>();
            foreach (Excel.Range row in range.Rows)
            {
                var cell1 = Convert.ToString(row.Cells[1, 1].Value2)?.ToLower();
                var cell2 = Convert.ToString(row.Cells[1, 2].Value2);
                result.Add(new KeyValuePair<string, string>(cell1, cell2));
            }
            return result;
        }

        private List<KeyValuePair<string, string>> RangeABToList(Excel.Range range)
        {
            List<KeyValuePair<string, string>> result = new List<KeyValuePair<string, string>>();
            foreach (Excel.Range row in range.Rows)
            {
                var cell1 = Convert.ToString(row.Cells[1, 1].Text);
                var cell2 = Convert.ToString(row.Cells[1, 2].Text);
                result.Add(new KeyValuePair<string, string>(cell1, cell2));
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

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            //Выделить выбранный диапозон
            int caseSwitch = dropDown1.SelectedItemIndex;
            string marker = dropDown1.SelectedItem.Label;

            switch (caseSwitch)
            {
                case 0:
                    PasteResults(R1_R2, marker);
                    break;
                case 1:
                    PasteResults(R2_R1, marker);
                    break;
                case 2:
                    PasteResults(intersectionOf_R1_R2, marker);
                    break;
                default:
                    break;
            }

        }

        public void Openfile(string file)
        {
            Process myProcess = new Process();
            Process.Start(@"c:\Program Files\Notepad++\notepad++.exe", "\""+file+"\"");
        }
        //find IDs
        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
            foreach (Excel.Range row in selection.Rows)
            {
                var cell = Convert.ToString(row.Cells[1, 1].Value2)?.ToLower();
                if (RangeWithIds.Any(item=> item.Key==cell))
                {
                    (row.Cells[1, 2] as Excel.Range).Value = RangeWithIds.Single(item => item.Key == cell).Value;

                }
                else row.Cells[1, 2].Value= "not found";
            }

        }
        //Set range with IDs
        private void button10_Click(object sender, RibbonControlEventArgs e)
        {
            RangeWithIds = RangeWithIDsToList(Globals.ThisAddIn.Application.Selection as Excel.Range);
        }

        //set r[a,b]
        private void button11_Click(object sender, RibbonControlEventArgs e)
        {
            RangeAB = RangeABToList(Globals.ThisAddIn.Application.Selection as Excel.Range);
        }
        //get a 
        private void button12_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
            foreach (Excel.Range row in selection.Rows)
            {
                var cell = ((Excel.Range)row.Cells[1, 1]).Text;
                if ((cell!= null ) && (RangeAB.Any(item => item.Value == cell)))
                {
                    try
                    {

                    (row.Cells[1, 1] as Excel.Range).Value = RangeAB.Single(item => item.Value!=null && item.Value == cell).Key;
                    }
                    catch (Exception)
                    {
                    string message = "Найдены значения:";
                    RangeAB.Where(item => item.Value != null && item.Value == cell).ToList().ForEach(i => message = message + " " + i.Key);
                    MessageBox.Show(message);
                    }

                }
            
            }
        }
        //get b
        private void button13_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
            foreach (Excel.Range row in selection.Rows)
            {
                var cell = (string)((Excel.Range)row.Cells[1, 1]).Text;
                if ((cell != null) && (RangeAB.Any(item => item.Key == cell)))
                {
                    try
                    {
                        (row.Cells[1, 1] as Excel.Range).Value = RangeAB.Single(item => item.Key != null && item.Key == cell).Value;
                    }
                    catch (Exception)
                    {
                        string message = "Найдены значения:";
                        RangeAB.Where(item => item.Key != null && item.Key == cell).ToList().ForEach(i=> message = message + " " + i.Value);
                        MessageBox.Show(message);
                    }

                }
            }
        }

        public void PasteResults(List<string> rangeList, string marker)
        {
            
            var selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
            foreach (Excel.Range row in selection.Rows)
            {
                var cell = (Excel.Range)row.Cells[1, 1];
                var celLValue = Convert.ToString(cell?.Value2)?.ToLower();
                if (celLValue != null)
                {
                    if (rangeList.Contains(celLValue))
                    {
                    (row.Cells[1, 2] as Excel.Range).Value = marker;

                }
                }
            }
        }
    }
}

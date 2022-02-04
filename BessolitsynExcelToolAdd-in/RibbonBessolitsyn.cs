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
        List<KeyValuePair<string, Excel.Range>> RangeWithIdsCol2;
        List<KeyValuePair<string, string>> RangeWithIdsCol3;
        List<KeyValuePair<string, string>> RangeWithIdsCol4;
        List<KeyValuePair<string, string>> RangeWithIdsCol5;
        List<KeyValuePair<string, string>> RangeWithIdsCol6;
        List<KeyValuePair<string, string>> RangeWithIdsCol7;
        List<KeyValuePair<string, string>> RangeWithIdsCol8;
        List<KeyValuePair<string, string>> RangeWithIdsCol9;
        List<KeyValuePair<string, string>> RangeWithIdsCol10;



        List<KeyValuePair<string, string>> RangeAB;
        // List as result of R1-R2
        List<string> R1_R2;

        // List as result of R2-R1
        List<string> R2_R1;

        //Общий диапазон
        List<string> intersectionOf_R1_R2;

        List<string> restOf_R1_R2 = new List<string>();

        List<string> ResultR1Rn;
        private void RibbonBessolitsyn_Load(object sender, RibbonUIEventArgs e)
        {
            //var Window = Globals.ThisAddIn.Application.ActiveWindow;
            //Excel.Workbook aWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //ActiveSheet = aWorkBook.ActiveSheet;
        }


        private void IntersectionR1Rn_Click(object sender, RibbonControlEventArgs e)
        {
            List<string> IntersectionR1R2(List<string> r1, List<string> r2)
            {
                //Общий диапазон
                var intersectionOf_R1_R2 = new List<string>();

                foreach (var item in r2)
                {

                    try
                    {
                        if (r1.Remove(item))
                        {
                            intersectionOf_R1_R2.Add(item);
                        }
                        

                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show(Ex.ToString());
                    }
                }
                return intersectionOf_R1_R2;
            }



            var R1toRnRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            var listOfRanges = new List<List<string>>();
            foreach (Excel.Range col in R1toRnRange.Columns)
            {
                listOfRanges.Add(RangeToList(col));

            }

            ResultR1Rn = listOfRanges[0];
            for (int i = 1; i < listOfRanges.Count; i++)
            {
                ResultR1Rn = IntersectionR1R2(ResultR1Rn, listOfRanges[i]);
            }

            foreach (Excel.Range col in R1toRnRange.Columns)
            {
                foreach (Excel.Range row in col.Rows)
                {
                    var cell = (Excel.Range)row.Cells[1, 1];
                    var celLValue = Convert.ToString(cell?.Value2);
                    if (celLValue != null)
                    {
                        if (!ResultR1Rn.Contains(celLValue))
                        {
                            cell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                        }
                    }

                }

            }


        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Range2 = RangeToListAndLowerCellValues(Globals.ThisAddIn.Application.Selection as Excel.Range);

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Range1 = RangeToListAndLowerCellValues(Globals.ThisAddIn.Application.Selection as Excel.Range);

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

        private List<string> RangeToListAndLowerCellValues(Excel.Range range)
        {
            List<string> result = new List<string>();
            foreach (Excel.Range row in range.Rows)
            {
                var cell = (Excel.Range)row.Cells[1, 1];
                
                result.Add(Convert.ToString(cell.Value2)?.ToLower());
            }
            return result;
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

        private List<KeyValuePair<string, Excel.Range>> RangeWithIDsColNToList(Excel.Range range, int Col=0)
        {
            List<KeyValuePair<string, Excel.Range>> result = new List<KeyValuePair<string, Excel.Range>>();
            foreach (Excel.Range row in range.Rows)
            {
                var cell1 = Convert.ToString(row.Cells[1, 1].Value2)?.ToLower();
                //var cell2 = Convert.ToString(row.Cells[1, Col].Value2);
                Excel.Range cell2 = row;
                result.Add(new KeyValuePair<string, Excel.Range>(cell1, cell2));
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
            
            //Process.Start(@"c:\Program Files\Notepad++\notepad++.exe", "\""+file+"\"");
            Process.Start(@"C:\Program Files (x86)\Notepad++\notepad++.exe", "\"" + file + "\"");

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
                    var range = RangeWithIds.Where(item => item.Key == cell).ToList();
                    if (range.Count==1)
                        (row.Cells[1, 2] as Excel.Range).Value = range.First().Value;
                    else
                        row.Cells[1, 2].Value = "more than one found";

                }
                else row.Cells[1, 2].Value= "not found";
            }

            if (int.Parse(editBox1.Text) > 1)
            {
                int Ncolumns = int.Parse(editBox1.Text);
                foreach (Excel.Range row in selection.Rows)
                {
                    var cell = Convert.ToString(row.Cells[1, 1].Value2)?.ToLower();
                    if (RangeWithIds.Any(item => item.Key == cell))
                    {
                        var range = RangeWithIdsCol2.Where(item => item.Key == cell).ToList();
                        if (range.Count == 1)
                            for (int i = 2; i < Ncolumns + 1; i++)
                            {
                                (row.Cells[1, i] as Excel.Range).Value = range.First().Value.Cells[1, i].Value;

                            }
                        else
                            row.Cells[1, 2].Value = "more than one found";

                       

                    }
                }

            }

        }
        //Set range with IDs
        private void button10_Click(object sender, RibbonControlEventArgs e)
        {
            RangeWithIds = RangeWithIDsToList(Globals.ThisAddIn.Application.Selection as Excel.Range);
            RangeWithIdsCol2 = RangeWithIDsColNToList(Globals.ThisAddIn.Application.Selection as Excel.Range);
            //RangeWithIdsCol4 = RangeWithIDsColNToList(Globals.ThisAddIn.Application.Selection as Excel.Range, 3);
            //RangeWithIdsCol5 = RangeWithIDsColNToList(Globals.ThisAddIn.Application.Selection as Excel.Range, 3);
            //RangeWithIdsCol6 = RangeWithIDsColNToList(Globals.ThisAddIn.Application.Selection as Excel.Range, 3);
            //RangeWithIdsCol7 = RangeWithIDsColNToList(Globals.ThisAddIn.Application.Selection as Excel.Range, 3);
            //RangeWithIdsCol8 = RangeWithIDsColNToList(Globals.ThisAddIn.Application.Selection as Excel.Range, 3);
            //RangeWithIdsCol9 = RangeWithIDsColNToList(Globals.ThisAddIn.Application.Selection as Excel.Range, 3);
            //RangeWithIdsCol10 = RangeWithIDsColNToList(Globals.ThisAddIn.Application.Selection as Excel.Range, 3);
            

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

        List<PBS2class> L1 = new List<PBS2class>();
        List<PBS2class> L2 = new List<PBS2class>();
        List<PBS2class> L3 = new List<PBS2class>();
        List<PBS2class> L4 = new List<PBS2class>();
        private void button14_Click(object sender, RibbonControlEventArgs e)

        {
            
            Excel.Range AtributeArea = Globals.ThisAddIn.Application.get_Range("A1:KA1");

            
            Excel.Range rng = Globals.ThisAddIn.Application.Selection as Excel.Range;
            //Excel.Range rng = Globals.ThisAddIn.Application.get_Range("A6:KA100");
            foreach (Excel.Range row in rng.Rows)
            {
                PBS2class obj = new PBS2class(row, AtributeArea);
                int length = obj.ProcessClassCode.Length;
                switch (length)
                {
                    case 1:
                        obj.Level = 1;
                        L1.Add(obj);
                        break;

                    case 2:
                        obj.Level = 2;
                        L2.Add(obj);
                        break;

                    case 3:
                        obj.Level = 3;
                        L3.Add(obj);
                        break;

                    case 5:
                        obj.Level = 4;
                        L4.Add(obj);
                        break;

                    default:
                        break;
                }

            }
            foreach (var item in L3)
            {
                var members = L4.Where(x => x.ProcessClassCode.Contains(item.ProcessClassCode)).ToList();
                item.Members.AddRange(members);
            }

            foreach (var item in L2)
            {
                var members = L3.Where(x => x.ProcessClassCode.Contains(item.ProcessClassCode)).ToList();
                item.Members.AddRange(members);
            }
            foreach (var item in L1)
            {
                var members = L2.Where(x => x.ProcessClassCode.Contains(item.ProcessClassCode)).ToList();
                item.Members.AddRange(members);
            }
            //List<string> Strings = new List<string>();
            //Strings.Add("L1;L2;L3;L4;ProcessClassCode;Attribute;TagIdentifier");

            //foreach (var item1 in L1)
            //{
            //    Strings.Add(String.Format("{0};{1};{2};{3};{4};{5};{6}", item1.ProcessClass, "", "", "", item1.ProcessClassCode, "", ""));
            //    foreach (var item2 in L2)
            //    {
            //        Strings.Add(String.Format("{0};{1};{2};{3};{4};{5};{6}", item1.ProcessClass, item2.ProcessClass, "", "", item2.ProcessClassCode, "", ""));
            //        foreach (var item3 in item2.Members)
            //        {
            //            Strings.Add(String.Format("{0};{1};{2};{3};{4};{5};{6}", item1.ProcessClass, item2.ProcessClass, item3.ProcessClass, "", item3.ProcessClassCode, "", ""));
            //            foreach (var item4 in item3.Members)
            //            {
            //                Strings.Add(String.Format("{0};{1};{2};{3};{4};{5};{6}", item1.ProcessClass, item2.ProcessClass, item3.ProcessClass, item4.ProcessClass, item4.ProcessClassCode, "", item4.TagIdentifier));
            //                foreach (var att in item4.Attributes)
            //                {
            //                    Strings.Add(String.Format("{0};{1};{2};{3};{4};{5};{6}", item1.ProcessClass, item2.ProcessClass, item3.ProcessClass, item4.ProcessClass, item4.ProcessClassCode, att.NewStandardAttribute, item4.TagIdentifier));

            //                }

            //            }
            //        }
            //    }
            //}
            //foreach (var item1 in L1)
            //{
            //    Strings.Add(String.Format("{0};{1};{2};{3};{4};{5};{6}", item1.ProcessClass, "", "", "", item1.ProcessClassCode, "", ""));
            //    foreach (var item2 in L2)
            //    {
            //        Strings.Add(String.Format("{0};{1};{2};{3};{4};{5};{6}", item1.ProcessClass, item2.ProcessClass, "", "", item2.ProcessClassCode, "", item2.TagIdentifier));
            //        foreach (var att in item2.Attributes)
            //        {
            //            Strings.Add(String.Format("{0};{1};{2};{3};{4};{5};{6}", item1.ProcessClass, item2.ProcessClass, "", "", item2.ProcessClassCode, att.NewStandardAttribute, item2.TagIdentifier));

            //        }
                    
            //    }
            //}


            //var path = "pbs2.csv";
            //var fileInf = new FileInfo(path);
            //if (fileInf.Exists)
            //{
            //    fileInf.Delete();
            //}


            //var file = new StreamWriter("pbs2.csv");
            //foreach (String s in Strings)
            //    file.WriteLine(s);
            //file.Close();
            //Openfile(fileInf.FullName);

        }

        private void button15_Click(object sender, RibbonControlEventArgs e)
        {
            string PreviosAttCat="";
            void PasteCommon(PBS2class item, int ii, dynamic aSheet, string att = "", string attOption="",string attCategory = "", string PartIs = "")
            {

                if (!(attCategory == ""))
                    PreviosAttCat = attCategory;
                ((Excel.Range)aSheet.Cells[ii, 1]).Value2 = item.TagList;
                ((Excel.Range)aSheet.Cells[ii, 2]).Value2 = item.DisciplineOwner;
                ((Excel.Range)aSheet.Cells[ii, 3]).Value2 = item.l1;
                ((Excel.Range)aSheet.Cells[ii, 4]).Value2 = item.l2;
                ((Excel.Range)aSheet.Cells[ii, 5]).Value2 = item.l3;
                ((Excel.Range)aSheet.Cells[ii, 6]).Value2 = item.l4;
                ((Excel.Range)aSheet.Cells[ii, 7]).Value2 = att;
                ((Excel.Range)aSheet.Cells[ii, 8]).Value2 = item.TagIdentifier;
                ((Excel.Range)aSheet.Cells[ii, 9]).Value2 = item.ProcessClassCode;
                ((Excel.Range)aSheet.Cells[ii, 10]).Value2 = item.ProcessClass;
                ((Excel.Range)aSheet.Cells[ii, 11]).Value2 = item.ProcessClassType;
                ((Excel.Range)aSheet.Cells[ii, 12]).Value2 = attOption;
                ((Excel.Range)aSheet.Cells[ii, 13]).Value2 = item.ProcessClassCode.Length.ToString();
                ((Excel.Range)aSheet.Cells[ii, 14]).Value2 = PreviosAttCat;
                ((Excel.Range)aSheet.Cells[ii, 15]).Value2 = PartIs;
            }

            //"L1;L2;L3;L4;ProcessClassCode;Attribute;TagIdentifier"
            Excel.Worksheet activeSheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);

            ((Excel.Range)activeSheet.Cells[1, 1]).Value2 = "tag list(Scope of work)";
            ((Excel.Range)activeSheet.Cells[1, 2]).Value2 = "Discipline owner";

            ((Excel.Range)activeSheet.Cells[1, 3]).Value2 = "L1";
            ((Excel.Range)activeSheet.Cells[1, 4]).Value2 = "L2";
            ((Excel.Range)activeSheet.Cells[1, 5]).Value2 = "L3";
            ((Excel.Range)activeSheet.Cells[1, 6]).Value2 = "L4";

            ((Excel.Range)activeSheet.Cells[1, 7]).Value2 = "Attribute";
            ((Excel.Range)activeSheet.Cells[1, 8]).Value2 = "TagIdentifier";

            ((Excel.Range)activeSheet.Cells[1, 9]).Value2 = "Process Class Code";
            ((Excel.Range)activeSheet.Cells[1, 10]).Value2 = "Process Class";
            ((Excel.Range)activeSheet.Cells[1, 11]).Value2 = "Process Class Type";
            ((Excel.Range)activeSheet.Cells[1, 12]).Value2 = "Attribute Option";
            ((Excel.Range)activeSheet.Cells[1, 13]).Value2 = "Level";
            ((Excel.Range)activeSheet.Cells[1, 14]).Value2 = "Attribute Category";
            ((Excel.Range)activeSheet.Cells[1, 15]).Value2 = "Attribute is part of ";
            int i = 2;


            foreach (var item1 in L2)
            {
                PasteCommon(item1, i, activeSheet);

                ((Excel.Range)activeSheet.Cells[i, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                ((Excel.Range)activeSheet.Cells[i, 4]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                i++;
                if (checkBox2.Checked)
                    foreach (var att in item1.Attributes)
                    {
                        PasteCommon(item1, i, activeSheet, att.NewStandardAttribute, att.Option, att.AttributeCategory, att.AttributeIsPartOf);
                        i++;
                    }
                foreach (var item2 in item1.Members)
                {

                    PasteCommon(item2, i, activeSheet);
                    ((Excel.Range)activeSheet.Cells[i, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    ((Excel.Range)activeSheet.Cells[i, 4]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    ((Excel.Range)activeSheet.Cells[i, 5]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    i++;
                    if (checkBox3.Checked)
                        foreach (var att in item2.Attributes)
                        {
                            PasteCommon(item2, i, activeSheet, att.NewStandardAttribute, att.Option, att.AttributeCategory);
                            i++;
                        }
                    foreach (var item3 in item2.Members)
                    {
                        PasteCommon(item3, i, activeSheet);
                        ((Excel.Range)activeSheet.Cells[i, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan);
                        ((Excel.Range)activeSheet.Cells[i, 4]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan);
                        ((Excel.Range)activeSheet.Cells[i, 5]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan);
                        ((Excel.Range)activeSheet.Cells[i, 6]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan);

                        i++;
                        if (checkBox3.Checked)
                            foreach (var att in item3.Attributes)
                            {
                                PasteCommon(item3, i, activeSheet, att.NewStandardAttribute, att.Option, att.AttributeCategory);
                                i++;
                            }
                    }

                }
            }

        }

        private void checkBox1_Click(object sender, RibbonControlEventArgs e)
        {
            if (checkBox1.Checked)
                if (dropDown1.SelectedItem != null)
                    if (dropDown1.SelectedItemIndex == 2)
                    { 
                        var range = Globals.ThisAddIn.Application.Selection as Excel.Range;
                        foreach (Excel.Range row in range.Rows)
                        {
                            var cell = (Excel.Range)row.Cells[1, 1];
                            var celLValue = Convert.ToString(cell?.Value2)?.ToLower();
                            if (intersectionOf_R1_R2.Contains(celLValue))
                            {
                                row.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }

                            string[] words = celLValue.Split(' ');
                            string NewComment="";
                            foreach (var word in words)
                            {
                                var result = R1_R2.Where(a=>a.Contains(word)).ToList();
                                if (result.Count()>0)
                                {
                                    foreach (var att in result)
                                    {
                                    NewComment += att + "\r\n";

                                    }
                                    
                                }
                            }
                            if (NewComment!="") row.AddComment(NewComment);

                        }

                    }

        }

        private void button16_Click(object sender, RibbonControlEventArgs e)
        {
            var selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
            for (int i = 0; i < ResultR1Rn.Count; i++)
            {
                (selection.Cells[i + 1, 1] as Excel.Range).Value = ResultR1Rn[i];
            }
            
        }
        //Summ
        private void SumR1Rn_Click(object sender, RibbonControlEventArgs e)
        {
            List<string> SumR1R2(List<string> r1, List<string> result)
            {
                //Общий диапазон
                

                foreach (var item in r1)
                {

                    try
                    {
                        if (!result.Any(str=>str==item))
                        {
                            result.Add(item);
                        }


                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show(Ex.ToString());
                    }
                }
                return result;
            }



            var R1toRnRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            var listOfRanges = new List<List<string>>();
            foreach (Excel.Range col in R1toRnRange.Columns)
            {
                listOfRanges.Add(RangeToList(col));

            }

            ResultR1Rn = listOfRanges[0];
            for (int i = 1; i < listOfRanges.Count; i++)
            {
                ResultR1Rn = SumR1R2(listOfRanges[i], ResultR1Rn);
            }
            ResultR1Rn.Remove("");


        }

        private void comboBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {

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

    public partial class Helper
    {
        public static string getValue(dynamic singlecell)
        {
            var cell = (Excel.Range)singlecell;
            string celLValue = Convert.ToString(cell?.Value2);
            if (celLValue != null)
            {
                return celLValue;
            }
            else
            {
                return "";
            }
        }

     
    }

    public class PBS2class
    {
        public string TagList { get; set; }
        public string DisciplineOwner { get; set; }
        public string ProcessClassCode { get; set; }
        public string ProcessClass { get; set; }
        public string TagIdentifier { get; set; }
        public int Level { get; set; }
        public List<PBS2class> Members { get; set; }
        public string l1 { get; set; }
        public string l2 { get; set; }
        public string l3 { get; set; }
        public string l4 { get; set; }
        public string ProcessClassType { get; set; }
        public PBS2class Owner { get; set; }
        public List<PBS2attribute> Attributes { get; set; }


        public PBS2class(Excel.Range row, Excel.Range AtributeArea)
        {
            TagList = Helper.getValue(row.Cells[1, 1]);
            DisciplineOwner = Helper.getValue(row.Cells[1, 2]);
            Attributes = new List<PBS2attribute>();
            Members = new List<PBS2class>();
            l1=Helper.getValue(row.Cells[1, 3]);
            l2=Helper.getValue(row.Cells[1, 4]);
            l3=Helper.getValue(row.Cells[1, 5]);
            l4=Helper.getValue(row.Cells[1, 6]);
            ProcessClassCode = Helper.getValue(row.Cells[1, 7]);
            ProcessClass = Helper.getValue(row.Cells[1, 8]);
            ProcessClassType = Helper.getValue(row.Cells[1, 9]);
            TagIdentifier = Helper.getValue(row.Cells[1, 10]);
            //12-23
            for (int i = 12; i < 288; i++)
            {
                var AttFlag = Helper.getValue(row.Cells[1, i]);
                if ((AttFlag == "X") || (AttFlag == "C") || (AttFlag == "O") || (AttFlag == "0"))
                {
                    var Att = Helper.getValue(AtributeArea.Cells[1, i]);

                    var Cat = Helper.getValue(AtributeArea.Cells[4, i]);
                    var AttributeIsPartOf = Helper.getValue(AtributeArea.Cells[2, i]);
                    Attributes.Add(new PBS2attribute(Att, AttFlag, Cat, AttributeIsPartOf));
                }
            }
        }
    }
    public class PBS2attribute
    {
        public string AttributeID { get; set; }
        public string NewStandardAttribute { get; set; }
        public string Option { get; set; }
        public string AttributeCategory { get; set; }
        public string AttributeIsPartOf { get; set; }

        public PBS2attribute(string name, string attFlag, string category="", string attributeIsPartOf="")
        {
            NewStandardAttribute = name;
            Option = attFlag;
            AttributeCategory = category;
            AttributeIsPartOf = attributeIsPartOf;
        }

    }
}

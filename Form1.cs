using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Scripting;
using Microsoft.Scripting.Hosting;
 
 
//파이선을 실행하기 위해 추가
using IronPython;
using IronPython.Hosting;
using IronPython.Runtime;
using IronPython.Modules;
using System.Text.RegularExpressions;
using Excel=Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;



namespace assignment1
{
    public partial class Form1 : Form
    {


        string scan = "";
        string vaccine = "";
        string result = "";
        string update = "";
        string version = "";
        string sha1 = "";
        string md5 = "";
        
        Microsoft.Office.Interop.Excel.Application oXL;
        Microsoft.Office.Interop.Excel._Workbook oWB;
        Microsoft.Office.Interop.Excel._Worksheet oSheet;
        Microsoft.Office.Interop.Excel.Range oRng;
        object misvalue = System.Reflection.Missing.Value;


        public Form1()

        {
                      InitializeComponent();
        }


        public void startExcel()
        {
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;

            //Get a new workbook.
            oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
        }

        public void endExcel()
        {


            oXL.Visible = false;
            oXL.UserControl = false;
            oWB.SaveAs("\\resultvirus.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing,
            false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            oWB.Close();
            oXL = null;
        }
        public void Excelcheck(int row, string vac, string re, string up, string ver,string filenames, string sha,string md)
        {
            try
        {

            if (row == 1)
            {
                
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oXL.Worksheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                oSheet.Name = filenames;
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                oSheet.Cells[1, 1] = "sha1";
                oRng = oSheet.get_Range("B1", "E1");
                oRng.Merge(true);
                oSheet.Cells[1, 2] = sha;
                oRng = oSheet.get_Range("A1", "E1");
                oRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                oRng.Borders.Weight = Excel.XlBorderWeight.xlThin;
                oRng.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                oRng.ColumnWidth = 17;

                oSheet.Cells[2, 1] = "md5";
                oRng = oSheet.get_Range("B2", "E2");
                oRng.Merge(true);
                oSheet.Cells[2, 2] = md;
                oRng = oSheet.get_Range("A2", "E2");
                oRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                oRng.Borders.Weight = Excel.XlBorderWeight.xlThin;
                oRng.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                oSheet.Cells[3, 1] = "제품";
                oSheet.Cells[3, 2] = "결과";
                oSheet.Cells[3, 3] = "업데이트";
                oSheet.Cells[3, 4] = "버전";

                oRng = oSheet.get_Range("A3", "D3");
                oRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                oRng.Borders.Weight = Excel.XlBorderWeight.xlThin;
                oRng.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                oSheet.Range["D:D"].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            }

 
        //Add table headers going cell by cell.

        oSheet.Cells[row + 3, 1] = vac;
        oSheet.Cells[row + 3, 2] = re;
        oSheet.Cells[row + 3, 3] = up;
        oSheet.Cells[row + 3, 4] = ver;
        
        }
        catch(Exception e)
        {
            MessageBox.Show(e.ToString());
        }
        }


        public void totalvirus(string path,string name)
        {
            var engine = IronPython.Hosting.Python.CreateEngine();
            //engine.SetSearchPaths(new[] { "C:\\Users\\Administrator\\Desktop\\IronPython-2.7.7\\Lib" });
            engine.SetSearchPaths(new[] { "IronPython-2.7.7\\Lib" });
            var scope = engine.CreateScope();

            try
            {
                //파이선 프로그램 파일 실행.
                var source = engine.CreateScriptSourceFromFile("vtpy.py");
              
                source.Execute(scope);

                
                // call class MyClass
                var myClass = scope.GetVariable<Func<object>>("vt");
               
                var myInstance = myClass();
               

                //Console.WriteLine(engine.Operations.GetMember(myInstance, "file"));

                
                var classMethod = engine.Operations.GetMember(myInstance, "getfile");
               
                 
                scan = classMethod(path);
               

                scan = scan.Replace("{\"scans\":", "");

                Match regex = Regex.Match(scan, "\"(.*?)\": {\"detected\": (.*?), \"version\": (.*?), \"result\": (.*?), \"update\": (.*?)}", RegexOptions.IgnoreCase);

                Match crypto = Regex.Match(scan, "\"sha1\": (.*?), (.*?), \"md5\": (.*?)}", RegexOptions.IgnoreCase);
                sha1 = crypto.Groups[1].Value;
                md5 = crypto.Groups[3].Value;
                sha1 = sha1.Replace("\"", "");
                md5 = md5.Replace("\"", "");

                //
               
                int row = 0;
                while (regex.Success)
                {
                    row++;
                    vaccine = regex.Groups[1].Value;
                    version = regex.Groups[3].Value;
                    result = regex.Groups[4].Value;
                    update = regex.Groups[5].Value;
                    

                    result = result.Replace("\"", "");
                    update = update.Replace("\"", "");
                    version = version.Replace("\"", "");
                  

                    if (result == "null")
                        result = "이상없음";
                   
                   Excelcheck(row, vaccine, result, update, version,name,sha1,md5);
                    regex = regex.NextMatch();
                }


                ListViewItem newitem = new ListViewItem(name);
                listView1.Items.Add(newitem);
                
               
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
                endExcel();
                return;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if(dialog.ShowDialog()==DialogResult.OK)
            {
                string selected = dialog.SelectedPath;

                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(selected);
                System.IO.FileInfo[] fi = di.GetFiles();
                if(fi.Length==0)
                {
                    MessageBox.Show("없음");
                    return;
                }
                else
                {
               
                    try
                    {
                        startExcel();
                    for(int i=0;i<fi.Length;i++)
                    {
                        //s += fi[i].FullName.ToString();// +Environment.NewLine;
                        totalvirus(fi[i].FullName.ToString(),fi[i].Name.ToString());
                        Thread.Sleep(200);
                    }
                        endExcel();
                    }
                    catch (Exception error)
                    {
                        MessageBox.Show(error.ToString());
                        endExcel();
                        Application.Exit();
                    }
                }
            }
        }
    }
}

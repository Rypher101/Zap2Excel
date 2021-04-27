using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Zap2Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnHTML_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Select HTML file";
            openFileDialog1.Filter = "HTML file | *.html";
            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtHTML.Text = openFileDialog1.FileName;
            }
        }

        private void btnOutput_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.Description = "Select output location";
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtOutput.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            txtLog.Text = "";
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.OptionFixNestedTags = true;
            var objTable = new List<ScanTable>();
            try
            {
                AppendToLog("Begining process");
                AppendToLog("Reading HTML");
                if (!File.Exists(txtHTML.Text))
                {
                    throw new FileNotFoundException();
                }
                htmlDoc.Load(txtHTML.Text);

                if (htmlDoc.ParseErrors != null && htmlDoc.ParseErrors.Count() >0)
                {
                    foreach (var item in htmlDoc.ParseErrors)
                    {
                        AppendToLog("\t" + item.Reason, true);
                    }
                }
                AppendToLog("HTML read successfull");

                AppendToLog("Selecting tables");
                var tables = htmlDoc.DocumentNode.Descendants().Where(x=>x.HasClass("results")).ToList();
                AppendToLog("Selecting tables completed. Table count :" + tables.Count);

                AppendToLog("Creating objects");
                int i = 0;
                foreach (var item in tables)
                {
                    var tmpTable = new ScanTable();
                    var rows = item.SelectNodes(".//tr").ToList();
                    tmpTable.Severity = rows.First().SelectNodes(".//th").ElementAt(0).InnerText;
                    tmpTable.Vulnerability = rows.First().SelectNodes(".//th").ElementAt(1).InnerText;
                    rows.RemoveAt(0);
                    tmpTable.Description = rows.First().SelectNodes(".//td").ElementAt(1).InnerText;
                    rows.RemoveAt(0);

                    var urlFound = false;
                    var tmpURL = new string[2];
                    foreach (var item2 in rows)
                    {
                        if (item2.SelectSingleNode(".//td").GetAttributeValue("class",null) == "indent1")
                        {
                            tmpURL[0]= item2.SelectNodes(".//td").ElementAt(1).InnerText;
                            urlFound = true;
                        }
                        else if (urlFound)
                        {
                            tmpURL[1] = item2.SelectNodes(".//td").ElementAt(1).InnerText;
                            tmpTable.URLs.Add(tmpURL);
                            urlFound = false;
                        }

                        if (item2.SelectSingleNode(".//td").InnerText == "CWE Id")
                        {
                            tmpTable.CWE = int.Parse(item2.SelectNodes(".//td").ElementAt(1).InnerText);
                        }
                    }

                    i++;
                    objTable.Add(tmpTable);
                    AppendToLog("\t Tables " + i + " of " + tables.Count + " completed.");
                }
                AppendToLog("Creating objects completed.");

                AppendToLog("Creating excel file.");
                Excel.Application xlApp = new Excel.Application();
                if (xlApp == null)
                {
                    AppendToLog("Creating excel file failed. Please try again.", true);
                    return;
                }

                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(Type.Missing);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "Severity";
                xlWorkSheet.Cells[1, 2] = "Vulnerability";
                xlWorkSheet.Cells[1, 3] = "Desctiption";
                xlWorkSheet.Cells[1, 4] = "URL";
                xlWorkSheet.Cells[1, 5] = "Method";
                xlWorkSheet.Cells[1, 6] = "CWE ID";
                xlWorkSheet.Cells[1, 7] = "Comments";

                AppendToLog("\t Adding data to table");
                var row = 2;
                foreach (var item in objTable)
                {
                    xlWorkSheet.Cells[row,1] = item.Severity;
                    xlWorkSheet.Cells[row,2] = item.Vulnerability;
                    xlWorkSheet.Cells[row,3] = item.Description;
                    xlWorkSheet.Cells[row,6] = item.CWE;
                    foreach (var item2 in item.URLs)
                    {
                        xlWorkSheet.Cells[row,4] = item2[0];
                        xlWorkSheet.Cells[row,5] = item2[1];
                        row++;
                    }
                    
                }
                AppendToLog("\t Adding data to table complete");

                AppendToLog("\t Merging cells");
                row = 2;
                foreach (var item in objTable)
                {
                    var severityRange = xlWorkSheet.Range[xlWorkSheet.Cells[row, 1], xlWorkSheet.Cells[row + item.URLs.Count-1, 1]];
                    var vulnerabilityRange = xlWorkSheet.Range[xlWorkSheet.Cells[row, 2], xlWorkSheet.Cells[row + item.URLs.Count-1, 2]];
                    var descriptionRange = xlWorkSheet.Range[xlWorkSheet.Cells[row, 3], xlWorkSheet.Cells[row + item.URLs.Count-1, 3]];
                    var cweRange = xlWorkSheet.Range[xlWorkSheet.Cells[row, 6], xlWorkSheet.Cells[row + item.URLs.Count-1, 6]];

                    severityRange.Merge();
                    vulnerabilityRange.Merge();
                    descriptionRange.Merge();
                    cweRange.Merge();

                    row += item.URLs.Count;
                }
                AppendToLog("\t Merging cells complete");

                AppendToLog("\t Saving excel file");
                string fileName = txtOutput.Text +"\\"+ Path.GetFileNameWithoutExtension(txtHTML.Text) + ".xlsx";
                xlWorkBook.SaveAs(fileName);
                xlWorkBook.Close();
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
                AppendToLog("\t Saving excel file complete");
                AppendToLog("Creating excel file completed.");

            }
            catch (Exception ex)
            {
                AppendToLog("Error : " + ex.Message.ToString(), true);
            }
        }

        private void AppendToLog(string txt, bool isError = false)
        {
            if (isError)
            {
                txtLog.SelectionColor = Color.Red;
            }
            txtLog.AppendText(txt);
            txtLog.SelectionColor = txtLog.ForeColor;
            txtLog.AppendText(Environment.NewLine);
            txtLog.ScrollToCaret();
        }
    }
}

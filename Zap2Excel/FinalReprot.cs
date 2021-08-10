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
    public partial class FinalReprot : Form
    {
        public FinalReprot()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Select HTML file";
            openFileDialog1.Filter = "HTML file | *.html";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtHTML.Text = openFileDialog1.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.Description = "Select output location";
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtOutput.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void button3_Click(object sender, EventArgs e)
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

                if (htmlDoc.ParseErrors != null && htmlDoc.ParseErrors.Count() > 0)
                {
                    foreach (var item in htmlDoc.ParseErrors)
                    {
                        AppendToLog("\t" + item.Reason, true);
                    }
                }
                AppendToLog("HTML read successfull");

                AppendToLog("Selecting tables");
                var tables = htmlDoc.DocumentNode.Descendants().Where(x => x.HasClass("results")).ToList();
                AppendToLog("Selecting tables completed. Table count :" + tables.Count);

                AppendToLog("Creating objects");
                int i = 0;
                foreach (var item in tables)
                {
                    var tmpTable = new ScanTable();
                    var rows = item.SelectNodes(".//tr").ToList();
                    var tmpSev = rows.First().SelectNodes(".//th").ElementAt(0).InnerText.Split(' ');
                    if (tmpSev[0] == "High")
                    {
                        tmpTable.Severity = "High";
                    }
                    else if (tmpSev[0] == "Medium")
                    {
                        tmpTable.Severity = "Medium";
                    }
                    else if (tmpSev[0] == "Low")
                    {
                        tmpTable.Severity = "Low";
                    }
                    else if (tmpSev[0] == "Informational")
                    {
                        tmpTable.Severity = "Informational";
                    }

                    rows.RemoveAt(0);
                    tmpTable.Description = rows.First().SelectNodes(".//td").ElementAt(1).InnerText;
                    tmpTable.FinalURLs = "";
                    rows.RemoveAt(0);

                    foreach (var item2 in rows)
                    {
                        if (item2.SelectSingleNode(".//td").GetAttributeValue("class", null) == "indent1")
                        {
                            tmpTable.FinalURLs = tmpTable.FinalURLs + item2.SelectNodes(".//td").ElementAt(1).InnerText + Environment.NewLine;
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
                var xlsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Zap to Excel", "FinalReport.xls");

                xlWorkBook = xlApp.Workbooks.Open(xlsPath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0); ;
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                AppendToLog("\t Adding data to table");
                var row = 3;
                foreach (var item in objTable)
                {
                    xlWorkSheet.Cells[row, 1] = row - 2;

                    xlWorkSheet.Cells[row, 2] = item.Severity;
                    xlWorkSheet.Cells[row, 2].Font.Color = ColorTranslator.ToOle(Color.White);
                    switch (item.Severity)
                    {
                        case "High":
                            xlWorkSheet.Cells[row, 2].Interior.Color = ColorTranslator.ToOle(Color.Red);
                            break;
                        case "Medium":
                            xlWorkSheet.Cells[row, 2].Interior.Color = ColorTranslator.ToOle(Color.Orange);
                            break;
                        case "Low":
                            xlWorkSheet.Cells[row, 2].Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                            xlWorkSheet.Cells[row, 2].Font.Color = ColorTranslator.ToOle(Color.Black);
                            break;
                        case "Informational":
                            xlWorkSheet.Cells[row, 2].Interior.Color = ColorTranslator.ToOle(Color.Blue);
                            break;
                        default:
                            break;
                    }

                    xlWorkSheet.Cells[row, 4] = item.Description;
                    xlWorkSheet.Cells[row, 5] = txtJira.Text;
                    xlWorkSheet.Cells[row, 8] = item.FinalURLs;
                    row++;
                }

                xlWorkSheet.Range[xlWorkSheet.Cells[3, 2], xlWorkSheet.Cells[row - 1, 11]].Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                xlWorkSheet.Range[xlWorkSheet.Cells[3, 1], xlWorkSheet.Cells[row - 1, 1]].Style.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Range[xlWorkSheet.Cells[3, 5], xlWorkSheet.Cells[row - 1, 5]].Style.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Range[xlWorkSheet.Cells[3, 1], xlWorkSheet.Cells[row - 1, 11]].Font.Size = 16;
                xlWorkSheet.Range[xlWorkSheet.Cells[3, 1], xlWorkSheet.Cells[row - 1, 11]].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                AppendToLog("\t Adding data to table complete");

                AppendToLog("\t Saving excel file");
                string fileName = txtOutput.Text + "\\" + Path.GetFileNameWithoutExtension(txtHTML.Text) + ".xls";
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

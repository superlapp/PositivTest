using System.Collections.Generic;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using System;

namespace PositivTest
{
    class clsWorker
    {
        string page = @"http://www.moyareklama.by/%D0%93%D0%BE%D0%BC%D0%B5%D0%BB%D1%8C/%D0%BA%D0%B2%D0%B0%D1%80%D1%82%D0%B8%D1%80%D1%8B_%D0%BF%D1%80%D0%BE%D0%B4%D0%B0%D0%B6%D0%B0/%D0%B2%D1%81%D0%B5/8/";

        public List<string> Parse()
        {
            List<string> result = new List<string>();
            //
            try
            {
                for (int i = 1; i <= 2; i++)
                {
                    var doc = new HtmlWeb().Load(page + i.ToString() + "/");
                    var nodeCol = doc.DocumentNode.SelectNodes("//div[@class='sa_header ']");
                    //
                    if (nodeCol != null)
                    {
                        foreach (HtmlNode hNode in nodeCol)
                        {
                            var hcolls = hNode.ChildNodes;
                            string r = "";
                            foreach (HtmlNode hcoll in hcolls)
                            {
                                if (hcoll.NodeType == HtmlNodeType.Element)
                                {
                                    string str = hcoll.InnerText.Remove(0, 2).Trim();
                                    r = r + str + ";";
                                }
                            }
                            if (r.Length != 0)
                            {
                                result.Add(r);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            //
            return result;
        }

        public void WriteToXLS(List<string> list)
        {
            var excel = new Microsoft.Office.Interop.Excel.Application();
            //
            if (excel == null)
            {
                System.Windows.Forms.MessageBox.Show("MS Excel не установлен");
            }
            else
            {
                try
                {
                    Workbook xBook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                    Worksheet xSheet = (Worksheet)xBook.Worksheets[1];
                    //
                    xSheet.Cells[1, 1] = "Название";
                    xSheet.Cells[1, 1].EntireRow.Font.Bold = true;
                    xSheet.Cells[1, 1].EntireRow.Font.Size = 12;
                    xSheet.Cells[1, 2] = "Цена";
                    xSheet.Cells[1, 2].EntireRow.Font.Bold = true;
                    xSheet.Cells[1, 2].EntireRow.Font.Size = 12;
                    //
                    for (int i = 0; i < list.Count; i++)
                    {
                        string[] t = list[i].Split(';');
                        //
                        xSheet.Cells[i + 2, 1] = t[1];
                        xSheet.Cells[i + 2, 2] = t[0];
                    }
                    //
                    xSheet.Columns.AutoFit();
                    //
                    string[] fileCount = System.IO.Directory.GetFiles(System.Windows.Forms.Application.StartupPath, "*.xlsx");
                    string fileName =
                        System.Windows.Forms.Application.StartupPath +
                        "\\Выборка " + (fileCount.Length + 1).ToString() + ".xlsx";
                    xBook.SaveAs(fileName);
                    //
                    excel.Quit();
                    //
                    System.Diagnostics.Process.Start(fileName);
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
            }
        }
    }
}

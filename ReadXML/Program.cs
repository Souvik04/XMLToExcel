using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Drawing;

namespace ReadXML
{
    class Program
    {
        private List<string> _lsColumnNames = new List<string>();
        static void Main(string[] args)
        {
            Program obj = new Program();
            string xmlFilePath = @ConfigurationManager.AppSettings["XMLFilePath"].ToString();

            #region XML Node or Attributes to fetch

            List<string> lsMessageTypeName = obj.ReadNodeAttribute(xmlFilePath, ConfigurationManager.AppSettings["Message.TypeName"]);
            List<string> lsMessageCategory = obj.ReadNodeAttribute(xmlFilePath, ConfigurationManager.AppSettings["Message.Category"]);
            List<string> lsIssueLevel = obj.ReadNodeAttribute(xmlFilePath, ConfigurationManager.AppSettings["Issue.Level"]);
            List<string> lsIssueValue = obj.ReadNodeValue(xmlFilePath, ConfigurationManager.AppSettings["Issue"]);

            //add your own List<string> and specify the XMLNodeName and Attribute to fetch the matched Node and Attribute values from XML
            //or specify the  NodeName to fetch the matched Node values from XML

            #endregion XML Node or Attributes to fetch

            string excelFilePath = @ConfigurationManager.AppSettings["ExcelFilePath"].ToString();

            obj.WriteToExcel(excelFilePath, lsMessageTypeName, lsMessageCategory, lsIssueLevel, lsIssueValue);
        }

        private List<string> ReadNodeAttribute(string xmlFilePath, string nodeAndAttribute)
        {
            string nodeName = nodeAndAttribute.Split('.')[0];
            string attribute = nodeAndAttribute.Split('.')[1];
            _lsColumnNames.Add(nodeAndAttribute);

            List<string> lsValues = new List<string>();
            XmlDocument doc = new XmlDocument();
            XmlNodeList nodeList = null;
            XmlElement root = null;
            try
            {
                doc.Load(xmlFilePath);
                root = doc.DocumentElement;
                nodeList = root.SelectNodes("//" + nodeName);

                lsValues = new List<string>();

                foreach (XmlNode node in nodeList)
                {
                    lsValues.Add(node.Attributes.GetNamedItem(attribute).Value);
                }
            }
            catch (Exception ex)
            {

            }
            return lsValues;
        }

        private List<string> ReadNodeValue(string xmlFilePath, string nodeName)
        {
            _lsColumnNames.Add(nodeName);
            List<string> lsValues = new List<string>();
            XmlDocument doc = new XmlDocument();
            XmlNodeList nodeList = null;
            XmlElement root = null;
            try
            {
                doc.Load(xmlFilePath);
                root = doc.DocumentElement;
                nodeList = root.SelectNodes("//" + nodeName);

                lsValues = new List<string>();

                foreach (XmlNode node in nodeList)
                {
                    lsValues.Add(node.InnerText);
                }
            }
            catch (Exception ex)
            {

            }
            return lsValues;
        }

        private void WriteToExcel(string excelFilePath, params List<string>[] listOfColumns)
        {
            Application excelApp = new Application();
            Worksheet ws = new Worksheet();
            Workbook wb = null;
            try
            {
                wb = excelApp.Workbooks.Open(excelFilePath, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing);
                ws = (Worksheet)wb.Sheets[1];

                //ws = (Worksheet)sheets.get_Item("Sheet1");

                ws.Cells.Clear();
                ws.Cells.ClearFormats();
                ws.Cells.ClearOutline();

                for (int i = 0; i < listOfColumns.Count(); i++)
                {
                    //ws.Cells.set_Item(1, i, _lsColumnNames[i]);
                    for (int j = 0; j < listOfColumns[i].Count; j++)
                    {
                        ws.Cells.set_Item(j + 1, i + 1, listOfColumns[i][j].ToString());

                        //set legend colors
                        switch (listOfColumns[i][j])
                        {
                            case "CriticalError":
                                ws.Rows[j + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                break;
                            case "Error":
                                ws.Rows[j + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Pink);
                                break;
                            case "CriticalWarning":
                                ws.Rows[j + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                                break;
                            case "Warning":
                                ws.Rows[j + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                break;
                                //add your own background cell formatting logic here
                            default:
                                break;
                        }
                    }
                }

                Range row = (Range)ws.Rows[1];
                row.Insert();

                for (int i = 0; i < _lsColumnNames.Count; i++)
                {
                    ws.Cells.set_Item(1, i + 1, _lsColumnNames[i]);
                }

                ws.Rows[1].EntireRow.Font.Bold = true;
                ws.UsedRange.Rows.Borders.LineStyle = XlLineStyle.xlContinuous;

                //ws.UsedRange.Font.Background.Borders.Color = System.Drawing.Color.Black.ToArgb();

                excelApp.DisplayAlerts = false;
                ws.SaveAs(excelFilePath);
                wb.SaveAs(Type.Missing, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
                    Type.Missing, Type.Missing,
                    true, false,
                    XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges,
                    Type.Missing, Type.Missing);

                //ws.SaveAs(excelFilePath);
            }
            catch (Exception ex)
            {

            }
            finally
            {
                wb.Close();
                excelApp.Quit();
                Marshal.ReleaseComObject(ws);
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(excelApp);
                ws = null;
                wb = null;
                excelApp = null;
            }
        }
    }
}

using Microsoft.Office.Interop.Word;
using Excel=Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using _Application = Microsoft.Office.Interop.Word._Application;
using System.Drawing;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace AITExportCsv
{
    class ExportFile
    {
        /// <summary>
        /// Export file csv to folder
        /// </summary>
        /// <param name="folder">export file to folder</param>
        /// <param name="nameFile">name file csv</param>
        /// <param name="listData">list data export</param>
        /// <param name="header">header csv (format header1,header2,...) ,if header=="" set property for header, if header==null not set header</param>
        /// <param name="squotes">add squotes for csv</param>
        /// <returns>true if export successful,false if export fail</returns>
        public Boolean Export<T>(string folder, string nameFile, List<T> listData, string typeFile, string header = null,Boolean squotes=false) where T : new()
        {
                switch (typeFile)
                {
                    case "CSV":
                        ExportCSV(folder, nameFile, listData, header, squotes);
                        break;
                    case "Excel":
                        ExportExcel(folder, nameFile, listData, header);
                        break;
                    case "PDF":
                        ExportPDF(folder, nameFile, listData, header);
                        break;
                    case "Word":
                        ExportWord(folder, nameFile, listData, header);
                        break;
                    default:
                        return false;
                }
                return true;
}
        /// <summary>
        /// export csv
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="folder">link path</param>
        /// <param name="nameFile">file name</param>
        /// <param name="listData">list data export</param>
        /// <param name="header">header csv (format header1,header2,...) ,if header=="" set property for header, if header==null not set header</param>
        /// <param name="squotes">squotes for csv</param>
        /// <returns>return true if export success</returns>
        private void ExportCSV<T>(string folder, string nameFile, List<T> listData, string header = null,Boolean squotes=false) where T : new()
        {
                var obj = new T();
                PropertyInfo[] properties = obj.GetType().GetProperties();
                string csv = string.Empty;
                if (!string.IsNullOrEmpty(header))
                {
                    //Build the CSV file data as a Comma separated string.
                    csv += header;
                    //Add new line.
                    csv += "\r\n";
                }
                if (listData != null)
                {
                    if (header.Equals(""))
                    {
                        foreach (PropertyInfo propti in properties)
                        {
                            //Add the Header.
                            if (squotes == true)
                            {
                                csv +='"'+ propti.Name.ToString() + "\","; 
                            }
                            else
                            {
                                csv += propti.Name.ToString() + ",";
                            }
                        }
                        csv = csv.Remove(csv.Length - 1);
                        //Add new line.
                        csv += "\r\n";
                    }
                    foreach (var data in listData)
                    {
                        foreach (PropertyInfo prop in properties)
                        {
                            //Add the Data rows.
                            var type = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                            if (type.Name=="DateTime"&& prop.GetValue(data, null)!=null)
                            {
                                var time=(DateTime)prop.GetValue(data, null);
                                if (squotes == true)
                                {
                                    csv += '"'+ GetCustomFormatString(time).ToString() + "\",";
                                }
                                else
                                {
                                    csv += GetCustomFormatString(time).ToString() + ",";
                                }
                            }
                            else
                            {
                                if (squotes == true)
                                {
                                    csv += prop.GetValue(data, null) == null ? "\"\"," : '"'+prop.GetValue(data, null).ToString() + "\",";
                                }
                                else
                                {
                                    csv += prop.GetValue(data, null) == null ? "," :prop.GetValue(data, null).ToString() + ",";
                                }
                            }
                        }
                        csv = csv.Remove(csv.Length - 1);
                        //Add new line.
                        csv += "\r\n";
                    }
                    csv = csv.Remove(csv.Length - 2);
                }
                //Exporting to CSV.
                string folderPath = folder;
                File.WriteAllText(folderPath +"\\"+ nameFile + ".csv", csv);
        }
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
        /// <summary>
        /// export excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="folderPath">path file</param>
        /// <param name="nameFile">file name</param>
        /// <param name="listData">list data export</param>
        /// <param name="header">header csv (format header1,header2,...) ,if header=="" set property for header, if header==null not set header</param>
        /// <returns>return true if export success</returns>
        private void ExportExcel<T>(string folderPath, string nameFile, List<T> listData, string header = "") where T : new()
        {

            var obj = new T();
            PropertyInfo[] properties = obj.GetType().GetProperties();
            // create a excel app along side with workbook and worksheet and give a name to it  
            Excel.Application excelApp = new Excel.Application();
            int id;
            // Find the Process Id
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            Process excelProcess = Process.GetProcessById(id);
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Add();
            Excel._Worksheet xlWorksheet = excelWorkBook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            //Add a new worksheet to workbook with the Datatable name  
            Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
            excelWorkSheet.Name = nameFile;
            if (!header.Equals(""))
            {
                var headers = header.Split(',');
                // add all the columns
                var headerTable = 1;
                foreach (var head in headers)
                {
                    excelWorkSheet.Cells[1, headerTable].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    excelWorkSheet.Cells[1, headerTable++] = head;

                    }
                }
            if (listData != null)
            {
                if (header.Equals(""))
                {
                    // add all the columns
                    var headerTable = 1;
                    foreach (PropertyInfo pi in properties)
                    {
                        excelWorkSheet.Cells[1, headerTable].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        excelWorkSheet.Cells[1, headerTable++] = pi.Name;
                    }
                }
                var j=0;
                foreach (var data in listData)
                {
                    var k = 0;
                    foreach (PropertyInfo pi in properties)
                    {
                            var dataString="";
                            var type = Nullable.GetUnderlyingType(pi.PropertyType) ?? pi.PropertyType;
                            if (type.Name == "DateTime" && pi.GetValue(data, null) != null)
                            {
                                var time = (DateTime)pi.GetValue(data, null);
                                dataString = GetCustomFormatString(time).ToString();
                            }
                            else
                                dataString = pi.GetValue(data) == null ? "" : pi.GetValue(data).ToString();
                            excelWorkSheet.Cells[j + 2, k + 1] = dataString;
                            excelWorkSheet.Cells[j + 2, k + 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        k++;
                    }
                    j++;
                }
            }
            var path = folderPath + "\\" + nameFile + ".xlsx";
            try
            {
                excelWorkBook.SaveAs(path.Replace("\\\\", "\\")); // -> this will do the custom  
            }
            catch
            {
                excelProcess.Kill();
            }
            excelWorkBook.Close();
            excelApp.Quit();
            excelProcess.Kill();
        }
        /// <summary>
        /// export pdf
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="folderPath">path file</param>
        /// <param name="nameFile">file name</param>
        /// <param name="listData">list data export</param>
        /// <param name="header">header csv (format header1,header2,...) ,if header=="" set property for header, if header==null not set header</param>
        /// <returns>return true if export success</returns>
        private void ExportPDF<T>(string folderPath, string nameFile, List<T> listData, string header = "") where T : new()
        {
                var obj = new T();
                PropertyInfo[] properties = obj.GetType().GetProperties();
                //Table start.
                string html = "<table style='border: 1px solid;width: 100%;border-collapse: collapse;'>";
                if (!header.Equals(""))
                {
                    var dataHeaders = header.Split(',');
                    //Adding HeaderRow.
                    html += "<tr>";
                    foreach (var headerData in dataHeaders)
                    {
                        html += "<th style='border: 1px solid;text-align: left;'>" + headerData + "</th>";
                    }
                    html += "</tr>";
                }
                if (listData != null)
                {
                    if (header.Equals(""))
                    {
                        //Adding HeaderRow.
                        html += "<tr>";
                        foreach (PropertyInfo pi in properties)
                        {
                            html += "<th style='border: 1px solid;text-align: left;'>" + pi.Name + "</th>";
                        }
                        html += "</tr>";
                    }
                    //Adding DataRow.
                    foreach (var data in listData)
                    {
                        html += "<tr>";
                        foreach (PropertyInfo pi in properties)
                        {
                            var dataString = "";
                            var type = Nullable.GetUnderlyingType(pi.PropertyType) ?? pi.PropertyType;
                            if (type.Name == "DateTime" && pi.GetValue(data, null) != null)
                            {
                                var time = (DateTime)pi.GetValue(data, null);
                                dataString = GetCustomFormatString(time).ToString();
                            }
                            else
                                dataString = pi.GetValue(data) == null ? "" : pi.GetValue(data).ToString();
                            html += "<td style='border: 1px solid;'>" + dataString + "</td>";
                        }
                        html += "</tr>";
                    }
                    //Table end.
                    html += "</table>";
                }
                //Save the HTML string as HTML File.
                string htmlFilePath = folderPath + nameFile + ".htm";
                File.WriteAllText(htmlFilePath, html);

                //Convert the HTML File to Word document.
                _Application pdf = new Application();
                _Document dwcumentPdf = pdf.Documents.Open(FileName: htmlFilePath, ReadOnly: false);
                dwcumentPdf.SaveAs(FileName: folderPath + "\\" + nameFile + ".pdf", FileFormat: WdSaveFormat.wdFormatPDF);
                ((_Document)dwcumentPdf).Close();
                ((_Application)pdf).Quit();

                //Delete the HTML File.
                File.Delete(htmlFilePath);
        }
        /// <summary>
        /// export pdf
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="folderPath">path file</param>
        /// <param name="nameFile">file name</param>
        /// <param name="listData">list data export</param>
        /// <param name="header">header csv (format header1,header2,...) ,if header=="" set property for header, if header==null not set header</param>
        /// <returns>return true if export success</returns>
        private void ExportWord<T>(string folderPath, string nameFile, List<T> listData, string header = "") where T : new()
        {
                var obj = new T();
                PropertyInfo[] properties = obj.GetType().GetProperties();
                //Table start.
                string html = "<table style='border: 1px solid;width: 100%;border-collapse: collapse;text-align: left;'>";
                if (!header.Equals(""))
                {
                    var dataHeaders = header.Split(',');
                    //Adding HeaderRow.
                    html += "<tr>";
                    foreach (var headerData in dataHeaders)
                    {
                        html += "<th style='border: 1px solid;text-align: left;'>" + headerData + "</th>";
                    }
                    html += "</tr>";
                }
                if (listData != null)
                {
                    if (header.Equals(""))
                    {
                        //Adding HeaderRow.
                        html += "<tr>";
                        foreach (PropertyInfo pi in properties)
                        {
                            html += "<th style='border: 1px solid;text-align: left;'>" + pi.Name + "</th>";
                        }
                        html += "</tr>";
                    }
                    //Adding DataRow.
                    foreach (var data in listData)
                    {
                        html += "<tr>";
                        foreach (PropertyInfo pi in properties)
                        {
                            var dataString = "";
                            var type = Nullable.GetUnderlyingType(pi.PropertyType) ?? pi.PropertyType;
                            if (type.Name == "DateTime" && pi.GetValue(data, null) != null)
                            {
                                var time = (DateTime)pi.GetValue(data, null);
                                dataString = GetCustomFormatString(time).ToString();
                            }
                            else
                                dataString = pi.GetValue(data) == null ? "" : pi.GetValue(data).ToString();
                            html += "<td style='border: 1px solid;'>" + dataString + "</td>";
                        }
                        html += "</tr>";
                    }
                    //Table end.
                    html += "</table>";
                }
                //Save the HTML string as HTML File.
                string htmlFilePath = folderPath + nameFile + ".htm";
                File.WriteAllText(htmlFilePath, html);

                //Convert the HTML File to Word document.
                _Application word = new Application();
                _Document wordDoc = word.Documents.Open(FileName: htmlFilePath, ReadOnly: false);
                wordDoc.SaveAs(FileName: folderPath + "\\" + nameFile + ".doc", FileFormat: WdSaveFormat.wdFormatRTF);
                ((_Document)wordDoc).Close();
                ((_Application)word).Quit();

                //Delete the HTML File.
                File.Delete(htmlFilePath);
        }
        private string GetCustomFormatString(DateTime input, bool excludeTimeIfZero = true)
        {
            return input.TimeOfDay == TimeSpan.Zero && excludeTimeIfZero
                ? input.ToString("yyyy/MM/dd")
                : input.ToString("yyyy/MM/dd HH:mm:ss");
        }
    }
}

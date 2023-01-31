using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace AITImportCSV
{
    class ImportCSV
    {
        /// <summary>
        /// import a file csv insert DB
        /// </summary>
        /// <typeparam name="T">model use</typeparam>
        /// <param name="quotes">true if have quotes, else false</param>
        /// <param name="columnSeparator">separator character between columns</param>
        /// <param name="header">save header or not</param>
        /// <param name="encoding">endcode type</param>
        /// <returns>list model</returns>
        public List<T> Import<T>(Boolean quotes=true,char columnSeparator=',', Boolean header =false,Encoding encoding=null) where T:new()
        {
            var filePath = string.Empty;
            var listdata = new List<T>();
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "csv File (.csv)|*.csv";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                filePath = openFileDialog.FileName;

                //Read the contents of the file into a stream
                var fileStream = openFileDialog.OpenFile();

                using (StreamReader reader = new StreamReader(fileStream, encoding == null ? Encoding.Default : encoding))
                {
                    var flagHeader = 0;
                    while (!reader.EndOfStream)
                    {
                        var obj = new T();
                        var line = reader.ReadLine();
                        var values = line.Split(columnSeparator);
                        if (quotes)
                        {
                            for (var i = 0; i < values.Length; i++)
                            {
                                values[i] = values[i].Trim('"');

                            }
                        }
                        if (flagHeader == 1 && header == false)
                        {
                            PropertyInfo[] properties = obj.GetType().GetProperties();
                            var i = 0;
                                foreach (PropertyInfo pi in properties)
                                {
                                    if (i + 1 <= values.Length)
                                    {
                                        pi.SetValue(obj, values[i], null);
                                    }else
                                        pi.SetValue(obj, "", null);
                                    i++;
                                }
                            listdata.Add(obj);
                        }
                        if(header == true)
                        {
                            PropertyInfo[] properties = obj.GetType().GetProperties();
                            var i = 0;
                                foreach (PropertyInfo pi in properties)
                                {
                                    if (i + 1 <= values.Length)
                                    {
                                        pi.SetValue(obj, values[i], null);
                                    }
                                    else
                                        pi.SetValue(obj, "", null);
                                    i++;
                                }
                            listdata.Add(obj);
                        }
                        flagHeader = 1;
                    }
                    string allText = File.ReadAllText(filePath);
                    var objData = new T();
                    PropertyInfo[] prope = objData.GetType().GetProperties();
                    foreach (PropertyInfo pi in prope)
                    {
                            pi.SetValue(objData, "", null);
                    }
                    if (allText != "" && ((allText.Length >= 2 && allText.Substring(allText.Length - 2) == "\r\n") || (allText.Length >= 1 && allText.Substring(allText.Length - 1) == "\n") || (allText.Length >= 1 && allText.Substring(allText.Length - 1) == "\r")))
                        listdata.Add(objData);
                }

            }
            return listdata;
        }
        /// <summary>
        /// import a file csv show table
        /// </summary>
        /// <typeparam name="T">model use</typeparam>
        /// <param name="quotes">true if have quotes, else false</param>
        /// <param name="columnSeparator">separator character between columns</param>
        /// <param name="header">save header or not</param>
        /// <param name="encoding">endcode type</param>
        /// <returns>list model</returns>
        public List<string[]> Import(Boolean quotes = true, char columnSeparator = ',', Boolean header = false, Encoding encoding=null)
        {
            var filePath = string.Empty;
            var listdata = new List<string[]>();
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "csv File (.csv)|*.csv";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                filePath = openFileDialog.FileName;

                //Read the contents of the file into a stream
                var fileStream = openFileDialog.OpenFile();

                using (StreamReader reader = new StreamReader(fileStream, encoding==null?Encoding.Default:encoding))
                {
                var flagHeader = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(columnSeparator);
                        if (quotes)
                        {
                            for(var i=0;i<values.Length; i++)
                            {
                                values[i] = values[i].Trim('"');

                            }
                        }
                        if (flagHeader == 1 && header == false)
                        {
                            listdata.Add(values);
                        }
                        if (header == true)
                        {
                            listdata.Add(values);
                        }
                        flagHeader = 1;
                    }
                    string allText = File.ReadAllText(filePath);
                    if (allText!=""&&((allText.Length>=2 && allText.Substring(allText.Length - 2) == "\r\n") || (allText.Length >= 1 && allText.Substring(allText.Length - 1) == "\n") || (allText.Length >= 1 && allText.Substring(allText.Length - 1) == "\r")))
                        listdata.Add("".Split());
                }
            }
            return listdata;
        }
    }
}

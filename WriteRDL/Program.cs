using System;
using System.Collections.Generic;
using System.IO;
using WriteRDL.RS2005;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace WriteRDL
{
    class Program
    {
        static void Main(string[] args)
        {
            string report = @"/<reportFolder>/<reportName>";

            Dictionary<string, string> dict = new Dictionary<string, string>();
            dict.Add("Year", "2017");
            dict.Add("StartMonth", "1");
            dict.Add("EndMonth", "5");

            string fileName = @"<location>\testreport";

            WriteRDLToFile.SaveRDL(report, "PDF", dict, fileName);
        }

        public enum SSRSServer { Old2012, New2016 };

        public class WriteRDLToFile
        {
            //formatType must be: EXCEL, PDF, WORD, CSV, XML, RTF
            public static void SaveRDL(string reportName, string formatType, Dictionary<string, string> dictValues, string _fileName, SSRSServer server = SSRSServer.New2016)
            {
                // Authenticate to the Web service using Windows credentials
                var netCredentials = new System.Net.NetworkCredential("<windowsUserName>", "<windowsUserPassword>", "<Domain>");

                // Create a new proxy to the web service

                ReportingService2005 rs = new ReportingService2005()
                {
                    Credentials = netCredentials
                };
                RE2005.ReportExecutionService rsExec = new RE2005.ReportExecutionService()
                {
                    Credentials = netCredentials
                };

                if (server == SSRSServer.New2016)
                {
                    rs.Url = $"http://<server>/reportservice2005.asmx";
                    rsExec.Url = $"http://<server>/reportexecution2005.asmx";
                }
                else
                {
                    rs.Url = $"http://<server>/reportservice2005.asmx";
                    rsExec.Url = $"http://<server>/reportexecution2005.asmx";
                }

                string historyId = null;
                bool forRendering = false;
                ParameterValue[] values = null;
                DataSourceCredentials[] credentials = null;
                ReportParameter[] _parameters = null;
                byte[] results;

                try
                {
                    _parameters = rs.GetReportParameters(reportName, historyId, forRendering, values, credentials);

                    if (_parameters != null)
                    {
                        foreach (ReportParameter rp in _parameters)
                        {
                            Console.WriteLine("Name: {0}", rp.Name);
                        }
                    }
                    RE2005.ParameterValue[] parameters = new RE2005.ParameterValue[dictValues.Count];

                    int count = 0;
                    foreach (var item in dictValues)
                    {
                        parameters[count] = new RE2005.ParameterValue();
                        parameters[count].Label = item.Key;
                        parameters[count].Name = item.Key;
                        parameters[count++].Value = item.Value;
                    }

                    RE2005.ExecutionInfo rpt = rsExec.LoadReport(reportName, null);
                    rsExec.SetExecutionParameters(parameters, "en-us");

                    //Render variables
                    string deviceInfo = null;
                    string encoding = String.Empty;
                    string mimeType = String.Empty;
                    string extension = String.Empty;
                    RE2005.Warning[] warnings = null;
                    string[] streamIDs = null;

                    string formatType2Pass;
                    switch (formatType)
                    {
                        case "RTF":
                            formatType2Pass = "WORD";  //And we convert afterwards
                            break;
                        case "DOCX":
                            formatType2Pass = "WORDOPENXML";
                            break;
                        default:
                            formatType2Pass = formatType;
                            break;
                    }

                    results = rsExec.Render(
                        formatType2Pass,
                        deviceInfo,
                        out extension,
                        out encoding,
                        out mimeType,
                        out warnings,
                        out streamIDs);

                    string fileNameExt;
                    switch (formatType)
                    {
                        case "EXCEL":
                            fileNameExt = ".XLSX";
                            break;
                        case "WORDOPENXML":
                            fileNameExt = ".DOCX";
                            break;
                        case "WORD":
                            fileNameExt = ".DOC";
                            break;
                        case "RTF":
                            fileNameExt = ".DOC";  //And we convert afterwards
                            break;
                        case "PDF":
                        default:
                            fileNameExt = "." + formatType;
                            break;
                    }

                    string fileName = (Path.GetExtension(_fileName) == "") ? _fileName + fileNameExt : _fileName;

                    using (FileStream stream = File.OpenWrite(fileName))
                    {
                        stream.Write(results, 0, results.Length);
                        stream.Dispose();
                    }
                    if (formatType == "RTF")
                    {
                        string fileNameRTF = _fileName + ".rtf";
                        WordToRtf(fileName, fileNameRTF);
                    }
                    File.SetAttributes(fileName, FileAttributes.Normal);
                    rs.Dispose();
                    rsExec.Dispose();
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show(e.Message);
                }
            }

            public static void WordToRtf(string infile, string outfile)
            {
                //We need try/catch to make sure we're not leaving open Word processes on the user's computer
                object missing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                try
                {
                    word.Visible = false;
                    Document doc = word.Documents.Open((Object)infile);
                    doc.Activate();
                    object fileFormat = WdSaveFormat.wdFormatRTF;

                    doc.SaveAs(outfile, fileFormat);

                    object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                    ((_Document)doc).Close(saveChanges);
                    doc = null;
                }
                catch
                {
                    //This will now go to finally to close the Word App
                }
                finally
                {
                    ((Microsoft.Office.Interop.Word._Application)word).Quit(ref missing, ref missing, ref missing);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(word);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(word);
                }
            }

            public static byte[] SaveRDLtoFileStream(string reportName, string formatType, Dictionary<string, string> dictValues)
            {
                // Create a new proxy to the web service
                ReportingService2005 rs = new ReportingService2005();
                RE2005.ReportExecutionService rsExec = new RE2005.ReportExecutionService();

                // Authenticate to the Web service using Windows credentials
                rs.Credentials = System.Net.CredentialCache.DefaultCredentials;
                rsExec.Credentials = System.Net.CredentialCache.DefaultCredentials;

                rs.Url = "http://<server>/reportservice2005.asmx";
                rsExec.Url = "http://<server>/reportexecution2005.asmx";

                string historyId = null;
                bool forRendering = false;
                ParameterValue[] values = null;
                DataSourceCredentials[] credentials = null;
                ReportParameter[] _parameters = null;
                byte[] results = new byte[64 * 1024];

                try
                {
                    _parameters = rs.GetReportParameters(reportName, historyId, forRendering, values, credentials);

                    if (_parameters != null)
                    {
                        foreach (ReportParameter rp in _parameters)
                        {
                            Console.WriteLine("Name: {0}", rp.Name);
                        }
                    }
                    RE2005.ParameterValue[] parameters = new RE2005.ParameterValue[dictValues.Count];

                    int count = 0;
                    foreach (var item in dictValues)
                    {
                        parameters[count] = new RE2005.ParameterValue();
                        parameters[count].Label = item.Key;
                        parameters[count].Name = item.Key;
                        parameters[count++].Value = item.Value;
                    }

                    RE2005.ExecutionInfo rpt = rsExec.LoadReport(reportName, null);
                    rsExec.SetExecutionParameters(parameters, "en-us");

                    //Render variables
                    string deviceInfo = null;
                    string encoding = String.Empty;
                    string mimeType = String.Empty;
                    string extension = String.Empty;
                    RE2005.Warning[] warnings = null;
                    string[] streamIDs = null;

                    results = rsExec.Render(
                        formatType,
                        deviceInfo,
                        out extension,
                        out encoding,
                        out mimeType,
                        out warnings,
                        out streamIDs);
                    return results;
                }
                catch (Exception e)
                {
                    //Console.WriteLine(e.Message);
                }
                return results;
            }
        }
    }
}

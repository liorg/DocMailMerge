using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using System.Net;
using System.Configuration;
using System.Diagnostics;

namespace Guardian.ELAL.Documents.MailMerge
{
    public class MailMergeOpenXml
    {
        const string AttributeValue = "w:val";
        string ConnectionString;
        Action<string, System.Diagnostics.EventLogEntryType> _log;

        public MailMergeOpenXml(Action<string, EventLogEntryType> log)
        {
            _log = log;
            ConnectionString =ConfigurationManager.ConnectionStrings["SqlConnectionString"].ConnectionString;
        }

        protected virtual System.Xml.XmlNamespaceManager PolpuateXmlNamespaceManager(XmlDocument xdoc)
        {
            var xmlnsManager = new System.Xml.XmlNamespaceManager(xdoc.NameTable);
            //Add the namespaces used in books.xml to the XmlNamespaceManager.
            xmlnsManager.AddNamespace("e", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            xmlnsManager.AddNamespace("pkg", "http://schemas.microsoft.com/office/2006/xmlPackage");
            xmlnsManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            return xmlnsManager;
        }

        protected byte[] GetBuffer(string path)
        {
            WebClient client = new WebClient();

            string userName = ConfigurationManager.AppSettings["userName"];
            if (String.IsNullOrEmpty(userName))
                throw new Exception("userName is null see config file");

            string domain = ConfigurationManager.AppSettings["domain"];
            if (String.IsNullOrEmpty(domain))
                throw new Exception("domain is null see config file");

            string passWord = ConfigurationManager.AppSettings["passWord"];
            if (String.IsNullOrEmpty(passWord))
                throw new Exception("passWord is null see config file");

            client.Credentials = new NetworkCredential(userName, passWord, domain);

            byte[] buffer = client.DownloadData(path);
            return buffer;
        }


        public virtual byte[] Modified(string path, string query)
        {

            byte[] buffer = GetBuffer(path);

            DataTable dt = new DataTable();
            try
            {
                using (DbDataAdapter adapter = new OleDbDataAdapter(query, ConnectionString))
                {
                    adapter.Fill(dt);
                }

                if (dt.Rows.Count > 0)
                {
                    DataRow dataRow = dt.Rows[0];
                    using (MemoryStream ms = new MemoryStream())
                    {
                        ms.Write(buffer, 0, buffer.Length);
                        using (WordprocessingDocument doc = WordprocessingDocument.Open(ms, true))
                        {
                            doc.MainDocumentPart.DocumentSettingsPart.Settings.RemoveAllChildren
                                <DocumentFormat.OpenXml.Wordprocessing.MailMerge>();
                            Body body = doc.MainDocumentPart.Document.Body;

                            var list =
                                body.Descendants<FieldChar>().Where(c => c.FieldCharType == FieldCharValues.Begin).Select(
                                    c => c.Parent as Run);

                            // Process complex MergeFields 
                            foreach (Run run in list)
                            {
                                Run current = run;
                                string column = "";
                                string format = "";

                                while (!(current.GetFirstChild<FieldChar>() != null &&
                                    current.GetFirstChild<FieldChar>().FieldCharType == FieldCharValues.End))
                                {
                                    if (current.GetFirstChild<FieldCode>() != null)
                                    {
                                        string[] columnParts = current.GetFirstChild<FieldCode>()
                                            .Text
                                            .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                                        if (columnParts.Length > 1)
                                        {
                                            if (columnParts[0] != "MERGEFIELD")
                                                break;
                                            column = columnParts[1];
                                            if (columnParts.Length > 3)
                                                format = columnParts[3].Replace('"', ' ');
                                        }

                                    }
                                    Text text = current.GetFirstChild<Text>();

                                    if (dt.Columns.Contains(column) && text != null)
                                    {
                                        if (dataRow[column] is DateTime)
                                            text.Text = ((DateTime)dataRow[column]).ToString(format);
                                        else
                                            text.Text = dataRow[column].ToString();
                                    }
                                    current = current.NextSibling<Run>();
                                }
                            }

                            // Process Simple MergeField
                            foreach (SimpleField field in body.Descendants<SimpleField>())
                            {
                                string column = "";
                                string format = "";
                                if (field.Instruction.HasValue)
                                {
                                    string[] columnParts = field.Instruction.Value
                                           .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                    if (columnParts.Length > 1)
                                    {
                                        column = columnParts[1];
                                        if (columnParts.Length > 3)
                                            format = columnParts[3].Replace('"', ' ');
                                    }
                                    Text text = field.Descendants<Text>().FirstOrDefault();
                                    if (dt.Columns.Contains(column) && text != null)
                                    {
                                        if (dataRow[column] is DateTime)
                                            text.Text = ((DateTime)dataRow[column]).ToString(format);
                                        else
                                            text.Text = dataRow[column].ToString();
                                    }
                                }
                            }

                        }
                        return ms.ToArray();
                    }
                }
                else
                {
                    _log(String.Format("No Data returned.\tQuery:\r\n{0}", query), EventLogEntryType.Information);
                }
            }
            catch (Exception ex)
            {
                _log(ex.Message, EventLogEntryType.Error);
                _log(ex.ToString(), EventLogEntryType.Information);
                _log(ex.StackTrace, EventLogEntryType.Information);
            }

#if false
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(buffer, 0, buffer.Length);
                using (WordprocessingDocument doc = WordprocessingDocument.Open(ms, true))
                {
                    doc.MainDocumentPart.DocumentSettingsPart.Settings.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.MailMerge>();
                    Body body = doc.MainDocumentPart.Document.Body;

                    var list = body.Descendants<FieldChar>().Where(c => c.FieldCharType == FieldCharValues.Begin).Select(c => c.Parent as Run);

                    // Process complex MergeFields 
                    foreach (Run run in list)
                    {
                        Run current = run;
                        string column = "";
                        string format = "";

                        while (!(current.GetFirstChild<FieldChar>() != null &&
                            current.GetFirstChild<FieldChar>().FieldCharType == FieldCharValues.End))
                        {
                            if (current.GetFirstChild<FieldCode>() != null)
                            {
                                string[] columnParts = current.GetFirstChild<FieldCode>()
                                    .Text
                                    .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                                if (columnParts.Length > 1)
                                {
                                    if (columnParts[0] != "MERGEFIELD")
                                        break;
                                    column = columnParts[1];
                                    if (columnParts.Length > 3)
                                        format = columnParts[3].Replace('"', ' ');
                                }

                            }
                            Text text = current.GetFirstChild<Text>();

                            if (_dt.Columns.Contains(column) && text != null)
                            {
                                if (dataRow[column] is DateTime)
                                    text.Text = ((DateTime)dataRow[column]).ToString(format);
                                else
                                    text.Text = dataRow[column].ToString();
                            }
                            current = current.NextSibling<Run>();
                        }
                    }

                    // Process Simple MergeField
                    foreach (SimpleField field in body.Descendants<SimpleField>())
                    {
                        string column = "";
                        string format = "";
                        if (field.Instruction.HasValue)
                        {
                            string[] columnParts = field.Instruction.Value
                                   .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            if (columnParts.Length > 1)
                            {
                                column = columnParts[1];
                                if (columnParts.Length > 3)
                                    format = columnParts[3].Replace('"', ' ');
                            }
                            Text text = field.Descendants<Text>().FirstOrDefault();
                            if (_dt.Columns.Contains(column) && text != null)
                            {
                                if (dataRow[column] is DateTime)
                                    text.Text = ((DateTime)dataRow[column]).ToString(format);
                                else
                                    text.Text = dataRow[column].ToString();
                            }
                        }
                    }
                }

                using (FileStream fs = new FileStream(path.Combine(OutputDir, targetPath), FileMode.Create))
                {
                    ms.WriteTo(fs);
                }
            }
#endif

            return null;
        }

#if false
        public virtual XmlDocument Modified(string path, string query, string org, string sqlMechineName)
        {
            try
            {
                var doc = new XmlDocument();
                // Create and load the XmlDocument.
                doc = new XmlDocument();
                _log("path source to resolve =" + path, LogType.Verbose);
                if (path.ToLower().Contains("http://"))
                {
                    var resolver = new XmlUrlResolver();
                    string userName = ConfigurationManager.AppSettings["userName"];
                    if (String.IsNullOrEmpty(userName))
                        throw new Exception("userName is null see config file");

                    string domain = ConfigurationManager.AppSettings["domain"];
                    if (String.IsNullOrEmpty(domain))
                        throw new Exception("domain is null see config file");

                    string passWord = ConfigurationManager.AppSettings["passWord"];
                    if (String.IsNullOrEmpty(passWord))
                        throw new Exception("passWord is null see config file");

                    resolver.Credentials = new NetworkCredential(userName, passWord, domain);
                    doc.XmlResolver = resolver;  // Set the resolver.
                }

                doc.Load(path);


                return Modified(doc, query, org, sqlMechineName);
            }
            catch (Exception ee)
            {
                _log(ee.ToString(), LogType.Error);
                throw;
            }

        }

#endif

        protected void SetNodeByXpath(XmlNode mailMergeNode, XmlNamespaceManager xmlnsManager, string xPathSelector, string value)
        {
            _log("xPathSelector = " + xPathSelector + " value=" + value, EventLogEntryType.Information);
            XmlNode cx = mailMergeNode.SelectSingleNode(xPathSelector, xmlnsManager);
            cx.Attributes[AttributeValue].Value = value;
        }

        protected virtual XmlDocument Modified(XmlDocument doc, string query, string org, string sqlMechineName)
        {
            _log("Modified = query" + query + " org=" + org + " sqlMechineName=" + sqlMechineName, EventLogEntryType.Information);
            var xmlnsManager = PolpuateXmlNamespaceManager(doc);
            var mailMergeNode = doc.SelectSingleNode("//pkg:part[@pkg:name = '/word/settings.xml']/pkg:xmlData/w:settings/w:mailMerge", xmlnsManager);

            SetNodeByXpath(mailMergeNode, xmlnsManager, "w:connectString", ConnectionString);
            SetNodeByXpath(mailMergeNode, xmlnsManager, "w:odso/w:udl",ConnectionString);
            SetNodeByXpath(mailMergeNode, xmlnsManager, "w:query", query);
            return doc;
        }

    }
}

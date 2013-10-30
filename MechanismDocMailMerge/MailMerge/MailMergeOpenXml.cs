using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using D = DocumentFormat.OpenXml.Packaging;
using XOPEN = DocumentFormat.OpenXml;

using System.Net;
using System.Configuration;
using System.Diagnostics;
using Guardian.Documents.MailMerge.Contract;

namespace Guardian.Documents.MailMerge
{
    public class MailMergeOpenXml
    {
        
        public string ConnectionString { get; protected set; }

        Action<string, EventLogEntryType> _log;

        public MailMergeOpenXml(Action<string, EventLogEntryType> log, string connection)
        {
            _log = log;
            ConnectionString = connection; 
        }

        public string Merge(string query,ISourceDoc sourceDoc,ITargetDoc targetDoc)
        {
            var buffer = sourceDoc.GetBuffer();
            var dataAfterModified = Modified(buffer, query);
            string targetPath = targetDoc.Save(dataAfterModified);
            return targetPath;

        }

        protected  byte[] Modified(byte[] buffer, string query)
        {
            using (var ms = new MemoryStream())
            {
                ms.Write(buffer, 0, buffer.Length);
                using (var doc = XOPEN.Packaging.WordprocessingDocument.Open(ms, true))
                {
                    int mailmergecount = doc.MainDocumentPart.DocumentSettingsPart.Settings.Elements<XOPEN.Wordprocessing.MailMerge>().Count();
                    if (mailmergecount != 1)
                        throw new ArgumentException("mailmergecount is not 1");

                    var mymerge = doc.MainDocumentPart.DocumentSettingsPart.Settings.Elements<XOPEN.Wordprocessing.MailMerge>().First();

                    mymerge.ConnectString.Val = ConnectionString;
                    mymerge.Query.Val = query;
                }
                return ms.ToArray();
            }
        }

        public  byte[] ReplaceMergeFieldValue(ISourceDoc sourceDoc, string query)
        {
            var buffer = sourceDoc.GetBuffer();

            var dt = new DataTable();
            try
            {
                using (var adapter = new OleDbDataAdapter(query, ConnectionString))
                {
                    adapter.Fill(dt);
                }

                if (dt.Rows.Count > 0)
                {
                    DataRow dataRow = dt.Rows[0];
                    using (var ms = new MemoryStream())
                    {
                        ms.Write(buffer, 0, buffer.Length);
                        using (var doc = XOPEN.Packaging.WordprocessingDocument.Open(ms, true))
                        {
                            doc.MainDocumentPart.DocumentSettingsPart.Settings.RemoveAllChildren
                                <XOPEN.Wordprocessing.MailMerge>();
                           var body = doc.MainDocumentPart.Document.Body;

                            var list =
                                body.Descendants<XOPEN.Wordprocessing.FieldChar>().Where(c => c.FieldCharType == XOPEN.Wordprocessing.FieldCharValues.Begin).Select(
                                    c => c.Parent as XOPEN.Wordprocessing.Run);

                            // Process complex MergeFields 
                            foreach (var run in list)
                            {
                                var current = run;
                                string column = "";
                                string format = "";

                                while (!(current.GetFirstChild<XOPEN.Wordprocessing.FieldChar>() != null &&
                                    current.GetFirstChild<XOPEN.Wordprocessing.FieldChar>().FieldCharType == XOPEN.Wordprocessing.FieldCharValues.End))
                                {
                                    if (current.GetFirstChild<XOPEN.Wordprocessing.FieldCode>() != null)
                                    {
                                        string[] columnParts = current.GetFirstChild<XOPEN.Wordprocessing.FieldCode>()
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
                                    var text = current.GetFirstChild<XOPEN.Wordprocessing.Text>();

                                    if (dt.Columns.Contains(column) && text != null)
                                    {
                                        if (dataRow[column] is DateTime)
                                            text.Text = ((DateTime)dataRow[column]).ToString(format);
                                        else
                                            text.Text = dataRow[column].ToString();
                                    }
                                    current = current.NextSibling<XOPEN.Wordprocessing.Run>();
                                }
                            }

                            // Process Simple MergeField
                            foreach (var field in body.Descendants<XOPEN.Wordprocessing.SimpleField>())
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
                                    var text = field.Descendants<XOPEN.Wordprocessing.Text>().FirstOrDefault();
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
                    Log(String.Format("No Data returned.\tQuery:\r\n{0}", query), EventLogEntryType.Information);
                }
            }
            catch (Exception ex)
            {
                Log(ex.Message, EventLogEntryType.Error);
                Log(ex.ToString(), EventLogEntryType.Information);
                Log(ex.StackTrace, EventLogEntryType.Information);
            }



            return null;
        }

        void Log(string s, EventLogEntryType t)
        {
            if (_log!=null)
            {
                _log(s, t);
            }
        }

    }
}
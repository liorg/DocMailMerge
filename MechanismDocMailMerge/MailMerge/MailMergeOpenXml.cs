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

        protected virtual byte[] Modified(byte[] buffer, string query)
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

        void Log(string s, EventLogEntryType t)
        {
            if (_log!=null)
            {
                _log(s, t);
            }
        }

    }
}
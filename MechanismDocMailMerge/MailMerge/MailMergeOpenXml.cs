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
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.Globalization;
using MechanismDocMailMerge.MailMerge;

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

        public DocPropertiey Merge(string query, ISourceDoc sourceDoc, ITargetDoc targetDoc, string udlPath, string macroTemplatePath = "", string serverAndPort = "")
        {
            var buffer = sourceDoc.GetBuffer();
            var dataAfterModified = Modified(buffer, query, udlPath, macroTemplatePath, serverAndPort);
            DocPropertiey targetPath = targetDoc.Save(dataAfterModified);
            return targetPath;
        }

        public DocPropertiey FillData(ISourceDoc sourceDoc, ITargetDoc targetDoc)
        {
            var dataAfterModified = sourceDoc.GetBuffer();
            var dataAfterChange = FillDataToDoc(dataAfterModified, true);
            DocPropertiey targetPath = targetDoc.Save(dataAfterChange);
            return targetPath;
        }

        bool FindRemoveWatermark(Run runWatermark)
        {
            bool success = false;
            //DocumentFormat.OpenXml.Vml.TextPath
            //Check, if run contains watermark
            if (runWatermark.Descendants<Picture>() != null)
            {
                var listPic = runWatermark.Descendants<Picture>().ToList();

                for (int n = listPic.Count; n > 0; n--)
                {
                    if (listPic[n - 1].Descendants<DocumentFormat.OpenXml.Vml.Shape>() != null)
                    {
                        if (listPic[n - 1].Descendants<DocumentFormat.OpenXml.Vml.Shape>().Where(s => s.Type == "#_x0000_t136").Count() > 0)
                        {
                            //Found -> remove
                            listPic[n - 1].Remove();
                            success = true;
                            break;
                        }
                    }
                }

            }

            return success;
        }

        byte[] FillDataToDoc(byte[] buffer, bool isRemoveWatermark)
        {

            DataTable dt = new DataTable();
            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    ms.Write(buffer, 0, buffer.Length);
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(ms, true))
                    {

                        var query = GetQuery(doc);
                        using (DbDataAdapter adapter = new OleDbDataAdapter(query, ConnectionString))
                        {
                            adapter.Fill(dt);
                        }

                        if (dt.Rows.Count > 0)
                        {
                            DataRow dataRow = dt.Rows[0];
                            // Remove datasource from settings part and save this part in the memory stream
                            doc.MainDocumentPart.DocumentSettingsPart.Settings.RemoveAllChildren
                                <DocumentFormat.OpenXml.Wordprocessing.MailMerge>();
                            doc.MainDocumentPart.DocumentSettingsPart.Settings.Save();


                            if (isRemoveWatermark)
                            {
                                foreach (var header in doc.MainDocumentPart.HeaderParts)
                                {
                                    //Remove
                                    if (header.Header.Descendants<Paragraph>() != null)
                                    {
                                        var isFound = false;
                                        foreach (var para in header.Header.Descendants<Paragraph>())
                                        {
                                            foreach (Run r in para.Descendants<Run>())
                                            {
                                                isFound = FindRemoveWatermark(r);
                                                if (isFound)
                                                    break;
                                            }
                                            if (isFound)
                                                header.Header.Save(header);
                                        }
                                    }
                                }
                            }
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

                                    IEnumerable<Text> texts = field.Descendants<Text>();
                                    Text text = texts.FirstOrDefault();

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
                        doc.MainDocumentPart.Document.Save();
                        return ms.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                Log(ex.Message, EventLogEntryType.Error);
                Log(ex.ToString(), EventLogEntryType.Information);
                Log(ex.StackTrace, EventLogEntryType.Error);
            }
            return null;
        }

        string GenerateSQL(string query, Guid activeDoc)
        {
            return String.Format(query, activeDoc);
        }

        protected byte[] Modified(byte[] buffer, string query, string udlPath, string macroTemplatePath, string serverAndPort)
        {
            using (var ms = new MemoryStream())
            {
                //   (@"\\Crm11mantad\c$\inetpub\wwwroot\WEBMentaService\Doctemplates\r.udl");
                ms.Write(buffer, 0, buffer.Length);
                using (var doc = XOPEN.Packaging.WordprocessingDocument.Open(ms, true))
                {
                    var propertyName = "server";
                    var settingPart = doc.MainDocumentPart.DocumentSettingsPart;
                    SetMailMergeSetting(doc, settingPart, udlPath, query, ConnectionString);
                    SetCustomProperty(doc, propertyName, serverAndPort);
                    SetMacroPath(doc, settingPart, macroTemplatePath);
                }
                return ms.ToArray();
            }
        }

        //set macro
        void SetMacroPath(WordprocessingDocument doc, DocumentSettingsPart settingPart, string macroTemplatePath)
        {
            if (!String.IsNullOrEmpty(macroTemplatePath))
            {
                var hasMacros = doc.MainDocumentPart.DocumentSettingsPart.Settings.Elements<XOPEN.Wordprocessing.AttachedTemplate>().Count();

                var newUriMacr = new Uri(macroTemplatePath);
                if (hasMacros > 0)
                {
                    AttachedTemplate attachTemplate = doc.MainDocumentPart.DocumentSettingsPart.Settings.Elements<XOPEN.Wordprocessing.AttachedTemplate>().First();

                    var attachTemplateRelationship = settingPart.GetExternalRelationship(attachTemplate.Id);

                    // attachTemplateRelationship.Uri=  new Uri(macroTemplatePath, UriKind.Absolute);

                    if (attachTemplateRelationship != null)
                    {
                        settingPart.DeleteExternalRelationship(attachTemplateRelationship);
                        settingPart.AddExternalRelationship(attachTemplateRelationship.RelationshipType, newUriMacr, attachTemplate.Id);
                    }
                }
            }
        }

        void SetCustomProperty(WordprocessingDocument doc, string propertyName, string propertyValue)
        {
            var newProp = new DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty();
            newProp.VTLPWSTR = new DocumentFormat.OpenXml.VariantTypes.VTLPWSTR(propertyValue);
            newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
            newProp.Name = propertyName;

            var customProps = doc.CustomFilePropertiesPart;
            if (customProps == null)
            {
                // No custom properties? Add the part, and the
                // collection of properties now.
                customProps = doc.AddCustomFilePropertiesPart();
                customProps.Properties =
                    new DocumentFormat.OpenXml.CustomProperties.Properties();
            }

            var props = customProps.Properties;
            if (props != null)
            {
                // This will trigger an exception if the property's Name 
                // property is null, but if that happens, the property is damaged, 
                // and probably should raise an exception.
                var prop =
                    props.Where(
                    p => ((DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty)p).Name.Value
                        == propertyName).FirstOrDefault();

                // Does the property exist? If so, get the return value, 
                // and then delete the property.
                if (prop != null)
                {
                    Debug.WriteLine(prop.InnerText);
                    prop.Remove();
                }

                // Append the new property, and 
                // fix up all the property ID values. 
                // The PropertyId value must start at 2.
                props.AppendChild(newProp);

                int pid = 2;
                foreach (DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty item in props)
                    item.PropertyId = pid++;
            }

        }

        void SetMailMergeSetting(WordprocessingDocument doc, DocumentSettingsPart settingPart, string udlPath, string query, string connectionString)
        {
            int mailmergecount = doc.MainDocumentPart.DocumentSettingsPart.Settings.Elements<XOPEN.Wordprocessing.MailMerge>().Count();
            if (mailmergecount != 1)
                throw new ArgumentException("mailmergecount is not 1");
            // Get the Document Settings Part

            var newUri = new Uri(udlPath);


            var mymerge = settingPart.Settings.Elements<XOPEN.Wordprocessing.MailMerge>().First();
            var myDataSourceReference = mymerge.DataSourceReference;
            if (myDataSourceReference != null)
            {
                var myOldRelationship = settingPart.GetExternalRelationship(myDataSourceReference.Id);
                if (myOldRelationship != null)
                {
                    settingPart.DeleteExternalRelationship(myOldRelationship);
                    settingPart.AddExternalRelationship(myOldRelationship.RelationshipType, newUri, myDataSourceReference.Id);
                }
            }
            if (mymerge.DataSourceObject != null && mymerge.DataSourceObject.SourceReference != null)
            {

                var myOldRelationship2 = settingPart.GetExternalRelationship(mymerge.DataSourceObject.SourceReference.Id);
                if (myOldRelationship2 != null)
                {
                    settingPart.DeleteExternalRelationship(myOldRelationship2);
                    settingPart.AddExternalRelationship(myOldRelationship2.RelationshipType, newUri, mymerge.DataSourceObject.SourceReference.Id);

                }
                if (mymerge.DataSourceObject.UdlConnectionString != null)
                    mymerge.DataSourceObject.UdlConnectionString.Val = connectionString;
            }
            mymerge.ConnectString.Val = connectionString;
            mymerge.Query.Val = query;
        }

        string GetQuery(WordprocessingDocument doc)
        {
            int mailmergecount = doc.MainDocumentPart.DocumentSettingsPart.Settings.Elements<XOPEN.Wordprocessing.MailMerge>().Count();
            if (mailmergecount != 1)
                throw new ArgumentException("mailmergecount is not 1");
            // Get the Document Settings Part


            var settingPart = doc.MainDocumentPart.DocumentSettingsPart;
            var mymerge = settingPart.Settings.Elements<XOPEN.Wordprocessing.MailMerge>().First();

            return mymerge.Query.Val;
        }

        void Log(string s, EventLogEntryType t)
        {
            if (_log != null)
            {
                _log(s, t);
            }
        }

    }
}
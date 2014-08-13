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
namespace Guardian.Documents.MailMerge
{
    /// <summary>
    /// Manipulate Open XML SDK 2.5 
    /// </summary>
    public class MailMergeOpenXml
    {
        public string ConnectionString { get; protected set; }

        Action<string, EventLogEntryType> _log;

        public MailMergeOpenXml(Action<string, EventLogEntryType> log, string connection)
        {
            _log = log;
            ConnectionString = connection;
        }

       /// <summary>
       /// 1. Change Data SOurce
       /// 2. Change SQL Query
       /// 3. Replace path of macro dotm template 
       /// 4. Add custom property for server name and port which connect from macro client side to server api side
       /// </summary>
       /// <param name="query"></param>
       /// <param name="sourceDoc"></param>
       /// <param name="targetDoc"></param>
       /// <param name="udlPath"></param>
       /// <param name="macroTemplatePath"></param>
       /// <param name="serverAndPort"></param>
       /// <returns></returns>
        public DocPropertiey Merge(string query, ISourceDoc sourceDoc, ITargetDoc targetDoc, string udlPath, string macroTemplatePath = "", string serverAndPort = "")
        {
            var buffer = sourceDoc.GetBuffer();
            var dataAfterModified = Modified(buffer, query, udlPath, macroTemplatePath, serverAndPort);
            DocPropertiey targetPath = targetDoc.Save(dataAfterModified);
            return targetPath;
        }

        /// <summary>
        /// Disconnect Data Source from Mail Merge and fill merge field with current query fro doc and also data source from doc
        /// </summary>
        /// <param name="sourceDoc"></param>
        /// <param name="targetDoc"></param>
        /// <returns></returns>
        public DocPropertiey FillData(ISourceDoc sourceDoc, ITargetDoc targetDoc)
        {
            var dataAfterModified = sourceDoc.GetBuffer();
            var dataAfterChange = FillDataToDocCP(dataAfterModified, true);
            DocPropertiey targetPath = targetDoc.Save(dataAfterChange);
            return targetPath;
        }

        //http://www.legalcube.de/post/Word-OpenXML-Create-change-or-delete-watermarks.aspx
        bool FindAndRemoveWatermark(Run runWatermark)
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

        byte[] FillDataToDocCP(byte[] buffer, bool isRemoveWatermark)
        {
            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    ms.Write(buffer, 0, buffer.Length);
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(ms, true))
                    {

                        var query = GetQueryFromDoc(doc);
                        var connStr = GetConnStrFromDoc(doc);
                        Dictionary<string, string> values = GetFieldsValues(query, connStr);
                        if (values.Count > 0)
                        {

                            // Remove datasource from settings part and save this part in the memory stream
                            RemoveMailMergeDataSource(doc);

                            if (isRemoveWatermark)
                                RemoveWatermark(doc);

                            UtilityFiller.GetWordReportPart(ms, doc, values);
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
                throw ex;
            }

        }

        /// <summary>
        /// Remove MailMerge DataSource (before fill mailmege fields)
        /// </summary>
        /// <param name="doc"></param>
        void RemoveMailMergeDataSource(WordprocessingDocument doc)
        {
            doc.MainDocumentPart.DocumentSettingsPart.Settings.RemoveAllChildren
                                   <DocumentFormat.OpenXml.Wordprocessing.MailMerge>();
            doc.MainDocumentPart.DocumentSettingsPart.Settings.Save();
        }

        /// <summary>
        /// Remove Watermark
        /// </summary>
        /// <param name="doc"></param>
        void RemoveWatermark(WordprocessingDocument doc)
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
                            isFound = FindAndRemoveWatermark(r);
                            if (isFound)
                                break;
                        }
                        if (isFound)
                            header.Header.Save(header);
                    }
                }
            }

        }

        /// <summary>
        /// Fill fields names and is values
        /// </summary>
        /// <param name="query"></param>
        /// <param name="connStr"></param>
        /// <returns></returns>
        Dictionary<string, string> GetFieldsValues(string query, string connStr)
        {
            DataRow dataRow = null;
            Dictionary<string, string> values = new Dictionary<string, string>();
            DataTable dt = new DataTable();

            using (System.Data.SqlClient.SqlDataAdapter adapter = new System.Data.SqlClient.SqlDataAdapter(query, connStr))
            {
                adapter.Fill(dt);
            }

            if (dt.Rows.Count > 0)
            {
                dataRow = dt.Rows[0];
            }
            if (dataRow == null)
                return values;
            // values.Add("businessunit", "שלום רב  manta בדיקה");
            foreach (DataColumn item in dt.Columns)
            {
                if (values.ContainsKey(item.ColumnName))
                    continue;
                values.Add(item.ColumnName, dataRow[item].ToString());

            }
            return values;
        }

        [Obsolete("roman  will be removed", false)]
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

                        var query = GetQueryFromDoc(doc);
                        var connStr = GetConnStrFromDoc(doc);
                        using (System.Data.SqlClient.SqlDataAdapter adapter = new System.Data.SqlClient.SqlDataAdapter(query, connStr))
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
                                                isFound = FindAndRemoveWatermark(r);
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
                throw ex;
            }

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

        string GetQueryFromDoc(WordprocessingDocument doc)
        {
            int mailmergecount = doc.MainDocumentPart.DocumentSettingsPart.Settings.Elements<XOPEN.Wordprocessing.MailMerge>().Count();
            if (mailmergecount != 1)
                throw new ArgumentException("mailmergecount is not 1");
            // Get the Document Settings Part


            var settingPart = doc.MainDocumentPart.DocumentSettingsPart;
            var mymerge = settingPart.Settings.Elements<XOPEN.Wordprocessing.MailMerge>().First();

            return mymerge.Query.Val;
        }

        string GetConnStrFromDoc(WordprocessingDocument doc)
        {
            int mailmergecount = doc.MainDocumentPart.DocumentSettingsPart.Settings.Elements<XOPEN.Wordprocessing.MailMerge>().Count();
            if (mailmergecount != 1)
                throw new ArgumentException("mailmergecount is not 1");
            // Get the Document Settings Part


            var settingPart = doc.MainDocumentPart.DocumentSettingsPart;
            var mymerge = settingPart.Settings.Elements<XOPEN.Wordprocessing.MailMerge>().First();

            return mymerge.ConnectString.Val;
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

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
using System.IO.Packaging;



namespace Guardian.Documents.MailMerge
{
    /// <summary>
    /// Manipulate Open XML SDK 2.5 
    /// </summary>
    public class MailMergeOpenXml : IMailMergeOpenXml
    {


        Action<string, EventLogEntryType> _log;

        public MailMergeOpenXml(Action<string, EventLogEntryType> log)
        {
            _log = log;
        }

        /// <summary>
        /// 1. Change Data SOurce
        /// 2. Change SQL Query
        /// 3. Replace path of macro dotm template 
        /// 4. Add custom property for server name and port which connect from macro client side to server api side
        /// </summary>
        public DocPropertiey Merge(string connStr, string query, ISourceDoc sourceDoc, ITargetDoc targetDoc, string udlPath, string macroTemplatePath = "", Dictionary<string, string> customProperties = null)
        {
            var buffer = sourceDoc.GetBuffer();
            var dataAfterModified = Modified(connStr, buffer, query, udlPath, macroTemplatePath, customProperties);
            DocPropertiey targetPath = targetDoc.Save(dataAfterModified);
            return targetPath;
        }

        /// <summary>
        /// Disconnect Data Source from Mail Merge and fill merge field with current query fro doc and also data source from doc
        /// </summary>
        /// <param name="sourceDoc"></param>
        /// <param name="targetDoc"></param>
        /// <returns></returns>
        public DocPropertiey FillData(ISourceDoc sourceDoc, ITargetDoc targetDoc, string connectionString = null)
        {
            var dataAfterModified = sourceDoc.GetBuffer();
            var dataAfterChange = FillDataToDocCP(dataAfterModified, true, connectionString);
            DocPropertiey targetPath = targetDoc.Save(dataAfterChange);
            return targetPath;
        }

        public DocPropertiey FillDataAndConvertDocx(ISourceDoc sourceDoc, ITargetDoc targetDoc, string connectionString = null)
        {
            var dataAfterModified = sourceDoc.GetBuffer();
            var dataAfterChange = FillDataToDocCP(dataAfterModified, true, connectionString);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(dataAfterChange, 0, dataAfterChange.Length);
                ChangeDocmToDocxUsingPackage(ms);
                RemoveMacroOpenXml(ms);
                DocPropertiey targetPath = targetDoc.Save(ms.ToArray());
                return targetPath;
            }
         
        }

        /// <summary>
        /// Convert Docm To Docx (By package because ChangeDocumentType not working well)
        /// </summary>
        /// <param name="sourceDoc"></param>
        /// <param name="targetDoc"></param>
        public void ChangeDocmToDocx(ISourceDoc sourceDoc, ITargetDoc targetDoc)
        {
            var buffer = sourceDoc.GetBuffer();
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(buffer, 0, buffer.Length);
                ChangeDocmToDocxUsingPackage(ms);
                RemoveMacroOpenXml(ms);
                DocPropertiey targetPath = targetDoc.Save(ms.ToArray());
            }
        }

        /// <summary>
        /// for safe remove vb code from word office
        /// </summary>
        /// <param name="ms"></param>
        public void RemoveMacroOpenXml(MemoryStream ms)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(ms, true))
            {
                var docPart = doc.MainDocumentPart;
                var vbaPart = docPart.VbaProjectPart;
                if (vbaPart != null)
                {
                    //    // Delete the vbaProject part and then save the document.
                    docPart.DeletePart(vbaPart);
                    docPart.Document.Save();
                    //  No work instead ChangeDocmToDocxUsingPackage will change format
                    // doc.ChangeDocumentType(WordprocessingDocumentType.Document);        
                }
                //var hasMacroAttachments  = doc.MainDocumentPart.DocumentSettingsPart.Settings.Elements<XOPEN.Wordprocessing.AttachedTemplate>().Count();
                //var settingPart = doc.MainDocumentPart.DocumentSettingsPart;
                // if (hasMacroAttachments > 0)
                //{
                //    foreach (var attachTemplate in doc.MainDocumentPart.DocumentSettingsPart.Settings.Elements<XOPEN.Wordprocessing.AttachedTemplate>())
                //    {
                //        var attachTemplateRelationship = settingPart.GetExternalRelationship(attachTemplate.Id);
                //        if (attachTemplateRelationship != null)
                //        {
                //            settingPart.DeleteExternalRelationship(attachTemplateRelationship);

                //        }
                //    }
                    
                //}

            }
        }

        private void CopyStream(Stream source, Stream target)
        {
            const int bufSize = 16384;
            byte[] buf = new byte[bufSize];
            int bytesRead = 0;
            while ((bytesRead = source.Read(buf, 0, bufSize)) > 0)
                target.Write(buf, 0, bytesRead);
        }

        /// <summary>
        ///By package because  ChangeDocumentType not working well
        /// </summary>
        /// <param name="documentStream"></param>
        private void ChangeDocmToDocxUsingPackage(Stream documentStream)
        {
            // Open the document in the stream and replace the custom XML part
            using (System.IO.Packaging.Package packageFile = System.IO.Packaging.Package.Open(documentStream, FileMode.Open, FileAccess.ReadWrite))
            {
                System.IO.Packaging.PackagePart packagePart = null;
                // Find part containing the correct namespace
                foreach (var part in packageFile.GetParts())
                {
                    if (part.ContentType.Equals("application/vnd.ms-word.document.macroEnabled.main+xml", StringComparison.OrdinalIgnoreCase))
                    {
                        packagePart = part;
                        break;
                    }
                }
                if (packagePart != null)
                {
                    using (MemoryStream source = new MemoryStream())
                    {
                        CopyStream(packagePart.GetStream(), source);

                        var saveRelationBeforeDelPart = new List<PackageRelationship>();
                        foreach (var item in packagePart.GetRelationships())
                        {
                            saveRelationBeforeDelPart.Add(item);
                        }

                        Uri uriData = packagePart.Uri;
                        // Delete the existing XML part
                        if (packageFile.PartExists(uriData))
                            packageFile.DeletePart(uriData);

                        // Load the custom XML data
                        var pkgprtData = packageFile.CreatePart(uriData, "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", System.IO.Packaging.CompressionOption.SuperFast);

                        source.Position = 0;//reset position
                        CopyStream(source, pkgprtData.GetStream(FileMode.Create));

                        foreach (var copyRel in saveRelationBeforeDelPart)
                        {
                            pkgprtData.CreateRelationship(copyRel.TargetUri, copyRel.TargetMode, copyRel.RelationshipType, copyRel.Id);
                        }
                    }
                }
            }
        }

        byte[] FillDataToDocCP(byte[] buffer, bool isRemoveWatermark, string connectionString = null)
        {
            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    ms.Write(buffer, 0, buffer.Length);
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(ms, true))
                    {

                        var query = GetQueryFromDoc(doc);
                        var connStr = connectionString == null ? GetConnStrFromDoc(doc) : connectionString;
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

        /// <summary>
        /// Remove MailMerge DataSource (before fill data on mailmerge fields)
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
            Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); // do not care about key case sensetive
            DataTable dt = new DataTable();

            // using (var adapter = new System.Data.SqlClient.SqlDataAdapter(query, connStr))
            using (var adapter = new System.Data.OleDb.OleDbDataAdapter(query, connStr))
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

        protected byte[] Modified(string connStr, byte[] buffer, string query, string udlPath, string macroTemplatePath, Dictionary<string, string> customProperties)
        {
            using (var ms = new MemoryStream())
            {
                //   (@"\\Crm11mantad\c$\inetpub\wwwroot\WEBMentaService\Doctemplates\r.udl");
                ms.Write(buffer, 0, buffer.Length);
                using (var doc = XOPEN.Packaging.WordprocessingDocument.Open(ms, true))
                {
                    //var propertyName = "server";
                    //SetCustomProperty(doc, propertyName, serverAndPort);
                    if (customProperties != null && customProperties.Any())
                    {
                        foreach (var propertyKey in customProperties.Keys)
                        {
                            SetCustomProperty(doc, propertyKey, customProperties[propertyKey]);
                        }
                    }
                    var settingPart = doc.MainDocumentPart.DocumentSettingsPart;
                    SetMailMergeSetting(doc, settingPart, udlPath, query, connStr);
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

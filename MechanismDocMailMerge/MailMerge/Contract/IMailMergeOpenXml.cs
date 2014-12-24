using System;
namespace Guardian.Documents.MailMerge
{
    public interface IMailMergeOpenXml
    {
        void ChangeDocmToDocx(global::Guardian.Documents.MailMerge.Contract.ISourceDoc sourceDoc, global::Guardian.Documents.MailMerge.Contract.ITargetDoc targetDoc);
        global::Guardian.Documents.MailMerge.DocPropertiey FillData(global::Guardian.Documents.MailMerge.Contract.ISourceDoc sourceDoc, global::Guardian.Documents.MailMerge.Contract.ITargetDoc targetDoc, string connectionString = null);
        global::Guardian.Documents.MailMerge.DocPropertiey Merge(string connStr, string query, global::Guardian.Documents.MailMerge.Contract.ISourceDoc sourceDoc, global::Guardian.Documents.MailMerge.Contract.ITargetDoc targetDoc, string udlPath, string macroTemplatePath = "", global::System.Collections.Generic.Dictionary<string, string> customProperties = null);
    }
}

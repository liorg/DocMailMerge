using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Guardian.Documents.MailMerge;
using UT.Implementation;

namespace UT
{
    class Program
    {
        static void Log(string s, System.Diagnostics.EventLogEntryType e)
        {
            Console.WriteLine(s);
        }
        static void Main(string[] args)
        {
            string connectionToChange = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MOIN_MSCRM;Data Source=crm11moin"; // "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MANTA_MSCRM;Data Source=CRM11MANTAD";
            var queryToChange = "select top 1 new_name as name from new_provider;";
            var pathTarget = @"C:\Users\lior_g\Documents\GitHub\DocMailMerge\TemplatesWords\output2.docx";
            var pathSource="http://localhost/TemplatesWords/name.docx";
            var mailMergeOpenXml = new MailMergeOpenXml(Log, connectionToChange);

            var source = new SourceWebDoc(pathSource);
            var target = new TargetLocalDoc(pathTarget);

            var result=mailMergeOpenXml.Merge(queryToChange, source, target);

        }
    }
}

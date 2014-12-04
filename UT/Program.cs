using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Guardian.Documents.MailMerge;
using UT.Implementation;
using Guardian.MailMerge.Implementation;

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
            var pathSource = "http://localhost/TemplatesWords/name.docx";
            var mailMergeOpenXml = new MailMergeOpenXml(Log);


            pathSource = "http://crm11mantad:8080/Doctemplates/output/tal.docm";
            pathTarget = @"c:\temp\tal.docm";

            var source = new SourceWebDoc(pathSource);
            var target = new TargetLocalDoc(pathTarget);

            mailMergeOpenXml.FillData(source, target);
            return;




            var customProperties = new Dictionary<string, string>();
            customProperties.Add("server", "");
            customProperties.Add("entityid", Guid.NewGuid().ToString());
            customProperties.Add("tempfolder", "c:\\temp\r.udl");
            var result = mailMergeOpenXml.Merge(connectionToChange, queryToChange, source, target, @"c:\\temp\r.udl", @"c:\\temp\TemplateCrmMenta.dotm", customProperties);

        }
    }
}

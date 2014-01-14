using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Guardian.Documents.MailMerge.Contract;

namespace Guardian.MailMerge.Implementation
{
    class SourceWebDoc : ISourceDoc
    {
        string _urlPath;
        public SourceWebDoc(string urlPath)
        {
            _urlPath = urlPath;
        }

        public byte[] GetBuffer()
        {
            WebClient client = new WebClient();

            //string userName = ConfigurationManager.AppSettings["userName"];
            //if (String.IsNullOrEmpty(userName))
            //    throw new Exception("userName is null see config file");

            //string domain = ConfigurationManager.AppSettings["domain"];
            //if (String.IsNullOrEmpty(domain))
            //    throw new Exception("domain is null see config file");

            //string passWord = ConfigurationManager.AppSettings["passWord"];
            //if (String.IsNullOrEmpty(passWord))
            //    throw new Exception("passWord is null see config file");

            //client.Credentials = new NetworkCredential(userName, passWord, domain);
            client.Credentials = System.Net.CredentialCache.DefaultCredentials;
            byte[] buffer = client.DownloadData(_urlPath);
            return buffer;
        }
    }
}

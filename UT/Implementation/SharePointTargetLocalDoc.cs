using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Protocols;
using Guardian.Documents.MailMerge.Contract;
using Guardian.Documents.MailMerge;


namespace UT.Implementation
{
    class SharePointTargetLocalDoc : ITargetDoc
    {
        string _webUrl, _folderTarget, _targetFileName;
        NetworkCredential _credentials;
        public SharePointTargetLocalDoc(string webUrl, string folderTarget, string targetFileName)
        {
            _webUrl = webUrl; _folderTarget = folderTarget; _targetFileName = targetFileName;
        }

        string GetTargetUrl(string targetFileName)
        {
            return new Uri(new Uri(_webUrl + _folderTarget), targetFileName).AbsoluteUri;
        }

        Tuple<bool, string> Upload(byte[] sourceContents = null)
        {
            try
            {

                var targetUrl = GetTargetUrl(_targetFileName);

                if (sourceContents != null)
                {
                    WebClient client = new WebClient();
                    client.Credentials = Credentials;
                    byte[] data = client.UploadData(targetUrl, "PUT", sourceContents);
                }

                return new Tuple<bool, string>(true, targetUrl);
            }
            catch (Exception ee)
            {
                string message = ee.ToString();
                if (ee is SoapException)
                {
                    SoapException soapExc = (SoapException)ee;
                    message += soapExc.Detail != null & soapExc.Detail.InnerText != null ? soapExc.Detail.InnerText : soapExc.Message;
                    //_log(soapExc.Detail.InnerText, LogType.Error);
                }
                //_log(message, LogType.Error);
                return new Tuple<bool, string>(false, String.Empty);
            } 
        }

        private NetworkCredential Credentials
        {
            get
            {
                if (_credentials == null)
                {
                    string userName = ConfigurationManager.AppSettings["userName"];
                    if (String.IsNullOrEmpty(userName))
                        throw new Exception("userName is null see config file");

                    string domain = ConfigurationManager.AppSettings["domain"];
                    if (String.IsNullOrEmpty(domain))
                        throw new Exception("domain is null see config file");

                    string passWord = ConfigurationManager.AppSettings["passWord"];
                    if (String.IsNullOrEmpty(passWord))
                        throw new Exception("passWord is null see config file");
                    _credentials = new NetworkCredential(userName, passWord, domain);
                }
                return _credentials;
            }
        }

        public DocPropertiey Save(byte[] data)
        {
            DocPropertiey d = new DocPropertiey();
            var s= Upload(data);
            d.Drl= s.Item2;
            return d;
        }
    }

}

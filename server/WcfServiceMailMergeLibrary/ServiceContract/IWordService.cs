using Guardian.Menta.MentaServicesLibrary.common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Activation;
using System.ServiceModel.Web;
using System.Text;

namespace Guardian.Menta.MentaServicesLibrary.WordService
{


    [ServiceContract(Namespace = "http://g-s.co.il/ServiceContract/WordService/IWordService")]
    
    public interface IWordService
    {
       // [WebInvoke(Method = "Get", ResponseFormat = WebMessageFormat.Json, RequestFormat = WebMessageFormat.Json)]
        [WebGet(UriTemplate = "/Ping/{idNum}")]
        [OperationContract]
        ResponseBase Ping(string idNum);

     
    }
}

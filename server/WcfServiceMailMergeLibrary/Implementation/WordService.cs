
using Guardian.Menta.MentaServicesLibrary.common;
using Guardian.Menta.MentaServicesLibrary.WordService;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Activation;
using System.Text;


namespace Guardian.Menta.MentaServicesLibrary.WordService
{
    [ServiceBehavior(Namespace = "http://g-s.co.il/ServiceBehavior/WordService/WordService", Name = "WordService")]
    //[AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Required)]
    [AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Allowed)]
    public class WordService : IWordService
    {
        public ResponseBase Ping(string param)
        {

            return new ResponseBase { IsError = false, ErrDescription =param+ " ok" };
        }

        
    }
}

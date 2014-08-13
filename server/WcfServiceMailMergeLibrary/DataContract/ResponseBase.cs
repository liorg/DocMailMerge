using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;

namespace Guardian.Menta.MentaServicesLibrary.common
{
    [DataContract(Namespace = "http://g-s.co.il/DataContract/ResponseBase")]
    public class ResponseBase
    {
        [DataMember]
        public bool IsError { get; set; }

        [DataMember]
        public string ErrDescription { get; set; }
    }
}

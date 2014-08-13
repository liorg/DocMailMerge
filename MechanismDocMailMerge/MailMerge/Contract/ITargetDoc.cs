using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Guardian.Documents.MailMerge.Contract
{
    public interface ITargetDoc
    {
        DocPropertiey Save(byte[] data);
       
    }
}

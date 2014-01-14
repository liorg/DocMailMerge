using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Guardian.Documents.MailMerge.Contract;

namespace Guardian.MailMerge.Implementation
{
    class TargetLocalDoc : ITargetDoc
    {
        string _pathTarget;
        public TargetLocalDoc(string pathTarget)
        {
            _pathTarget = pathTarget;
        }
        public string Save(byte[] data)
        {
            using (FileStream stream = new FileStream(_pathTarget, FileMode.OpenOrCreate))
            {
                using (BinaryWriter writer = new BinaryWriter(stream))
                {
                    writer.Write(data);

                }
            }
            return _pathTarget;
        }
    }

}

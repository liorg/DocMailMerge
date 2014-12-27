
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace exportTemplate.DataModel
{
    [Serializable]
    public class DocumentTemplate
    {
        public Guid Id { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }

        public string SourceSp { get; set; }//new_sourcesp

        [XmlIgnore]
        public string Sql { get; set; }

        [XmlElement("Sql")]
        public System.Xml.XmlCDataSection SqlQuery
        {

            get
            {
                return new System.Xml.XmlDocument().CreateCDataSection(Sql);
            }

            set

            {
                Sql = value.Value;
            }
        }

        public Guid DocType { get; set; }
        //new_days_to_get_reminder
        public int? DaysReminder { get; set; }

        //[XmlIgnore]
        //public EntityReference New_Reminder_Template
        //{
        //    get
        //    {
        //        if (ReminderTemplate.HasValue)
        //            return new EntityReference("new_doc_template", ReminderTemplate.Value);
        //        else return null;
        //    }
        //    set
        //    {
        //        if (value == null)
        //        {
        //            ReminderTemplate = null;
        //        }
        //        else
        //        {
        //            ReminderTemplate = value.Id;
        //        }
        //    }

        //}

        public Guid? ReminderTemplate { get; set; }

    }
}

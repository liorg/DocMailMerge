using exportTemplate.DataModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace WinTestMerge
{
    public partial class frmTest : Form
    {

        public frmTest()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (lstXml.SelectedItems != null && lstXml.SelectedItems.Count > 0)
            {
                MessageBox.Show(lstXml.SelectedItems[0].Tag.ToString());
                var sql = _documentsTemplates.Where(s => s.Id ==Guid.Parse( lstXml.SelectedItems[0].Tag.ToString())).Select(ss=>ss.Sql).FirstOrDefault();
                MessageBox.Show(sql);
            }

        }
        List<DocumentTemplate> _documentsTemplates = null;
        void LoadXml()
        {
            var _pathDocsTemplates = @"C:\export\docsTemplates.xml";
            XmlSerializer serializerDocsTypes = new XmlSerializer(typeof(List<DocumentTemplate>));

          

            using (StreamReader reader = new StreamReader(_pathDocsTemplates))
            {
                _documentsTemplates = (List<DocumentTemplate>)serializerDocsTypes.Deserialize(reader);
            }
           // lstXml.Items = _documentsTemplates;


            foreach (var docUri in _documentsTemplates)
            {

               // string fileName = Path.GetFileName(file);
                ListViewItem item = new ListViewItem(docUri.Name);
                item.Tag = docUri.Id;
                item.Text = docUri.Name;
                lstXml.Items.Add(item);
            }
      
        }

        private void frmTest_Load(object sender, EventArgs e)
        {
            LoadXml();

        }
    }
}

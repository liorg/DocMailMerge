using Guardian.Documents.MailMerge;
using Guardian.MailMerge.Implementation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
 
namespace WinTestMerge
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
 
         const string SourceUri="http://mvmentcrmqa:8080//Doctemplates/output/merge";
         const string TargetFolder = @"\\mvmentcrmqa\c$\inetpub\wwwroot\WEBMentaService\Doctemplates\output\disconnect";
         const string SourceFolder = @"\\mvmentcrmqa\Doctemplates\output\merge";
 
 
        private void Form1_Load(object sender, EventArgs e)
        {
            Refresh();
 
        }
 
        void Refresh()
        {
            lstFiles.Items.Clear();
            lstFiles.Dock = DockStyle.Fill;
            string[] files = Directory.GetFiles(SourceFolder);
            foreach (string file in files)
            {
 
                string fileName = Path.GetFileName(file);
                ListViewItem item = new ListViewItem(fileName);
                item.Tag = file;
                item.Text = fileName;
 
                lstFiles.Items.Add(item);
 
 
            }
 
            files = Directory.GetFiles(TargetFolder);
            lstTarget.Items.Clear();
            lstTarget.Dock = DockStyle.Fill;
 
            foreach (string file in files)
            {
                string fileName = Path.GetFileName(file);
                ListViewItem item = new ListViewItem(fileName);
                item.Tag = file;
                item.Text = fileName;
 
                lstTarget.Items.Add(item);
            }
        }
 
        private void btnDisconnect_Click(object sender, EventArgs e)
        {
            if (lstFiles.SelectedItems != null && lstFiles.SelectedItems.Count > 0)
            {
                 String fileName = lstFiles.SelectedItems[0].Text.ToString();
                 var pathTarget = TargetFolder +@"\" +fileName;
                 Log("pathTarget=" + pathTarget, EventLogEntryType.Information);
 
                 var pathSource = SourceUri + "/" + fileName;
                 Log("pathSource=" + pathSource, EventLogEntryType.Information);
 
 
                 if (File.Exists(pathTarget))
                 {
                     Log("pathTarget is deleted" + pathTarget, EventLogEntryType.Information);
 
                     File.Delete(pathTarget);
                 }
              
                var mailMergeOpenXml = new MailMergeOpenXml(Log);
 
 
                var source = new SourceWebDoc(pathSource);
                var target = new TargetLocalDoc(pathTarget);
 
                mailMergeOpenXml.FillData(source, target);
                Log("done disconnect", EventLogEntryType.Information);
 
            }
           
            return;
        }
 
      
        /// <summary>
        /// Open specified word document.
        /// </summary>
        static void OpenMicrosoftWord(string file)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = "WINWORD.EXE";
            startInfo.Arguments = file;
            Process.Start(startInfo);
        }
 
      
        void Log(string s, System.Diagnostics.EventLogEntryType e)
        {
            Console.WriteLine(s);
            ListViewItem item = new ListViewItem(s);
            item.Tag = s;
            item.Text = s;
 
            lstLog.Items.Add(item);
        }
 
        private void btnRefresh_Click(object sender, EventArgs e)
        {
            Refresh();
        }
 
        private void btnRefresh_Click_1(object sender, EventArgs e)
        {
            Refresh();
        }
 
        private void btnTarget_Click_1(object sender, EventArgs e)
        {
            if (lstTarget.SelectedItems != null && lstTarget.SelectedItems.Count == 1)
            {
                String tage = lstTarget.SelectedItems[0].Tag.ToString();
                OpenMicrosoftWord(tage);
            }
            else
                MessageBox.Show("יש בחוק קובץ אחד בלבד ביעד");
        }
 
        private void btnOpen_Click_1(object sender, EventArgs e)
        {
            if (lstFiles.SelectedItems != null && lstFiles.SelectedItems.Count > 0)
            {
                String tage = lstFiles.SelectedItems[0].Tag.ToString();
                OpenMicrosoftWord(tage);
            }
            else
                MessageBox.Show("יש בחוק קובץ אחד בלבד במקור -MERGE");
        }
 
        private void button1_Click(object sender, EventArgs e)
        {
            //var mailMergeOpenXml = new MailMergeOpenXml(Log);
            //mailMergeOpenXml.Cc("http://mvmentcrmqa:8080//Doctemplates/output/disconnect/2.DOCM");
 
        }
    }
}


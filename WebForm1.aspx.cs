using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace _3e_sample
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        private TXTextControl.DocumentServer.MailMerge mailMerge1;
        private System.ComponentModel.IContainer components;
        private TXTextControl.ServerTextControl serverTextControl1;
    
        protected void Page_Load(object sender, EventArgs e)
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.serverTextControl1 = new TXTextControl.ServerTextControl();
            this.mailMerge1 = new TXTextControl.DocumentServer.MailMerge(this.components);
            // 
            // serverTextControl1
            // 
            this.serverTextControl1.SpellChecker = null;
            // 
            // mailMerge1
            // 
            this.mailMerge1.TextComponent = this.serverTextControl1;

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            // create a new sample data source
            DataTable dt = new DataTable();
            dt.Columns.Add("company");
            dt.Columns.Add("name");

            dt.Rows.Add(new string[] { "Text Control", "Peter Jackson" });
            dt.Rows.Add(new string[] { "Microsoft", "Sue Williamson" });
            dt.Rows.Add(new string[] { "Google", "Peter Gustavsson" });

            // load the template
            mailMerge1.LoadTemplate(Server.MapPath("template.docx"), TXTextControl.DocumentServer.FileFormat.WordprocessingML);

            // attached the DataRowMerged event
            mailMerge1.DataRowMerged += mailMerge1_DataRowMerged;

            // merge with the appenbd parameter "false"
            mailMerge1.Merge(dt, false);
        }

        void mailMerge1_DataRowMerged(object sender, TXTextControl.DocumentServer.MailMerge.DataRowMergedEventArgs e)
        {
            string htmlDocument = "";

            // load the merged records into a temporary ServerTextControl
            using (TXTextControl.ServerTextControl tx = new TXTextControl.ServerTextControl())
            {
                tx.Create();
                tx.Load(e.MergedRow, TXTextControl.BinaryStreamType.InternalUnicodeFormat);

                // save the results as HTML
                tx.Save(out htmlDocument, TXTextControl.StringStreamType.HTMLFormat);
            }

            // write the HTML to the browser
            Response.Write(htmlDocument);
            Response.Write("<hr />");
        }
    }
}
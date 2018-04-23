using System;
using System.Data;
using System.Windows.Forms;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace RichEditCalculatedField {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();

            richEditControl1.Options.MailMerge.DataSource = ProductsTable.CreateData();
            RestoreTemplate();
        }

        private void button1_Click(object sender, EventArgs e) {
            MailMergeOptions mailMergeOptions = richEditControl1.Document.CreateMailMergeOptions();
            mailMergeOptions.MergeMode = MergeMode.JoinTables;

            RichEditDocumentServer server = new RichEditDocumentServer();

            server.CalculateDocumentVariable += server_CalculateDocumentVariable;

            richEditControl1.Document.MailMerge(mailMergeOptions, server.Document);

            richEditControl1.LoadDocument("HeaderTemplate.rtf");
            richEditControl1.Document.AppendDocumentContent(server.Document.Range);
        }

        void server_CalculateDocumentVariable(object sender, CalculateDocumentVariableEventArgs e) {
            if (e.VariableName == "Prod") {
                int productId = -1;

                if (Int32.TryParse(e.Arguments[0].Value, out productId)) {
                    DataRow row = ((DataTable)richEditControl1.Options.MailMerge.DataSource).Rows.Find(productId);
                    int unitsInStock = Convert.ToInt32(row[e.Arguments[1].Value]);
                    decimal unitPrice = Convert.ToDecimal(row[e.Arguments[2].Value]);

                    e.Value = unitsInStock * unitPrice;
                    e.Handled = true;
                }
            }
        }

        #region Helper Methods
        private void button2_Click(object sender, EventArgs e) {
            RestoreTemplate();
        }

        void RestoreTemplate() {
            richEditControl1.LoadDocument("DetailTemplate.rtf");
            ShowFieldCodes();
        }

        private void ShowFieldCodes() {
            Document doc = richEditControl1.Document;
            doc.BeginUpdate();
            foreach (Field f in doc.Fields) f.ShowCodes = true;
            doc.EndUpdate();
        }
        #endregion Helper Methods
    }
}
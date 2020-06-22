using Dapper;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sage100c_SysLibreExt
{
    public partial class Form1 : Form
    {
        private dynamic fields;

        public Form1()
        {
            InitializeComponent();
            this.Text = "Informations libres étendues " + Program.paramTableName;

            string dbQuery = "SELECT * FROM " + Program.paramTableName + Program.paramWhereClause;
            var result = Program.db.Query(dbQuery);
            if (result.Count() == 0) {
                MessageBox.Show("No record found for this query:\n" + dbQuery);
                this.Close();
            } else if (result.Count() > 1) {
                MessageBox.Show("More than 1 record found for this query:\n" + dbQuery);
                this.Close();
            }
            IDictionary<string, object> values = result.First();

            dbQuery = "SELECT t.*, CASE WHEN c.user_type_id IN (231,239) THEN c.max_length/2 ELSE c.max_length END AS max_length"
                    + "  FROM [" + Program.dbTableConfig + "] t"
                    + "  JOIN sys.[columns] c ON c.[object_id] = OBJECT_ID(t.[CB_File]) AND c.[name] = t.[CB_Name]"
                    + " WHERE t.[CB_File] = '" + Program.paramTableName + "'";
            this.fields = Program.db.Query(dbQuery);
            int i = 0;
            foreach (var field in this.fields) {
                //Label
                Label label = new Label();
                this.panel.Controls.Add(label);
                label.Top = i * 25 + 8;
                label.Width = 145;
                label.TextAlign = ContentAlignment.MiddleRight;
                label.Text = field.CB_Name;

                //Textbox
                TextBox textbox = new TextBox();
                this.panel.Controls.Add(textbox);
                textbox.Top = i * 25 + 10;
                textbox.Left = 150;
                textbox.Width = 300;
                textbox.Text = values[field.CB_Name];
                textbox.Tag = field.CB_Name; //Store database field name in Taf so we can use it to generate UPDATE statement below
                textbox.MaxLength = field.max_length;
                if (field.IsMultiline) {
                    textbox.Multiline = true;
                    textbox.Height += 25; //Display 3 lines per default
                    i++; //Increment here because multiline textbox is higher
                }
                if (field.IsReadOnly) {
                    textbox.ReadOnly = true;
                }

                //TODO: Handle selectbox + regex validation or mask

                i++;
            }
        }

        private void btnOK_Click(object sender, EventArgs e) {
            //Save to database
            List<string> lstUpdates = new List<string>();
            foreach (var field in this.fields) {
                TextBox textbox = this.panel.Controls.OfType<TextBox>().FirstOrDefault(c => c.Tag == field.CB_Name);
                if (textbox.Modified) {
                    lstUpdates.Add(field.CB_Name + " = N'" + textbox.Text + "'");
                }
            }

            string dbQuery = lstUpdates.Count == 0 ? ""
                : "EXEC sys.sp_set_session_context @key = N'cbLockDisable', @value = 1;\n" //Bypass lock, compatible with sage100c-lock-rebuild
                + "UPDATE " + Program.paramTableName + " SET " + string.Join(", ", lstUpdates) + Program.paramWhereClause + "\n"
                + "EXEC sys.sp_set_session_context @key = N'cbLockDisable', @value = NULL;";
            try {
                Program.db.Execute(dbQuery);
            } catch (Exception ex) { //Handle exception (ex trigger)
                MessageBox.Show(ex.Message);
                return;
            }

            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e) {
            this.Close();
        }
    }
}

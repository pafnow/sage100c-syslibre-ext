using Dapper;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sage100c_SysLibreExt
{
    public class Contexte
    {
        public string Name;  //Name of the contexte
        public string Table; //Related SQL table name
        public string Key;   //Key field (passed as RecordId)
        public string Where; //Where clause to apply

        public Contexte(string name, string table, string key, string where = "") {
            this.Name = name;
            this.Table = table;
            this.Key = key;
            this.Where = where;
        }
    }

    static class Program {
        public static List<Contexte> allowedContextes = new List<Contexte>() {
            //new Contexte("Global", "", ""),
            new Contexte("Tiers", "F_COMPTET", "CT_Num"),
            new Contexte("SectionsAnalytiques", "F_COMPTEA", "CA_Num"),
            new Contexte("Banques", "F_BANQUE", "BQ_No"),
            new Contexte("Collaborateurs", "F_COLLABORATEUR", "CO_No"),
            new Contexte("Clients", "F_COMPTET", "CT_Num", "CT_Type = 0"),
            new Contexte("Articles", "F_ARTICLE", "AR_Ref"),
            new Contexte("DocumentsDesVentes", "F_DOCENTETE", "DO_Piece", "DO_Domaine = 0"),
            new Contexte("DocumentsDesAchats", "F_DOCENTETE", "DO_Piece", "DO_Domaine = 1"),
            new Contexte("DocumentsDesStocks", "F_DOCENTETE", "DO_Piece", "DO_Domaine = 2"),
            new Contexte("DocumentsInternes", "F_DOCENTETE", "DO_Piece", "DO_Domaine = 4"),
            //new Contexte("LignesDeDocument", "F_DOCLIGNE", "?"), //Cannot find Key information
            new Contexte("Ressources", "F_RESSOURCEPROD", "RP_Code"),
            new Contexte("Depots", "F_DEPOT", "DE_No"),
            //new Contexte("ProjetsDAffaire", "", ""),
            new Contexte("ProjetsDeFabrication", "F_PROJETFABRICATION", "PF_Num")
        };

        public static string paramGcmPath;
        public static string paramTableName;
        public static string paramWhereClause;
        public static SqlConnection db = null;
        public static string dbTableConfig = "ZZ1_SysLibreExt";

        /// Legacy function to read parameters from ini files
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        /// The main entry point for the application.
        [STAThread]
        static void Main(string[] args)
        {
            try {
                //Verif number of command line parameters
                if (args.Count() != 3) {
                    MessageBox.Show("This program should be called with exactly 3 parameters.\nSage100c_SysLibreExt.exe [gcm path] [contexte] [record identifier]\n" 
                        + JsonConvert.SerializeObject(args));
                    return;
                }

                //Verif contexte parameter
                Contexte contexte = allowedContextes.FirstOrDefault(c => c.Name == args[1]);
                if (contexte == null) {
                    MessageBox.Show("Unknown contexte parameter, verify external program configuration.\n" + JsonConvert.SerializeObject(args));
                    return;
                }
                paramGcmPath = args[0];
                paramTableName = contexte.Table;
                paramWhereClause = " WHERE " + contexte.Key + " = '" + args[2] + "'" + (contexte.Where == null ? "" : " AND " + contexte.Where);

                //Gather db information from gcm file
                StringBuilder strTemp = new StringBuilder(255);
                GetPrivateProfileString("CBASE", "ServeurSQL", "", strTemp, 255, paramGcmPath);
                string dbConnectionString = "Server=" + strTemp.ToString() + ";Database=" + System.IO.Path.GetFileNameWithoutExtension(paramGcmPath).ToString() + ";Trusted_Connection=True;";
                db = new SqlConnection(dbConnectionString);

                //Display form
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            } finally {
                if (db != null) db.Close();
            }
        }
    }
}

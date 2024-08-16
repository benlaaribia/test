using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Windows.Forms;
namespace WebApplication1
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        public SqlConnection cn = new SqlConnection("server=192.168.1.252;database=CAPRAMSAGE;user id=SAGE;password=Lun@2020;");
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                if (Request["ETAT"].Equals("R_CA_SAP"))
                {
                    DataSet1 dd = new DataSet1();
                    SqlCommand cmd = new SqlCommand("R_CA_SAP", cn);
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("ANNEE", Request["ANNEE"]);

                    new SqlDataAdapter(cmd).Fill(dd, "etatCASAP");
                    if (dd.Tables["etatCASAP"].Rows.Count == 0)
                    {
                         ClientScript.RegisterStartupScript(GetType(), "Avertissement", "alertme('Erreur de traitement! Aucune ligne trouvée ...');", true);
                        
                    }
                    else
                    {
                        Page.Header.Title = "Visualisation";
                        Page.Title = "Visualisation";
                        ReportDocument rd = new ReportDocument();
                        string path = Server.MapPath("./ETAT/EtatCASAP.rpt");
                        rd.Load(@path);
                        rd.SetDataSource((DataSet)dd);
                        rd.ExportToHttpResponse(ExportFormatType.PortableDocFormat, Response, false, "BC " + Request["ANNEE"]);
                        Page.Header.Title = "Visualisation";
                        Page.Title = "Visualisation";
                        Response.End();
                    }
                }
                else
                {
                  
                    ClientScript.RegisterStartupScript(GetType(), "Avertissement", "alertme(' Aucune ligne trouvée ...');", true);
                }
          
            }
            catch (Exception ex)
            {
                ClientScript.RegisterStartupScript(GetType(), "Avertissement", "alertme('Erreur !" + ex.Message.Replace(Environment.NewLine, "").Replace("'", " ") + "');", true);
              
            }
        }
    }
}
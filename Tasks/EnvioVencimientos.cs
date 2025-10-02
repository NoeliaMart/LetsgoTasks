using NLog;
using NortConsultingTasks.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace NortConsultingTasks.Tasks
{
    public class EnvioVencimientos
    {
        public static void Ejecuta(Logger Logger)
        {
            string sql = "SELECT OINV2.*, INV6.*, OCPR.\"E_MailL\", ADP2.\"EmailSbj\", ADP2.\"EmailBody\", CASE WHEN OINV.\"DocEntry\" = OINV2.\"DocEntry\" THEN 'Y' ELSE 'N' END AS \"Actualiza\" " +
                        " FROM OINV INNER JOIN INV6 ON OINV.\"DocEntry\" = INV6.\"DocEntry\" " +
                        " INNER JOIN OCRD ON OINV.\"CardCode\" = OCRD.\"CardCode\" " +
                        " INNER JOIN OCPR ON OINV.\"CardCode\" = OCPR.\"CardCode\" AND OINV.\"CntctCode\" = OCPR.\"CntctCode\" AND COALESCE(OCPR.\"E_MailL\", '') <> '' " +
                        " INNER JOIN ADP2 ON ADP2.\"RptType\" = 2 " +
                        " INNER JOIN OINV OINV2 ON OINV.\"CardCode\" = OINV2.\"CardCode\" AND OINV2.\"DocTotal\" > OINV2.\"PaidToDate\" AND OINV2.\"DocDueDate\" <= INV6.\"DueDate\" " +
                        " WHERE OCRD.\"U_NC_ENVREC\" = 'Y' AND COALESCE(INV6.\"U_NC_RECENVIADO\", 'N') = 'N' AND INV6.\"Status\" = 'O' AND OINV.\"CANCELED\" = 'N' AND INV6.\"InsTotal\" - OINV.\"PaidToDate\" > 0 AND INV6.\"DueDate\" = ADD_DAYS(CURRENT_DATE, 15) ";

            System.Data.DataTable dt = cBBDD.ExecDBQuery(sql);
            SAPbobsCOM.Company company = cDIAPI.ConnectionCompany();

            string errMsg = string.Empty;
            int iFactura = 0;
            string to = string.Empty;
            string cardCode = string.Empty;
            string subject = string.Empty;
            string body = string.Empty;
            string sErr = string.Empty;
            String path = string.Empty;
            string cc = Properties.Settings.Default.CC;


            var facturasCliente = dt.AsEnumerable()
            .GroupBy(dr => dr["CardCode"].ToString())
            .ToDictionary(g => g.Key, g => g.ToList());

            foreach (var kvp in facturasCliente)
            {
                cardCode = kvp.Key;
                List<DataRow> facturas = kvp.Value;

                List<DataRow> facturasDistinct = facturas
                .GroupBy(dr => dr["DocEntry"].ToString())
                .Select(grp => grp.First())
                .ToList();

                try
                {
                    to = facturas.First()["E_MailL"].ToString();
                    subject = facturas.First()["EmailSbj"].ToString();
                    subject = ReemplazarVariables(subject, facturas.First());
                    body = facturas.First()["EmailBody"].ToString();
                    //body = ReemplazarVariables(body, facturas.First());

                    List<string> attachments = new List<string>();

                    //"cobros@nortconsulting.com"

                    foreach (DataRow dr in facturasDistinct)
                    {
                        iFactura = (int)dr["DocEntry"];

                        path = System.IO.Path.GetTempPath() + "Factura_" + iFactura.ToString() + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                        
                        Dictionary<string, object> parametros = new Dictionary<string, object>();
                        parametros.Add("Dockey@", iFactura);
                        parametros.Add("ObjectId@", 13);
                        cDIAPI.ExportPDFCrystalReport(Logger, Properties.Settings.Default.ReportFactura, parametros, path);

                        Logger.Error("Factura exportada correctamente: " + path);
                        attachments.Add(path);
                    }


                    if (cDIAPI.SendEmailViaMessageService(Logger, company, cardCode, to, cc, subject, body, out sErr, attachments.ToArray()))
                    {
                        //Actualizamos el recordatorio
                        var facturasMarcadas = facturas
                                                .Where(dr => dr["Actualiza"] != DBNull.Value
                                                          && dr["Actualiza"].ToString() == "Y")
                                                .ToList();

                        ActualizaRecordatorio(facturasMarcadas);

                        Logger.Error("Mensaje enviado a " + to);
                    }
                    else
                    {
                        Logger.Error("Error enviando mensaje a " + to + ": " + sErr);
                    }




                }
                catch (Exception ex)
                {
                    Logger.Error(ex.Message);
                }


            }

        }


        static string ReemplazarVariables(string texto, DataRow row)
        {
            var matches = Regex.Matches(texto, @"@(\w+)@");

            foreach (Match match in matches)
            {
                string variable = match.Groups[1].Value;

                if (row.Table.Columns.Contains(variable))
                {
                    string valor = row[variable]?.ToString() ?? "";
                    texto = texto.Replace(match.Value, valor);
                }
            }

            return texto;
        }

        static void ActualizaRecordatorio(List<DataRow> facturas)
        {
            int iDocEntry = 0;
            int iInstlmntID = 0;

            foreach (var dr in facturas)
            {
                iDocEntry = (int)dr["DocEntry"];
                iInstlmntID = int.Parse(dr["InstlmntID"].ToString());

                cBBDD.ExecNonQuery(String.Format("UPDATE INV6 SET \"U_NC_RECENVIADO\" = 'Y' WHERE \"ObjType\" = 13 AND \"DocEntry\" = {0} AND \"InstlmntID\" = {1}", iDocEntry, iInstlmntID));
            }
        }
    }
}

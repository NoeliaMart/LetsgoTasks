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
    public class GeneraPersonasContacto
    {
        public static void Ejecuta(Logger Logger)
        {
            string sql = "SELECT OCPR.\"CardCode\", OCPR.\"Name\", OCPR.\"E_MailL\" FROM OCPR INNER JOIN OCRD ON OCPR.\"CardCode\" = OCRD.\"CardCode\" WHERE OCRD.\"CardType\" = 'C' AND OCPR.\"E_MailL\" LIKE '%;%'"; string serr = string.Empty;
            DataTable dt = cBBDD.ExecDBQuery(sql);
            SAPbobsCOM.Company company = cDIAPI.ConnectionCompany();

            string errMsg = string.Empty;

            foreach (DataRow dr in dt.Rows)
            {
                try
                {
                    if (cDIAPI.CreaPersonaContacto(company, dr["CardCode"].ToString(), dr["Name"].ToString(), dr["E_MailL"].ToString(), false, ref serr))
                    {
                        Logger.Info("Personas creadas: " + cBBDD.ToSQL(dr["CardCode"]).ToString() + " " + cBBDD.ToSQL(dr["Name"]).ToString());
                    }
                    else
                    {
                        Logger.Info("Error creando personas: " + cBBDD.ToSQL(dr["CardCode"]).ToString() + " " + cBBDD.ToSQL(dr["Name"]).ToString() + ". " + serr);
                    }

                }
                catch (Exception ex)
                {
                    Logger.Error(ex.Message);
                }


            }


            sql = "SELECT \"CardCode\", \"CardName\", \"E_Mail\" FROM OCRD WHERE \"CardType\" = 'C' AND \"E_Mail\" IS NOT NULL AND \"E_Mail\" <> ''";
            dt = cBBDD.ExecDBQuery(sql);

            foreach (DataRow dr in dt.Rows)
            {
                try
                {
                    string cardCode = dr["CardCode"].ToString();
                    string cardName = dr["CardName"].ToString();
                    string emailRaw = dr["E_Mail"].ToString();

                    // Separar correos por comas y quitar espacios en blanco
                    string[] emails = emailRaw.Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                  .Select(e => e.Trim())
                                  .Where(e => !string.IsNullOrEmpty(e))
                                  .ToArray();

                    foreach (string email in emails)
                    {
                        if (!ExisteEmailEnOCPR(cardCode, email))
                        {
                            string nombrePersona = GenerarNombreUnico(cardCode, cardName);

                            try
                            {
                                if (cDIAPI.CreaPersonaContacto(company, cardCode, nombrePersona, email, true, ref serr))
                                {
                                    Logger.Info($"Persona creada: {cBBDD.ToSQL(cardCode)} {cBBDD.ToSQL(nombrePersona)}");
                                }
                                else
                                {
                                    Logger.Info($"Error creando persona: {cBBDD.ToSQL(cardCode)} {cBBDD.ToSQL(nombrePersona)}. {serr}");
                                }
                            }
                            catch (Exception ex)
                            {
                                Logger.Error(ex.Message);
                            }
                        }
                    }

                }
                catch (Exception ex)
                {
                    Logger.Error(ex.Message);
                }

            }

        }

        public static bool ExisteEmailEnOCPR(string cardCode, string email)
        {
            string sql = $"SELECT 1 FROM OCPR WHERE \"CardCode\" = '{cardCode}' AND \"E_MailL\" = '{email}'";
            DataTable dt = cBBDD.ExecDBQuery(sql);
            return dt.Rows.Count > 0;
        }

        public static string GenerarNombreUnico(string cardCode, string cardName)
        {
            string sql = $"SELECT COUNT(*) FROM OCPR WHERE \"CardCode\" = '{cardCode}' AND \"Name\" LIKE '{cardName}%'";
            int count = Convert.ToInt32(cBBDD.ExecDBQuery(sql).Rows[0][0]);

            return count == 0 ? cardName : $"{cardName} {count + 1}";
        }
    }
}

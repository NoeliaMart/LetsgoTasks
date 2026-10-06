using LetsGoTasks.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace LetsGoTasks.Tasks
{
    public class ActualizaCambioMoneda
    {
        public static void Ejecuta()
        {
            string empresas = Properties.Settings.Default.EmpresasCambioMoneda;
            Dictionary<string, double> cambios = ObtenerCambiosBCE();
            SAPbobsCOM.Company company = null;


            foreach (string companyDB in empresas.Split(',').Select(x => x.Trim()).Where(x => x != ""))
            {
                try
                {
                    company = cDIAPI.ConnectionCompany(companyDB);

                    string monedaLocal = ObtenerMonedaLocal(company.CompanyDB);

                    ActualizarMonedas(company, monedaLocal, cambios);

                    Console.WriteLine($"{companyDB}: tipos de cambio actualizados");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"{companyDB}: {ex.Message}");
                }
                finally
                {
                    if (company != null && company.Connected)
                        company.Disconnect();
                }
            }

        }

        private static string ObtenerMonedaLocal(string companyDB)
        {
            string sql = $"SELECT \"MainCurncy\" FROM OADM";
            var dt = cBBDD.ExecDBQuery(companyDB, sql);
            return dt.Rows.Count > 0 ? dt.Rows[0]["MainCurncy"].ToString() : "";
        }

        private static void ActualizarMonedas(SAPbobsCOM.Company company, string monedaLocal, Dictionary<string, double> cambios)
        {
            string sql = "SELECT \"CurrCode\" FROM OCRN WHERE \"Locked\" = 'N'";
            var dt = cBBDD.ExecDBQuery(company.CompanyDB, sql);

            foreach (System.Data.DataRow row in dt.Rows)
            {
                string moneda = row["CurrCode"].ToString();

                if (moneda != monedaLocal && cambios.ContainsKey(moneda))
                {
                    string err = "";

                    double tipoCambio = CalcularCambio(moneda, monedaLocal, cambios);

                    if (!cDIAPI.ActualizarMoneda(company, moneda, tipoCambio, DateTime.Today, ref err))
                        Log.InfoFichero("ActualizaCambioMoneda", $"Error {company.CompanyDB}: {err}");
                   
                }
            }
        }

        private static double CalcularCambio(string moneda, string monedaLocal, Dictionary<string, double> cambios) =>
         monedaLocal == "EUR" ? 1 / cambios[moneda] :
         moneda == "EUR" ? cambios[monedaLocal] :
         cambios[monedaLocal] / cambios[moneda];


        public static Dictionary<string, double> ObtenerCambiosBCE()
        {
            Dictionary<string, double> cambios = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

            using (HttpClient client = new HttpClient())
            {
                string xml = client.GetStringAsync("https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml").GetAwaiter().GetResult();

                XDocument doc = XDocument.Parse(xml);

                XNamespace cube ="http://www.ecb.int/vocabulary/2002-08-01/eurofxref";

                foreach (XElement item in doc.Descendants(cube + "Cube").Where(x => x.Attribute("currency") != null))
                {
                    string moneda = item.Attribute("currency").Value;

                    double cambio = double.Parse(item.Attribute("rate").Value,System.Globalization.CultureInfo.InvariantCulture);

                    cambios[moneda] = cambio;
                }

                cambios["EUR"] = 1;
            }

            return cambios;
        }
    }
}

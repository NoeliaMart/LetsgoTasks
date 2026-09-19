using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using LetsGoTasks.Utils;
using LetsGoTasks.Utils.Models;
using ClosedXML.Excel;

namespace LetsGoTasks.Tasks
{
    public class ProcesaCargaFacturas
    {
        public static void Ejecuta()
        {
            string folder = Properties.Settings.Default.CSVProvFolder;
            string procesados = Path.Combine(folder, "Procesados");
            string erroneos = Path.Combine(folder, "Erroneos");

            foreach (string fichero in Directory.GetFiles(folder, "*.xlsx"))
            {
                string nombre = Path.GetFileName(fichero);
                string companyDB = Path.GetFileNameWithoutExtension(fichero).Substring(7, nombre.LastIndexOf('_') - 7);
                string err = "", cardCode = "";
                List<string> lErrors = new List<string>();
                List<int> lDocumentos = new List<int>();
                List<string> lProv = new List<string>();
                int docEntry = 0;
                SAPbobsCOM.Company company = null;

                try
                {
                    company = cDIAPI.ConnectionCompany(companyDB);
                    var registros = LeerExcel(companyDB, fichero);

                    if (!registros.Any())
                        throw new Exception("El fichero no contiene registros válidos.");

                    if (company.InTransaction)
                        company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                    company.StartTransaction();

                    foreach (var factura in registros.GroupBy(r => new
                    {
                        NumAtCard = string.IsNullOrEmpty(r.NumAtCard) ? r.U_COD_Ocs : r.NumAtCard,
                        r.LicTradNum
                    }))
                    {
                        var r = factura.First();
                        cardCode = r.CardCode;

                        if (string.IsNullOrEmpty(cardCode))
                        {
                            if (!cDIAPI.CreaProveedor(company, r, ref err, ref cardCode))
                                lErrors.Add(err);
                            else
                                lProv.Add(cardCode);
                        }

                        if (!cDIAPI.CreaFacturaCompra(company, cardCode, factura.ToList(), ref err, ref docEntry))
                            lErrors.Add(err);
                        else
                            lDocumentos.Add(docEntry);
                    }

                    if (lErrors.Count > 0)
                    {
                        foreach (string sError in lErrors)
                            Log.Error(companyDB, "ProcesaCargaFacturasExcel",sError);

                        throw new Exception("Fichero con errores.");
                    }

                    if (company.InTransaction)
                        company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                    foreach (string sCardCode in lProv)
                        Log.Info(companyDB, "ProcesaCargaFacturasExcel",
                            $"Proveedor creado. CardCode: {sCardCode}", sCardCode, "OCRD");

                    foreach (int idocEntry in lDocumentos)
                        Log.Info(companyDB, "ProcesaCargaFacturasExcel",
                            $"Factura creada. Número: {GetDocNum(companyDB, idocEntry)}",
                            idocEntry.ToString(), "OPCH");

                    Directory.CreateDirectory(procesados);
                    File.Move(fichero, Path.Combine(procesados,
                        $"{Path.GetFileNameWithoutExtension(nombre)}_{DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(nombre)}"));

                    Log.Info(companyDB, "ProcesaCargaFacturasExcel", "Fichero procesado correctamente.");
                }
                catch (Exception ex)
                {
                    string mensaje = ex.Message;

                    if (company != null)
                    {
                        try
                        {
                            if (company.InTransaction)
                                company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        }
                        catch (Exception exRollback)
                        {
                            company.GetLastError(out int code, out string sapError);
                            mensaje += $" | ERROR ROLLBACK: {code} - {sapError} | {exRollback.Message}";
                        }
                    }

                    if (company == null)
                        Log.ErrorFichero("ProcesaCargaFacturasExcel", mensaje);
                    else
                        Log.Error(companyDB, "ProcesaCargaFacturasExcel", mensaje);

                    Directory.CreateDirectory(erroneos);
                    File.Move(fichero, Path.Combine(erroneos,
                        $"{Path.GetFileNameWithoutExtension(nombre)}_{DateTime.Now:yyyyMMddHHmmss}{Path.GetExtension(nombre)}"));
                }
                finally
                {
                    if (company != null)
                    {
                        if (company.Connected)
                            company.Disconnect();

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(company);
                    }
                }
            }
        }


        private static List<RegistroFactura> LeerExcel(string companyDB, string fichero)
        {
            var lista = new List<RegistroFactura>();
            string[] campos = Properties.Settings.Default.CSVProvColumns.Split(',').Select(x => x.Trim()).ToArray();

            using (var wb = new XLWorkbook(fichero))
                foreach (var fila in wb.Worksheet(1).RowsUsed().Skip(1))
                {
                    var r = new RegistroFactura
                    {
                        LicTradNum = V(fila, campos, "CIF"),
                        CardCode = ExisteProveedor(companyDB, V(fila, campos, "CIF")),
                        Series = GetSeries(companyDB),
                        CardName = V(fila, campos, "PROVEEDOR"),
                        Address = $"{V(fila, campos, "TIPO_VIA")} {V(fila, campos, "CALLE")} {V(fila, campos, "NUM_CALLE")}".Trim(),
                        City = V(fila, campos, "CIUDAD"),
                        ZipCode = V(fila, campos, "CODIGO_POSTAL"),
                        Country = V(fila, campos, "PAIS"),
                        PeymentMethodCode = V(fila, campos, "FORMA_PAGO"),
                        CardType = V(fila, campos, "CardType"),
                        NumAtCard = V(fila, campos, "NUM_FRA"),
                        DocDate = Fecha(fila, campos, "FECHA_FACTURA"),
                        DocDueDate = Fecha(fila, campos, "VTO_OC"),
                        TaxDate = Fecha(fila, campos, "TaxDate"),
                        AccountCode = V(fila, campos, "CUENTA_ANALITICA"),
                        LineTotal = double.Parse(V(fila, campos, "BASE")),
                        DocCurrency = V(fila, campos, "MONEDA"),
                        VatGroup = V(fila, campos, "COD_IVA"),
                        Description = V(fila, campos, "DESCRIPCION")?.Length > 100 ? V(fila, campos, "DESCRIPCION").Substring(0, 100) : V(fila, campos, "DESCRIPCION"),
                        U_Observaciones = V(fila, campos, "OBSERVACIONES"),
                        U_Linea_COD_Ocs = V(fila, campos, "LINEA_DE_OC"),
                        U_COD_Ocs = V(fila, campos, "COD_Ocs"),
                        ProjectCode = V(fila, campos, "COD_PROYECTO"),
                        CostingCode = V(fila, campos, "PLAZA"),
                        CostingCode2 = V(fila, campos, "MES_IMPUTACION"),
                        Archivo_adjunto = V(fila, campos, "ARCHIVO_ADJUNTO"),
                        WTCode = V(fila, campos, "COD_RET"),
                        State = V(fila, campos, "PROVINCIA")
                    };

                    ValidarRegistro(r);
                    lista.Add(r);
                }

            return lista;
        }

        private static string V(IXLRow fila, string[] campos, string campo)
        {
            int pos = Array.FindIndex(campos, x => string.Equals(x.Trim(), campo.Trim(), StringComparison.OrdinalIgnoreCase));
            if (pos < 0) throw new Exception($"El campo {campo} no existe en la configuración del Excel.");
            return fila.Cell(pos + 1).GetString().Trim();
        }


        private static DateTime Fecha(IXLRow fila, string[] campos, string campo) =>
            DateTime.ParseExact(V(fila, campos, campo), "yyyyMMdd",
                System.Globalization.CultureInfo.InvariantCulture);

        private static void ValidarRegistro(RegistroFactura r)
        {
            if (r.LicTradNum.Contains(".") || r.LicTradNum.Contains("-"))
                throw new Exception($"El CIF {r.LicTradNum} contiene caracteres no permitidos.");

            if (r.LineTotal.ToString().Contains("."))
                throw new Exception($"El importe {r.LineTotal} no debe tener punto decimal.");

            if (r.CardType != "cSupplier" && r.CardType != "cLid")
                throw new Exception($"Tipo de interlocutor no válido: {r.CardType}.");
        }

        public static string ExisteProveedor(string companyDB, string LicTradNum)
        {
            string sql = $"SELECT \"CardCode\" FROM OCRD WHERE \"CardType\" IN ('S','L') AND \"LicTradNum\" = {cBBDD.ToSQL(LicTradNum)}";
            var dt = cBBDD.ExecDBQuery(companyDB, sql);
            return dt.Rows.Count > 0 ? dt.Rows[0]["CardCode"].ToString() : "";
        }

        public static int GetSeries(string companyDB)
        {
            string sql = $"SELECT \"Series\" FROM NNM1 WHERE \"ObjectCode\" = 2 AND \"DocSubType\" = 'S' AND \"IsManual\" = 'N' AND \"Locked\" = 'N'";
            var dt = cBBDD.ExecDBQuery(companyDB, sql);
            return dt.Rows.Count > 0 ? int.Parse(dt.Rows[0]["Series"].ToString()) : 0;
        }

        public static string GetDocNum(string companyDB, int iDocEntry)
        {
            string sql = $"SELECT \"DocNum\" FROM OPCH WHERE \"DocEntry\" = {cBBDD.ToSQL(iDocEntry)}";
            var dt = cBBDD.ExecDBQuery(companyDB, sql);
            return dt.Rows.Count > 0 ? dt.Rows[0]["DocNum"].ToString() : "";
        }


    }
}
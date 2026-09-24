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

            foreach (string fichero in Directory.GetFiles(folder, "*.xlsx"))
            {
                string nombre = Path.GetFileName(fichero);
                string companyDB = Path.GetFileNameWithoutExtension(fichero)
                    .Substring(7, nombre.LastIndexOf('_') - 7);

                SAPbobsCOM.Company company = null;
                var lErrors = new List<string>();
                var lDocumentos = new List<int>();
                var lProv = new List<string>();

                try
                {
                    company = cDIAPI.ConnectionCompany(companyDB);
                    var registros = LeerExcel(companyDB, fichero);

                    if (!registros.Any())
                        throw new Exception("El fichero no contiene registros válidos.");

                    var grupos = registros.GroupBy(r => new
                    {
                        NumAtCard = string.IsNullOrEmpty(r.NumAtCard) ? r.U_COD_Ocs : r.NumAtCard,
                        r.LicTradNum
                    }).ToList();

                    // PROVEEDORES SIN TRANSACCIÓN
                    foreach (var grupo in grupos)
                    {
                        var r = grupo.First();

                        if (string.IsNullOrEmpty(r.CardCode))
                        {
                            string err = "", cardCode = "";

                            if (!cDIAPI.CreaProveedor(company, r, ref err, ref cardCode))
                                lErrors.Add(err);
                            else
                            {
                                r.CardCode = cardCode;
                                lProv.Add(cardCode);
                            }
                        }
                    }

                    if (lErrors.Any())
                        throw new Exception("Error creando proveedores.");

                    // FACTURAS CON TRANSACCIÓN
                    company.StartTransaction();

                    foreach (var grupo in grupos)
                    {
                        var r = grupo.First();
                        string err = "";
                        int docEntry = 0;

                        if (!cDIAPI.CreaFacturaCompra(
                            company, r.CardCode, grupo.ToList(), ref err, ref docEntry))
                        {
                            lErrors.Add(err);
                            break;
                        }

                        lDocumentos.Add(docEntry);
                    }

                    if (lErrors.Any())
                    {
                        if (company.InTransaction)
                            company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                        foreach (string sError in lErrors)
                            Log.Error(companyDB, nombre, sError);


                        throw new Exception("Error creando facturas.");
                    }

                    if (company.InTransaction)
                        company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                    foreach (string cardCode in lProv)
                        Log.Info(companyDB, nombre, $"Proveedor creado. CardCode: {cardCode}", cardCode, "OCRD");

                    foreach (int docEntry in lDocumentos)
                        Log.Info(companyDB, nombre,
                            $"Factura creada. Número: {GetDocNum(companyDB, docEntry)}",
                            docEntry.ToString(), "OPCH");

                    cUtils.MoverFicheroSFTP(
                        fichero, Properties.Settings.Default.SFTPFolderProcesado);

                    Log.Info(companyDB, nombre, "Fichero procesado correctamente.");
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
                        catch (Exception e)
                        {
                            company.GetLastError(out int code, out string sapError);
                            mensaje += $" | ERROR ROLLBACK: {code} - {sapError} | {e.Message}";
                        }
                    }

                    Log.Error(companyDB, nombre, mensaje);

                    cUtils.MoverFicheroSFTP(
                        fichero, Properties.Settings.Default.SFTPFolderError);
                }
                finally
                {
                    if (company != null)
                    {
                        if (company.Connected) company.Disconnect();
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
            {
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
                        County = V(fila, campos, "PROVINCIA"),
                        Country = V(fila, campos, "PAIS"),
                        Email = V(fila, campos, "EMAIL"),
                        PeymentMethodCode = V(fila, campos, "FORMA_PAGO"),
                        CardType = V(fila, campos, "CardType"),
                        NumAtCard = V(fila, campos, "NUM_FRA"),
                        DocDate = Fecha(fila, campos, "TaxDate"),
                        DocDueDate = Fecha(fila, campos, "VTO_OC"),
                        TaxDate = Fecha(fila, campos, "FECHA_FACTURA"),
                        AccountCode = V(fila, campos, "CUENTA_ANALITICA"),
                        LineTotal = double.Parse(V(fila, campos, "BASE").Replace(",", "."), System.Globalization.CultureInfo.InvariantCulture),
                        DocCurrency = V(fila, campos, "MONEDA"),
                        VatGroup = V(fila, campos, "COD_IVA"),
                        Description = V(fila, campos, "DESCRIPCION"),
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

                    if (r.Description?.Length > 100)
                        r.Description = r.Description.Substring(0, 100);

                    ValidarRegistro(r);
                    lista.Add(r);
                }
            }

            return lista;
        }

        private static string V(IXLRow fila, string[] campos, string campo)
        {
            int pos = Array.FindIndex(campos,
                x => string.Equals(x.Trim(), campo.Trim(), StringComparison.OrdinalIgnoreCase));

            if (pos < 0) throw new Exception($"El campo {campo} no existe en la configuración del Excel.");
            return fila.Cell(pos + 1).GetString().Trim();
        }

        private static DateTime Fecha(IXLRow fila, string[] campos, string campo)
        {
            string valor = V(fila, campos, campo);
            return DateTime.ParseExact(V(fila, campos, campo), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
        }
            

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
            string sql = "SELECT \"Series\" FROM NNM1 WHERE \"ObjectCode\" = 2 AND \"DocSubType\" = 'S' AND \"IsManual\" = 'N' AND \"Locked\" = 'N'";
            var dt = cBBDD.ExecDBQuery(companyDB, sql);
            return dt.Rows.Count > 0 ? int.Parse(dt.Rows[0]["Series"].ToString()) : 0;
        }

        public static string GetDocNum(string companyDB, int docEntry)
        {
            string sql = $"SELECT \"DocNum\" FROM OPCH WHERE \"DocEntry\" = {cBBDD.ToSQL(docEntry)}";
            var dt = cBBDD.ExecDBQuery(companyDB, sql);
            return dt.Rows.Count > 0 ? dt.Rows[0]["DocNum"].ToString() : "";
        }
    }
}
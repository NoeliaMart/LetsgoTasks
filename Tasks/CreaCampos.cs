using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using LetsGoTasks.Utils;
using SAPbobsCOM;

namespace LetsGoTasks.Tasks
{
    public class CreaCampos
    {
        public static void Ejecuta(string companyDB)
        {
            try
            {
                string err = string.Empty;

                CreaTabla(companyDB,"LTG_LOG", "Log", BoUTBTableType.bott_NoObject);
                CreaCampo(companyDB, "FECHA", "Fecha", 20, BoFieldTypes.db_Date, BoFldSubTypes.st_None, "@LTG_LOG", ref err);
                CreaCampo(companyDB,"AUTOR", "Autor", 100, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, "@LTG_LOG", ref err);
                CreaCampo(companyDB,"DESCRIPCION", "Descripción", 5000, BoFieldTypes.db_Memo, BoFldSubTypes.st_None, "@LTG_LOG", ref err);
                CreaCampo(companyDB,"ID", "Id", 50, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, "@LTG_LOG", ref err);
                CreaCampo(companyDB,"TABLA", "Tabla", 50, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, "@LTG_LOG", ref err);
                CreaCampo(companyDB,"TIPO", "Tipo", 20, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, "@LTG_LOG", ref err);

                CreaCampo(companyDB, "COD_Ocs", "COD_Ocs", 20, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, "PCH1", ref err);
                CreaCampo(companyDB, "Linea_COD_Ocs", "Linea COD_Ocs", 20, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, "PCH1", ref err);
                CreaCampo(companyDB, "Observaciones", "Observaciones", 5000, BoFieldTypes.db_Memo, BoFldSubTypes.st_None, "PCH1", ref err);
                CreaCampo(companyDB, "Archivo_ad", "Adjunto", 5000, BoFieldTypes.db_Memo, BoFldSubTypes.st_Link, "PCH1", ref err);

                Console.WriteLine("Creación terminada correctamente");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static int CreaCampo(string companyDB, string name, string description, int size, SAPbobsCOM.BoFieldTypes fldtype, SAPbobsCOM.BoFldSubTypes fldsubtype, string tablename, ref string err,
          string linkedTable = "", bool validvalue = false)
        {

            int i = 0;

            SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;

            try
            {
                if (!cBBDD.ExisteCampo(companyDB,tablename, name))
                {
                    SAPbobsCOM.Company company = cDIAPI.ConnectionCompany(companyDB);

                    oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                    oUserFieldsMD.Name = name;
                    oUserFieldsMD.Type = fldtype;
                    oUserFieldsMD.Description = description;
                    oUserFieldsMD.Size = size;
                    oUserFieldsMD.SubType = fldsubtype;
                    oUserFieldsMD.TableName = tablename;
                    oUserFieldsMD.LinkedTable = linkedTable;

                    i = oUserFieldsMD.Add();
                    if (i != 0)
                        throw new Exception("Error creando campo " + name + ": cod. " + i.ToString() + " - " + company.GetLastErrorDescription());
                }


                return i;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                if (oUserFieldsMD != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);

                    GC.WaitForPendingFinalizers();

                    GC.Collect();
                }

            }


        }

        public static int CreaTabla(string companyDB, string TableName, string TableDescription, SAPbobsCOM.BoUTBTableType TableType)
        {
            Int32 i = 0;
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;

            try
            {
                if (!cBBDD.ExisteTabla(companyDB,TableName))
                {
                    SAPbobsCOM.Company company = cDIAPI.ConnectionCompany(companyDB);

                    oUserTablesMD = (SAPbobsCOM.UserTablesMD)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

                    oUserTablesMD.TableName = TableName;
                    oUserTablesMD.TableDescription = TableDescription;
                    oUserTablesMD.TableType = TableType;

                    i = oUserTablesMD.Add();

                    if (i != 0)
                        throw new Exception("Error creando tabla " + TableName + ": cod. " + i.ToString() + " - " + company.GetLastErrorDescription());
                }


                return i;
            }
            catch (Exception ex)
            {

                throw new Exception(ex.Message);
            }
            finally
            {
                if (oUserTablesMD != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);

                    GC.WaitForPendingFinalizers();

                    GC.Collect();
                }
            }


        }

       

    }
}

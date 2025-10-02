using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using NLog;
using NortConsultingTasks.Utils;
using SAPbobsCOM;

namespace NortConsultingTasks.Tasks
{
    public class CreaCampos
    {
        public static void Ejecuta(Logger Logger)
        {
            try
            {
                string err = string.Empty;

                CreaCampo("NC_RECENVIADO", "Recordatorio enviado", 1, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, "INV6", ref err);
                Logger.Error("Campo creado");
            }
            catch (Exception ex)
            {

                Logger.Error(ex.Message);
            }
        }

        public static int CreaCampo(string name, string description, int size, SAPbobsCOM.BoFieldTypes fldtype, SAPbobsCOM.BoFldSubTypes fldsubtype, string tablename, ref string err,
          string linkedTable = "", bool validvalue = false)
        {

            int i = 0;

            SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;

            try
            {
                if (!cBBDD.ExisteCampo(tablename, name))
                {
                    SAPbobsCOM.Company company = cDIAPI.ConnectionCompany();

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
    }
}

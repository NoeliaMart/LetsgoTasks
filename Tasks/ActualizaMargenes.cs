using NLog;
using NortConsultingTasks.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace NortConsultingTasks.Tasks
{
    public class ActualizaMargenes
    {
        public static void Ejecuta(Logger Logger)
        {
            string sql = "SELECT \"DocEntry\", \"LineNum\", \"Beneficio\", \"PrecioCompra\", \"PrecioCompraTot\" " +
                        "FROM V_ActMargenes ";

            System.Data.DataTable dt = cBBDD.ExecDBQuery(sql);

            string errMsg = string.Empty;

            foreach (DataRow dr in dt.Rows)
            {
                try
                {
                    Logger.Info("Linea actualizada: " + cBBDD.ToSQL(dr["DocEntry"]).ToString() + " " + cBBDD.ToSQL(dr["LineNum"]).ToString());


                    cBBDD.ExecNonQuery(String.Format("UPDATE RDR1 SET \"GrossBuyPr\" = {0} WHERE \"DocEntry\" = {1} AND \"LineNum\" = {2}", cBBDD.ToSQL(dr["PrecioCompra"]), cBBDD.ToSQL(dr["DocEntry"]), cBBDD.ToSQL(dr["LineNum"])));
                    cBBDD.ExecNonQuery(String.Format("UPDATE RDR1 SET \"GPTtlBasPr\" = \"GrossBuyPr\" * \"Quantity\", \"GrssProfit\" = (\"Price\" - \"GrossBuyPr\") * \"Quantity\" WHERE \"DocEntry\" = {0}", cBBDD.ToSQL(dr["DocEntry"])));

                    cBBDD.ExecNonQuery(String.Format("UPDATE ORDR" +
                              " SET \"GrosProfit\" = (SELECT SUM(\"GrssProfit\") FROM RDR1 L WHERE L.\"DocEntry\" = ORDR.\"DocEntry\")" +
                              "   , \"GrosProfSy\" = (SELECT SUM(\"GrssProfit\") FROM RDR1 L WHERE L.\"DocEntry\" = ORDR.\"DocEntry\")" +
                              " WHERE \"DocEntry\" = {0}", cBBDD.ToSQL(dr["DocEntry"])));

                }
                catch (Exception ex)
                {
                    Logger.Error(ex.Message);
                }


            }

        }
    }
}

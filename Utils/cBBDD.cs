using System.Data.Common;
using System.Data;
using System;
using System.Configuration;

namespace NortConsultingTasks.Utils
{
    public class cBBDD
    {
        public static System.Data.Common.DbProviderFactory SetDbProvider(string bbddType)
        {
            return System.Data.Common.DbProviderFactories.GetFactory(bbddType);
        }
        public static DbConnection SetDbConnection()
        {
            DbConnection connection;
            string serverIP = Properties.Settings.Default.ServerIP;
            string BBDD = Properties.Settings.Default.BBDD;
            string userName = Properties.Settings.Default.UserBBDD;
            string Passw =  EncryptionHelper.DecryptPass(Properties.Settings.Default.PwdBBDD);
            string connStr;
            DbProviderFactory odbProvider = SetDbProvider(Properties.Settings.Default.DbProviderFactory);


#if DEBUG
            connStr = "SERVER=" + serverIP + ";DATABASE=" + BBDD + ";UID=" + userName + ";PWD=" + Passw;
#else
            connStr = "Server=" + serverIP + ";UserID=" + userName + ";Password=" + Passw + ";current schema=" + BBDD;
#endif

            connection = odbProvider.CreateConnection();

            try
            {
                connection.ConnectionString = connStr;
                connection.Open();
            }
            catch (Exception ex)
            {
                throw new Exception("Error conectando BBDD: " + ex.Message);
            }

            return connection;

        }


        public static object ExecScalarQuery(string sql)
        {
            DbConnection connection = null;

            try
            {
                connection = SetDbConnection();
                object result;

                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }

                using (DbCommand command = connection.CreateCommand())
                {
                    command.CommandText = sql;
                    command.CommandType = CommandType.Text;

                    result = command.ExecuteScalar();
                }

                return result;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                    connection.Dispose();
                    connection = null;
                }
            }

        }

        public static void ExecNonQuery(string sql, string companyDB = "")
        {
            DbConnection connection = null;

            try
            {
                connection = SetDbConnection();
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }

                using (DbCommand command = connection.CreateCommand())
                {
                    command.CommandText = sql;
                    command.CommandType = CommandType.Text;

                    command.ExecuteNonQuery();
                }

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                    connection.Dispose();
                    connection = null;
                }
            }

        }

        public static string ToSQL(object value)
        {
            if (value.GetType() == System.Type.GetType("System.String"))
            {
                return "'" + value.ToString().Replace("'", "''") + "'";
            }
            else if (value.GetType() == System.Type.GetType("System.Date") || value.GetType() == System.Type.GetType("System.DateTime"))
            {
                return "'" + (Convert.ToDateTime(value)).ToString("o").Replace("0000", "") + "'";
            }
            else if (value.GetType() == System.Type.GetType("System.Double") || value.GetType() == System.Type.GetType("System.Decimal"))
            {
                return value.ToString().Replace(",", ".");
            }
            else
            {
                return value.ToString();
            }

        }

        public static DataTable ExecDBQuery(string sql)
        {
            DbConnection connection = null;
            DbProviderFactory dbProvider;
            DataTable dt = new DataTable();


            try
            {
                connection = SetDbConnection();
                dbProvider = SetDbProvider(Properties.Settings.Default.DbProviderFactory);

                using (DbCommand command = connection.CreateCommand())
                {
                    command.CommandText = sql;
                    command.CommandType = CommandType.Text;

                    using (DbDataAdapter adapter = dbProvider.CreateDataAdapter())
                    {
                        adapter.SelectCommand = command;
                        adapter.Fill(dt);
                    }
                }
            }
            catch (Exception ex)
            {

                throw new Exception(ex.Message);
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                    connection.Dispose();
                    connection = null;
                }
            }

            return dt;
        }

        public static bool ExisteCampo(string Tablename, string Fieldname)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            //cUtilsSAP.GetApplication().SetStatusBarMessage("Actualizando OFs", SAPbouiCOM.BoMessageTime.bmt_Short, false);
            try
            {
                string sql;

#if DEBUG
                sql = "select * from CUFD WITH (NOLOCK) where upper(\"TableID\") = '" + Tablename + "' AND upper(\"AliasID\") = '" + Fieldname + "'";
#else
                sql = "select * from CUFD where trim(upper(\"TableID\")) = '" + Tablename + "' AND trim(upper(\"AliasID\")) = '" + Fieldname + "'";
#endif

                System.Data.DataTable dt = cBBDD.ExecDBQuery(sql);

                return dt.Rows.Count > 0;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                if (oRecordSet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

    }
}

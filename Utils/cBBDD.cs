using System.Data.Common;
using System.Data;
using System;
using System.Configuration;

namespace LetsGoTasks.Utils
{
    public class cBBDD
    {
        public static System.Data.Common.DbProviderFactory SetDbProvider(string bbddType)
        {
            return System.Data.Common.DbProviderFactories.GetFactory(bbddType);
        }
        public static DbConnection SetDbConnection(string companyDB)
        {
            DbConnection connection;
            string serverIP = Properties.Settings.Default.ServerIP;
            string userName = Properties.Settings.Default.UserBBDD;
            string Passw =  EncryptionHelper.DecryptPass(Properties.Settings.Default.PwdBBDD);
            string connStr;
            DbProviderFactory odbProvider = SetDbProvider(Properties.Settings.Default.DbProviderFactory);


#if DEBUG
            connStr = "SERVER=" + serverIP + ";DATABASE=" + companyDB + ";UID=" + userName + ";PWD=" + Passw;
#else
            connStr = "Server=" + serverIP + ";UserID=" + userName + ";Password=" + Passw + ";current schema=" + companyDB;
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


        public static object ExecScalarQuery(string companyDB, string sql)
        {
            DbConnection connection = null;

            try
            {
                connection = SetDbConnection(companyDB);
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

        public static void ExecNonQuery(string companyDB, string sql)
        {
            DbConnection connection = null;

            try
            {
                connection = SetDbConnection(companyDB);
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

        public static DataTable ExecDBQuery(string companyDB,string sql)
        {
            DbConnection connection = null;
            DbProviderFactory dbProvider;
            DataTable dt = new DataTable();


            try
            {
                connection = SetDbConnection(companyDB);
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

        public static bool ExisteCampo(string companyDB, string Tablename, string Fieldname)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            //cUtilsSAP.GetApplication().SetStatusBarMessage("Actualizando OFs", SAPbouiCOM.BoMessageTime.bmt_Short, false);
            try
            {
                string sql;

#if DEBUG
                sql = "select * from CUFD WITH (NOLOCK) where upper(\"TableID\") = '" + Tablename + "' AND upper(\"AliasID\") = '" + Fieldname + "'";
#else
                sql = "select * from CUFD where trim(upper(\"TableID\")) = '" + Tablename + "' AND \"AliasID\" = '" + Fieldname + "'";
#endif

                System.Data.DataTable dt = cBBDD.ExecDBQuery(companyDB, sql);

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

        public static bool ExisteTabla(string companyDB, string TableName)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                string sql;

#if DEBUG
                sql = "select * from OUTB WITH (NOLOCK) where upper(\"TableName\") = '" + TableName + "'";
#else
                sql = "select * from OUTB where trim(upper(\"TableName\")) = '" + TableName + "'";
#endif


                System.Data.DataTable dt = cBBDD.ExecDBQuery(companyDB,sql);

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

using System;
using System.IO;

namespace LetsGoTasks.Utils
{
    public static class Log
    {
        public static void Info(string companyDB, string descripcion)
        {
            Guardar(companyDB, "", "INFO", descripcion);
        }

        public static void Info(string companyDB, string autor, string descripcion, string id = "", string tabla = "")
        {
            Guardar(companyDB, autor, "INFO", descripcion, id, tabla);
        }

        public static void Error(string companyDB, string autor, string descripcion, string id = "", string tabla = "")
        {
            Guardar(companyDB, autor, "ERROR", descripcion, id, tabla);
        }

        private static void Guardar(string companyDB, string autor, string tipo, string descripcion, string id = "", string tabla = "")
        {
            try
            {
                string code = DateTime.Now.ToString("yyyyMMddHHmmssfff");

                string sql = $"INSERT INTO \"@LTG_LOG\" (\"Code\", \"Name\", \"U_FECHA\",\"U_AUTOR\",\"U_DESCRIPCION\",\"U_ID\",\"U_TABLA\",\"U_TIPO\") " +
                             $"VALUES ({cBBDD.ToSQL(code)}, {cBBDD.ToSQL(code)}, CURRENT_TIMESTAMP,{cBBDD.ToSQL(autor)},{cBBDD.ToSQL(descripcion)},{cBBDD.ToSQL(id)},{cBBDD.ToSQL(tabla)},{cBBDD.ToSQL(tipo)})";

                cBBDD.ExecNonQuery(companyDB, sql);
            }
            catch {
                GuardarFichero("ERROR", autor, descripcion);
            }
        }

        public static void InfoFichero(string autor, string descripcion)
        {
            GuardarFichero("INFO", autor, descripcion);
        }

        public static void ErrorFichero(string autor, string descripcion)
        {
            GuardarFichero("ERROR", autor, descripcion);
        }

        private static void GuardarFichero(string tipo, string autor, string descripcion)
        {
            try
            {
                string carpeta = Properties.Settings.Default.LogFolder;
                Directory.CreateDirectory(carpeta);

                string fichero = Path.Combine(carpeta, $"{DateTime.Now:yyyy-MM-dd}.txt");

                File.AppendAllText(fichero,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {tipo} | {autor} | {descripcion}{Environment.NewLine}");
            }
            catch { }
        }
    }
}
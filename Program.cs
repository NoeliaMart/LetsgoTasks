using System;
using NLog;
using NortConsultingTasks.Tasks;
using NortConsultingTasks.Utils;

namespace NortConsultingTasks
{
    class Program
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        static void Main(string[] args)
        {
            logger.Info("Application started.");

            try
            {
                if (args.Length == 0)
                {
                    return;
                }

                string command = args[0].ToLower();

                switch (command.ToUpper())
                {
                    case "ENCRIPTA":
                        Console.WriteLine(NortConsultingTasks.Utils.EncryptionHelper.EncryptPass(args[1]));
                        break;
                    case "TESTCONEXIONBBDD":
                        try
                        {
                            cBBDD.SetDbConnection();
                            Console.WriteLine("Conexión correcta");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error realizando conexion: " + ex.Message);
                        }

                        break;
                    case "TESTCONEXIONSAP":
                        try
                        {
                            string err = string.Empty;
                            cDIAPI.ConnectionCompany();

                            if (string.IsNullOrEmpty(err))
                            {
                                Console.WriteLine("Conexión correcta");
                            }
                            else
                            {
                                Console.WriteLine(err);
                            }

                           
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error realizando conexion: " + ex.Message);
                        }

                        break;
                    case "MARGENHW":
                        ActualizaMargenes.Ejecuta(logger);
                        break;
                    case "CONTACTOS":
                        GeneraPersonasContacto.Ejecuta(logger);
                        break;
                    case "ENVIAVENCI":
                        EnvioVencimientos.Ejecuta(logger);
                        break;
                    case "CREACAMPO":
                        CreaCampos.Ejecuta(logger);
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error :" + ex.Message);
                logger.Error(ex, "An unexpected error occurred.");
            }
            finally
            {
                logger.Info("Application ended.");
            }
        }
    }
}

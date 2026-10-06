using System;
using LetsGoTasks.Tasks;
using LetsGoTasks.Utils;

namespace LetsGoTasks
{
    class Program
    {
        static void Main(string[] args)
        {
            Log.Info("Main", "Application started.");

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
                        Console.WriteLine(LetsGoTasks.Utils.EncryptionHelper.EncryptPass(args[1]));
                        Console.WriteLine("Pulsa una tecla para continuar...");
                        Console.ReadKey();
                        break;
                    case "TESTCONEXIONBBDD":
                        try
                        {

                            cBBDD.SetDbConnection(args[1]);
                            Console.WriteLine("Conexión correcta");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error realizando conexion: " + ex.Message);
                        }

                        Console.WriteLine("Pulsa una tecla para continuar...");
                        Console.ReadKey();
                        break;
                    case "TESTCONEXIONSAP":
                        try
                        {
                            string err = string.Empty;
                            cDIAPI.ConnectionCompany(args[1]);

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

                        Console.WriteLine("Pulsa una tecla para continuar...");
                        Console.ReadKey();
                        break;
                    case "CSVPROVEEDORES":
                        ProcesaCargaFacturas.Ejecuta();
                        break;
                    case "CREACAMPOS":
                        CreaCampos.Ejecuta(args[1]);
                        Console.WriteLine("Pulsa una tecla para continuar...");
                        Console.ReadKey();
                        break;
                    case "DESCARGASFTP":
                        DescargaFicherosSFTP.Ejecuta();
                        break;
                    case "ACTUALIZACAMBIO":
                        ActualizaCambioMoneda.Ejecuta();
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error :" + ex.Message);
            }
        }
    }
}

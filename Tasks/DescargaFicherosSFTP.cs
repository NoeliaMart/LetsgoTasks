using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using LetsGoTasks.Utils;
using Renci.SshNet;

namespace LetsGoTasks.Tasks
{
    public class DescargaFicherosSFTP
    {
        public static void Ejecuta()
        {
            string host = Properties.Settings.Default.SFTPHost;
            int port = Properties.Settings.Default.SFTPPort;
            string user = Properties.Settings.Default.SFTPUser;
            string password = EncryptionHelper.DecryptPass(Properties.Settings.Default.SFTPPassword);
            string carpetaRemota = Properties.Settings.Default.SFTPFolder;
            string carpetaLocal = Properties.Settings.Default.CSVProvFolder;
            string destinoRemoto = carpetaRemota.TrimEnd('/') + "/" + "SAP_PROCESADO";

            var connectionInfo = new PasswordConnectionInfo(host,port,user,password);

            connectionInfo.Timeout = TimeSpan.FromSeconds(30);

            try
            {
                using (var sftp = new SftpClient(connectionInfo))
                {

                    Log.InfoFichero("DescargaFicherosSFTP", $"Conectando {host}:{port} con usuario {user}");

                    sftp.Connect();

                    foreach (var fichero in sftp.ListDirectory(carpetaRemota)
                                               .Where(x => !x.IsDirectory && x.Name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)))
                    {
                        string destino = Path.Combine(carpetaLocal, fichero.Name);

                        using (var stream = File.Create(destino))
                            sftp.DownloadFile(fichero.FullName, stream);

                        Log.Info("DescargaFicherosSFTP", $"Fichero descargado por SFTP: {fichero.Name}");

                        string ficheroDestinoRemoto = destinoRemoto.TrimEnd('/')+ "/"+ fichero.Name;

                        sftp.RenameFile(fichero.FullName, ficheroDestinoRemoto);
                    }

                    sftp.Disconnect();
                }
            }
            catch (Exception ex)
            {
                Log.ErrorFichero("DescargaFicherosSFTP", $"Error descargando ficheros SFTP: {ex.Message}");
            }
        }
    }
}
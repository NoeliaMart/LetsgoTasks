using Renci.SshNet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace LetsGoTasks.Utils
{
    public class cUtils
    {
        public static bool EsEmailValido(string email)
        {
			string emailPattern = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+.[a-zA-Z]{2,}$";

			if (string.IsNullOrEmpty(email))
				return false;

			Regex regex = new Regex(emailPattern);
			return regex.IsMatch(email);
		}

        public static void MoverFicheroSFTP(string fichero, string destino)
        {
            string host = Properties.Settings.Default.SFTPHost;
            int port = Properties.Settings.Default.SFTPPort;
            string user = Properties.Settings.Default.SFTPUser;
            string password = EncryptionHelper.DecryptPass(Properties.Settings.Default.SFTPPassword);
            string carpetaRemota = Properties.Settings.Default.SFTPFolder;
            string carpetaLocal = Properties.Settings.Default.CSVProvFolder;
            string destinoRemoto = carpetaRemota.TrimEnd('/') + "/" + destino + "/" + Path.GetFileName(fichero);

            var connectionInfo = new PasswordConnectionInfo(host, port, user, password);

            connectionInfo.Timeout = TimeSpan.FromSeconds(30);

            try
            {
                using (var sftp = new SftpClient(connectionInfo))
                {
                    sftp.Connect();

                    using (var stream = File.OpenRead(fichero))
                        sftp.UploadFile(stream, destinoRemoto);

                    sftp.Disconnect();

                    File.Delete(fichero);
                }
            }
            catch (Exception ex)
            {
                Log.ErrorFichero("MoverFicheroSFTP", $"Error moviendo fichero {fichero} a SFTP: {ex.Message}");
                throw;
            }
        }
    }
}

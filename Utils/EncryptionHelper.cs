using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace NortConsultingTasks.Utils
{
    public static class EncryptionHelper
    {
        // Función de encriptación
        public static string EncryptPass(string textToEncrypt)
        {
            try
            {
                string ToReturn = "";
                string publickey = "87654321";  // Clave pública (8 bytes)
                string secretkey = "nortCsb1";  // Clave secreta (8 bytes)

                byte[] privatekeyByte = Encoding.UTF8.GetBytes(secretkey);
                byte[] publickeybyte = Encoding.UTF8.GetBytes(publickey);

                if (privatekeyByte.Length != 8 || publickeybyte.Length != 8)
                {
                    throw new Exception("Las claves deben tener una longitud de 8 bytes para DES.");
                }

                byte[] inputByteArray = Encoding.UTF8.GetBytes(textToEncrypt);

                using (DESCryptoServiceProvider des = new DESCryptoServiceProvider())
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        using (CryptoStream cs = new CryptoStream(ms, des.CreateEncryptor(publickeybyte, privatekeyByte), CryptoStreamMode.Write))
                        {
                            cs.Write(inputByteArray, 0, inputByteArray.Length);
                            cs.FlushFinalBlock();
                        }

                        ToReturn = Convert.ToBase64String(ms.ToArray());
                    }
                }

                return ToReturn;
            }
            catch (Exception ae)
            {
                throw new Exception("Error al intentar encriptar el texto: " + ae.Message, ae);
            }
        }

        // Función de desencriptación
        public static string DecryptPass(string textToDecrypt)
        {
            try
            {
                string ToReturn = "";
                string publickey = "87654321";  // Clave pública (8 bytes)
                string secretkey = "nortCsb1";  // Clave secreta (8 bytes)

                byte[] privatekeyByte = Encoding.UTF8.GetBytes(secretkey);
                byte[] publickeybyte = Encoding.UTF8.GetBytes(publickey);

                if (privatekeyByte.Length != 8 || publickeybyte.Length != 8)
                {
                    throw new Exception("Las claves deben tener una longitud de 8 bytes para DES.");
                }

                // Limpiar el texto Base64, asegurándonos de que tiene el padding correcto
                string cleanedText = CleanBase64String(textToDecrypt);

                byte[] inputByteArray = Convert.FromBase64String(cleanedText);

                using (DESCryptoServiceProvider des = new DESCryptoServiceProvider())
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        using (CryptoStream cs = new CryptoStream(ms, des.CreateDecryptor(publickeybyte, privatekeyByte), CryptoStreamMode.Write))
                        {
                            cs.Write(inputByteArray, 0, inputByteArray.Length);
                            cs.FlushFinalBlock();
                        }

                        ToReturn = Encoding.UTF8.GetString(ms.ToArray());
                    }
                }

                return ToReturn;
            }
            catch (Exception ae)
            {
                throw new Exception("Error al intentar desencriptar el texto: " + ae.Message, ae);
            }
        }

        // Función para asegurarse de que la cadena Base64 tiene el padding adecuado
        private static string CleanBase64String(string base64)
        {
            string cleanedText = base64.Replace(" ", "+");

            // Ajustar padding si es necesario
            int padding = cleanedText.Length % 4;
            if (padding > 0)
            {
                cleanedText = cleanedText.PadRight(cleanedText.Length + (4 - padding), '=');
            }

            return cleanedText;
        }
    }
}
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace RBS_Reports
{
    class Crypting
    {
        public static string key = "THIS_IS_MPT";
        public static byte[] IV = new ASCIIEncoding().GetBytes("THIS_IS_GOOD_AES");

        public static string getHash(string text)
        {
            byte[] data = new UTF8Encoding().GetBytes(text);
            SHA256 shaM = new SHA256Managed();
            return BitConverter.ToString(shaM.ComputeHash(data)).Replace("-", "").ToLower();
        }

        public static byte[] getHashBytes(string text)
        {
            byte[] data = new UTF8Encoding().GetBytes(text);
            SHA256 shaM = new SHA256Managed();
            return shaM.ComputeHash(data);
        }

        public static string encryptAES(string text)
        {
            byte[] bytes = Encoding.Unicode.GetBytes(text);
            //Encrypt
            SymmetricAlgorithm crypt = Aes.Create();
            HashAlgorithm hash = MD5.Create();
            crypt.BlockSize = 128;
            crypt.Key = hash.ComputeHash(Encoding.Unicode.GetBytes(key));
            crypt.IV = IV;

            using (MemoryStream memoryStream = new MemoryStream())
            {
                using (CryptoStream cryptoStream =
                   new CryptoStream(memoryStream, crypt.CreateEncryptor(), CryptoStreamMode.Write))
                {
                    cryptoStream.Write(bytes, 0, bytes.Length);
                }

                return Convert.ToBase64String(memoryStream.ToArray());
            }
        }

        public static string decryptAES(string text)
        {
            byte[] bytes = Convert.FromBase64String(text);
            SymmetricAlgorithm crypt = Aes.Create();
            HashAlgorithm hash = MD5.Create();
            crypt.Key = hash.ComputeHash(Encoding.Unicode.GetBytes(key));
            crypt.IV = IV;

            using (MemoryStream memoryStream = new MemoryStream(bytes))
            {
                using (CryptoStream cryptoStream =
                   new CryptoStream(memoryStream, crypt.CreateDecryptor(), CryptoStreamMode.Read))
                {
                    byte[] decryptedBytes = new byte[bytes.Length];
                    cryptoStream.Read(decryptedBytes, 0, decryptedBytes.Length);
                    return Encoding.Unicode.GetString(decryptedBytes);
                }
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Send_Email
{
    public static class EncryptExtend
    {
        //CMMS: 2899a4bf
        //POPsystem: cv*8788
        private static string _BasekeyString = "vjitgmes"; //= "popsystem"; // "2899a4bf";
        static string strPasswordKey = "vjitgmes";//  "popsystem"; //cv*8788
        public static string ToEncryptString(this string PlainText)
        {
            if (string.IsNullOrEmpty(PlainText)) return PlainText;

            byte[] buffer = Encrypt(PlainText);

            return Convert.ToBase64String(buffer, 0, buffer.Length);
        }

        private static byte[] Encrypt(string PlainText)
        {
            DESCryptoServiceProvider key = new DESCryptoServiceProvider();

            key.Key = System.Text.Encoding.Default.GetBytes(_BasekeyString);
            key.IV = System.Text.Encoding.Default.GetBytes(_BasekeyString);

            // Create a memory stream.
            MemoryStream ms = new MemoryStream();

            // Create a CryptoStream using the memory stream and the 
            // CSP DES key.
            CryptoStream encStream = new CryptoStream(ms, key.CreateEncryptor(), CryptoStreamMode.Write);

            // Create a StreamWriter to write a string
            // to the stream.
            StreamWriter sw = new StreamWriter(encStream);

            // Write the plaintext to the stream.
            sw.WriteLine(PlainText);

            // Close the StreamWriter and CryptoStream.
            sw.Close();

            encStream.Close();

            // Get an array of bytes that represents
            // the memory stream.
            byte[] buffer = ms.ToArray();

            // Close the memory stream.
            ms.Close();

            // Return the encrypted byte array.
            return buffer;
        }

        private static string Decrypt(byte[] CypherText)
        {
            try
            {
                DESCryptoServiceProvider key = new DESCryptoServiceProvider();

                key.Key = System.Text.Encoding.Default.GetBytes(_BasekeyString);
                key.IV = System.Text.Encoding.Default.GetBytes(_BasekeyString);

                // Create a memory stream to the passed buffer.
                MemoryStream ms = new MemoryStream(CypherText);

                // Create a CryptoStream using the memory stream and the 
                // CSP DES key. 
                CryptoStream encStream = new CryptoStream(ms, key.CreateDecryptor(), CryptoStreamMode.Read);

                // Create a StreamReader for reading the stream.
                StreamReader sr = new StreamReader(encStream);

                // Read the stream as a string.
                string val = sr.ReadLine();

                // Close the streams.
                sr.Close();

                encStream.Close();

                ms.Close();

                return val;
            }
            catch { return ""; }
        }

        public static string ToDecryptString(this string CypherText)
        {
            if (string.IsNullOrEmpty(CypherText)) return CypherText;

            byte[] buffer = Convert.FromBase64String(CypherText);

            return Decrypt(buffer).ToString();
        }


        public static string EncryptString(string text)
        {

            byte[] inputText = System.Text.Encoding.Unicode.GetBytes(text);
            byte[] passwordSalt = Encoding.ASCII.GetBytes(strPasswordKey.Length.ToString());

            PasswordDeriveBytes secretKey = new PasswordDeriveBytes(strPasswordKey, passwordSalt);
            Rijndael rijAlg = Rijndael.Create();


            rijAlg.Key = secretKey.GetBytes(32);
            rijAlg.IV = secretKey.GetBytes(16);

            ICryptoTransform encryptor = rijAlg.CreateEncryptor(rijAlg.Key, rijAlg.IV);
            MemoryStream msEncrypt = new MemoryStream();
            CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write);


            csEncrypt.Write(inputText, 0, inputText.Length);
            csEncrypt.FlushFinalBlock();


            byte[] encryptBytes = msEncrypt.ToArray();
            msEncrypt.Close();
            csEncrypt.Close();

            // Base64
            string encryptedData = Convert.ToBase64String(encryptBytes);
            return encryptedData;
        }


        public static string DescryptString(string text)
        {

            byte[] encryptedData = Convert.FromBase64String(text);

            byte[] passwordSalt = Encoding.ASCII.GetBytes(strPasswordKey.Length.ToString());


            PasswordDeriveBytes secretKey = new PasswordDeriveBytes(strPasswordKey, passwordSalt);

            Rijndael rijAlg = Rijndael.Create();


            rijAlg.Key = secretKey.GetBytes(32);
            rijAlg.IV = secretKey.GetBytes(16);


            ICryptoTransform decryptor = rijAlg.CreateDecryptor(rijAlg.Key, rijAlg.IV);
            MemoryStream msDecrypt = new MemoryStream(encryptedData);
            CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read);

            int decryptedCount = csDecrypt.Read(encryptedData, 0, encryptedData.Length);

            msDecrypt.Close();
            csDecrypt.Close();



            // Base64
            string decryptedData = Encoding.Unicode.GetString(encryptedData, 0, decryptedCount);
            return decryptedData;

        }
    }
}

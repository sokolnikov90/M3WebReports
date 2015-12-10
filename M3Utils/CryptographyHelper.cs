using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace M3Utils
{
    public static class CryptographyHelper
    {
        public static string CryptPassword8(string passw)
        {
            if (passw.Length < 8)
            {
                for (int h = passw.Length; h < 9; h++)
                    passw += "\0";
            }
            else if (passw.Length > 8)
            {
                passw = passw.Substring(0, 8);
            }

            return CryptPassword(passw);
        }

        private static string CryptPassword(string passw)
        {
            byte[] key = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16 };

            TripleDESCryptoServiceProvider des = new TripleDESCryptoServiceProvider { Key = key, Mode = CipherMode.ECB };

            byte[] buff = Encoding.Default.GetBytes(passw);

            byte[] tmp = des.CreateEncryptor().TransformFinalBlock(buff, 0, buff.Length);

            return tmp.Aggregate("", (current, b) => current + b.ToString("X2"));
        }

        public static string DecryptPassword(string passw)
        {
            if (passw == null)
                return "";
            if (passw.Length == 0)
                return "";
            byte[] key = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16 };

            TripleDESCryptoServiceProvider des = new TripleDESCryptoServiceProvider();
            des.Key = key;
            des.Mode = CipherMode.ECB;

            byte[] buff = StringHelper.HexStr2ByteArr(passw);

            byte[] tmp = des.CreateDecryptor().TransformFinalBlock(buff, 0, buff.Length);

            string encrypted = Encoding.Default.GetString(tmp);

            return encrypted;
        }
    }
}

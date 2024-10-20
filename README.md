using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

public class ClsPDScoreCardEncryption
{
    public static string Encrypt(string key, string data)
    {
        string encData = null;
        byte[][] keys = GetHashKeys(key);
        encData = EncryptString_Aes(data, keys[0], keys[1]);

        return encData;
    }

    public static string Decrypt(string key, string data)
    {
        string decData = null;
        byte[][] keys = GetHashKeys(key);

        decData = DecryptString_Aes(data, keys[0], keys[1]);

        return decData;
    }

    public static byte[][] GetHashKeys(string key)
    {
        byte[][] result = new byte[2][];
        Encoding enc = Encoding.UTF8;
        using (SHA256 sha2 = new SHA256CryptoServiceProvider())
        {
            byte[] rawKey = enc.GetBytes(key);
            byte[] rawIV = enc.GetBytes(key);
            byte[] hashKey = sha2.ComputeHash(rawKey);
            byte[] hashIV = sha2.ComputeHash(rawIV);
            Array.Resize(ref hashIV, 16);
            result[0] = hashKey;
            result[1] = hashIV;
        }
        return result;
    }

    public static string EncryptString_Aes(string plainText, byte[] Key, byte[] IV)
    {
        if (string.IsNullOrEmpty(plainText))
            throw new ArgumentNullException(nameof(plainText));
        if (Key == null || Key.Length <= 0)
            throw new ArgumentNullException(nameof(Key));
        if (IV == null || IV.Length <= 0)
            throw new ArgumentNullException(nameof(IV));

        byte[] encrypted;

        using (AesManaged aesAlg = new AesManaged())
        {
            aesAlg.Padding = PaddingMode.None;
            aesAlg.Key = Key;
            aesAlg.IV = IV;

            ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);

            using (MemoryStream msEncrypt = new MemoryStream())
            {
                using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                {
                    using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                    {
                        swEncrypt.Write(plainText);
                    }
                    encrypted = msEncrypt.ToArray();
                }
            }
        }

        return Convert.ToBase64String(encrypted);
    }

    public static string DecryptString_Aes(string cipherTextString, byte[] Key, byte[] IV)
    {
        byte[] cipherText = Convert.FromBase64String(cipherTextString);

        if (cipherText == null || cipherText.Length <= 0)
            throw new ArgumentNullException(nameof(cipherText));
        if (Key == null || Key.Length <= 0)
            throw new ArgumentNullException(nameof(Key));
        if (IV == null || IV.Length <= 0)
            throw new ArgumentNullException(nameof(IV));

        string plaintext = null;

        using (AesManaged aesAlg = new AesManaged())
        {
            aesAlg.Padding = PaddingMode.Zeros;
            aesAlg.Key = Key;
            aesAlg.IV = IV;

            ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);

            using (MemoryStream msDecrypt = new MemoryStream(cipherText))
            {
                using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                {
                    using (StreamReader srDecrypt = new StreamReader(csDecrypt))
                    {
                        plaintext = srDecrypt.ReadToEnd();
                    }
                }
            }
        }

        return Regex.Replace(plaintext, @"[^a-zA-Z0-9\\:_\- ]", "");
    }
}

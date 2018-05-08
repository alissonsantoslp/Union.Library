using System;
using System.Collections.Generic;
using System.Data;
using System.Security.Cryptography;
using System.Text;

namespace Union.Library
{
    public class Lib
    {
        const string senha = "!@#$%¨&*()_+'1234567890-=qwertyuiop´[asdfghjklç~]zxcvbnm,.;/QWERTYUIOP`{ASDFGHJKLÇ^}ZXCVBNM<>:?ªº°§";

        #region Serializa_Jason
        public string Serializa_Jason(DataSet dsInput)
        {
            try
            {
                DataSet ds = new DataSet();
                ds = dsInput;
                StringBuilder JsonString = new StringBuilder();
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    JsonString.Append("[");
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        JsonString.Append("{");
                        for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                        {
                            if (j < ds.Tables[0].Columns.Count - 1)
                            {
                                JsonString.AppendFormat("\"{0}\":\"{1}\",", ds.Tables[0].Columns[j].ColumnName, ds.Tables[0].Rows[i][j]);
                            }
                            else if (j == ds.Tables[0].Columns.Count - 1)
                            {
                                JsonString.AppendFormat("\"{0}\":\"{1}\"", ds.Tables[0].Columns[j].ColumnName, ds.Tables[0].Rows[i][j]);
                            }
                        }
                        if (i == ds.Tables[0].Rows.Count - 1)
                        {
                            JsonString.Append("}");
                        }
                        else
                        {
                            JsonString.Append("},");
                        }
                    }
                    JsonString.Append("]");
                    return JsonString.ToString();
                }
                else
                    return string.Empty;
            }
            catch (Exception)
            {
                throw;
            }
        }
        #endregion

        #region Metodo que valida o CPF
        public bool ValidaCPF(string vrCPF)
        {
            string valor = vrCPF.Replace(".", "");
            valor = valor.Replace("-", "");
            
            if (valor.Length != 11)
                return false;
 
            bool igual = true;
            for (int i = 1; i < 11 && igual; i++)
                if (valor[i] != valor[0])
                    igual = false;
 
            if (igual || valor == "12345678909")
                return false;
 
            int[] numeros = new int[11];
            for (int i = 0; i < 11; i++)
                numeros[i] = int.Parse(
                valor[i].ToString());
 
            int soma = 0;
            for (int i = 0; i < 9; i++)
                soma += (10 - i) * numeros[i];
           
            int resultado = soma % 11;
            if (resultado == 1 || resultado == 0)
            {
                if (numeros[9] != 0)
                    return false;
            }
            else if (numeros[9] != 11 - resultado)
                return false;
 
            soma = 0;
            for (int i = 0; i < 10; i++)
                soma += (11 - i) * numeros[i];
 
            resultado = soma % 11;
 
            if (resultado == 1 || resultado == 0)
            {
                if (numeros[10] != 0)
                    return false;
 
            }
            else
                if (numeros[10] != 11 - resultado)
                    return false;
            return true;
 
        }
        #endregion

        #region Metodo que valida o CNPJ
        public bool ValidaCNPJ(string vrCNPJ)
        {
 
            string CNPJ = vrCNPJ.Replace(".", "");
            CNPJ = CNPJ.Replace("/", "");
            CNPJ = CNPJ.Replace("-", "");
 
            int[] digitos, soma, resultado;
            int nrDig;
            string ftmt;
            bool[] CNPJOk;
 
            ftmt = "6543298765432";
            digitos = new int[14];
            soma = new int[2];
            soma[0] = 0;
            soma[1] = 0;
            resultado = new int[2];
            resultado[0] = 0;
            resultado[1] = 0;
            CNPJOk = new bool[2];
            CNPJOk[0] = false;
            CNPJOk[1] = false;
 
            try
            {
                for (nrDig = 0; nrDig < 14; nrDig++)
                {
                    digitos[nrDig] = int.Parse(
                     CNPJ.Substring(nrDig, 1));
                    if (nrDig <= 11)
                        soma[0] += (digitos[nrDig] *
                        int.Parse(ftmt.Substring(
                          nrDig + 1, 1)));
                    if (nrDig <= 12)
                        soma[1] += (digitos[nrDig] *
                        int.Parse(ftmt.Substring(
                          nrDig, 1)));
                  }
 
                for (nrDig = 0; nrDig < 2; nrDig++)
                {
                    resultado[nrDig] = (soma[nrDig] % 11);
                    if ((resultado[nrDig] == 0) || (resultado[nrDig] == 1))
                        CNPJOk[nrDig] = (
                        digitos[12 + nrDig] == 0);
 
                    else
                        CNPJOk[nrDig] = (
                        digitos[12 + nrDig] == (
                        11 - resultado[nrDig]));
 
                }
 
                return (CNPJOk[0] && CNPJOk[1]);
 
            }
            catch
            {
                return false;
            }
 
        }
        #endregion

        #region Metodo que valida o Cep
        public bool ValidaCep(string cep)
        {
            if (cep.Length == 8)
            {
                cep = cep.Substring(0, 5) + "-" + cep.Substring(5, 3);
                //txt.Text = cep;
            }
            return System.Text.RegularExpressions.Regex.IsMatch(cep, ("[0-9]{5}-[0-9]{3}"));
        }
        #endregion

        #region Metodo que valida o Email
        public bool ValidaEmail(string email)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(email, ("(?<user>[^@]+)@(?<host>.+)"));
        }
        #endregion

        public static string Encrypt(string Message)
        {
            byte[] Results;
            System.Text.UTF8Encoding UTF8 = new System.Text.UTF8Encoding();
            MD5CryptoServiceProvider HashProvider = new MD5CryptoServiceProvider();
            byte[] TDESKey = HashProvider.ComputeHash(UTF8.GetBytes(senha));
            TripleDESCryptoServiceProvider TDESAlgorithm = new TripleDESCryptoServiceProvider() { Key = TDESKey, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 };
            byte[] DataToEncrypt = UTF8.GetBytes(Message);

            try
            {
                ICryptoTransform Encryptor = TDESAlgorithm.CreateEncryptor();
                Results = Encryptor.TransformFinalBlock(DataToEncrypt, 0, DataToEncrypt.Length);
            }
            finally
            {
                TDESAlgorithm.Clear();
                HashProvider.Clear();
            }
            return Convert.ToBase64String(Results);
        }

        public static string Decrypt(string Message)
        {
            byte[] Results;
            System.Text.UTF8Encoding UTF8 = new System.Text.UTF8Encoding();
            MD5CryptoServiceProvider HashProvider = new MD5CryptoServiceProvider();
            byte[] TDESKey = HashProvider.ComputeHash(UTF8.GetBytes(senha));
            TripleDESCryptoServiceProvider TDESAlgorithm = new TripleDESCryptoServiceProvider() { Key = TDESKey, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 };
            byte[] DataToDecrypt = Convert.FromBase64String(Message);

            try
            {
                ICryptoTransform Decryptor = TDESAlgorithm.CreateDecryptor();
                Results = Decryptor.TransformFinalBlock(DataToDecrypt, 0, DataToDecrypt.Length);
            }
            finally
            {
                TDESAlgorithm.Clear();
                HashProvider.Clear();
            }

            return UTF8.GetString(Results);
        }



    }
}

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace Union.Library
{
    public class connect_BaseSQLServer : IDisposable
    {
        private SqlConnection conn = null;
        private SqlCommand cmd = null;
        private SqlDataReader dr = null;
        private SqlDataAdapter da = null;
        private DataSet ds = null;
        private DataTable dt = null;
        private SqlTransaction transacao = null;
        private string pStringdeConexao = string.Empty;

        public connect_BaseSQLServer(string StringdeConexao)
        {
            pStringdeConexao = StringdeConexao;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (conn != null)
                {
                    conn.Dispose();
                    conn = null;
                }
                if (cmd != null)
                {
                    cmd.Dispose();
                    cmd = null;
                }
                if (dr != null)
                {
                    dr.Dispose();
                    dr = null;
                }
                if (da != null)
                {
                    da.Dispose();
                    da = null;
                }
                if (ds != null)
                {
                    ds.Dispose();
                    ds = null;
                }
                if (dt != null)
                {
                    dt.Dispose();
                    dt = null;
                }
                if (transacao != null)
                {
                    transacao.Dispose();
                    transacao = null;
                }
            }
        }

        ~connect_BaseSQLServer()
        {
            Dispose(false);
        }

        #region Abre Conexão com banco de dados
        private SqlConnection AbreBanco()
        {
            conn = new SqlConnection(pStringdeConexao);
            try
            {
                conn.Open();
            }
            catch
            {
                throw;
            }
            return conn;
        }
        #endregion

        #region Fecha Conexão com banco de dados
        private void FechaBanco(SqlConnection conn)
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
        #endregion

        public void ExecutaComando(string strQuery)
        {
            try
            {
                conn = AbreBanco();
                cmd = new SqlCommand();
                cmd.CommandText = strQuery;
                cmd.Connection = conn;
                cmd.ExecuteNonQuery();
            }
            catch
            {
                throw;
            }
            finally
            {
                FechaBanco(conn);
                conn.Dispose();
                cmd.Dispose();
            }
        }

        public int ExecutaComandoRetorno(string strQuery)
        {
            conn = new SqlConnection();
            int flag = 0;
            try
            {
                conn = AbreBanco();
                cmd = new SqlCommand();
                cmd.CommandText = strQuery;
                cmd.Connection = conn;
                cmd.ExecuteNonQuery();
                dr = cmd.ExecuteReader();
                flag = 1;
                dr.Read();
                return flag;
            }
            catch
            {
                flag = 0;
                return flag;
            }
            finally
            {
                FechaBanco(conn);
                conn.Dispose();
                cmd.Dispose();
                dr.Dispose();
            }
        }

        public void ExecutaSqlComando(SqlCommand Commando)
        {
            conn = new SqlConnection();
            try
            {
                conn = AbreBanco();
                cmd.Connection = conn;
                Commando.ExecuteNonQuery();
            }
            catch
            {
                throw;
            }
            finally
            {
                FechaBanco(conn);
                conn.Dispose();
                cmd.Dispose();
            }
        }

        public DataSet RetornaDataSet(string strQuery)
        {
            conn = new SqlConnection();
            try
            {
                conn = AbreBanco();
                cmd = new SqlCommand();
                cmd.CommandText = strQuery;
                cmd.Connection = conn;
                da = new SqlDataAdapter();
                ds = new DataSet();
                da.SelectCommand = cmd;
                da.Fill(ds);
                return ds;
            }
            catch
            {
                throw;
            }
            finally
            {
                FechaBanco(conn);
            }
        }

        public DataTable RetornaDataTable(string strQuery)
        {
            conn = new SqlConnection();
            try
            {
                conn = AbreBanco();
                cmd = new SqlCommand();
                cmd.CommandText = strQuery;
                cmd.Connection = conn;
                da = new SqlDataAdapter();
                dt = new DataTable();
                da.SelectCommand = cmd;
                da.Fill(dt);
                return dt;
            }
            catch
            {
                throw;
            }
            finally
            {
                FechaBanco(conn);
                conn.Dispose();
                cmd.Dispose();
                da.Dispose();
                dt.Dispose();
            }
        }

        public DataTable RetornaDataTableCommando(SqlCommand commando)
        {
            conn = new SqlConnection();
            try
            {
                conn = AbreBanco();
                cmd.Connection = conn; ;
                dt = new DataTable();
                da = new SqlDataAdapter();
                da.SelectCommand = commando;
                da.Fill(dt);
                return dt;
            }
            catch
            {
                throw;
            }
            finally
            {
                FechaBanco(conn);
                conn.Dispose();
                cmd.Dispose();
                dt.Dispose();
                da.Dispose();
            }
        }

        public SqlDataReader RetornaDataReader(string strQuery)
        {
            conn = new SqlConnection();
            try
            {
                conn = AbreBanco();
                cmd = new SqlCommand();
                cmd.CommandText = strQuery;
                cmd.Connection = conn;
                return cmd.ExecuteReader();
            }
            catch
            {
                throw;
            }
            finally
            {
                FechaBanco(conn);
                conn.Dispose();
                //cmd.Dispose();
            }
        }

        public void ExecutaComando(SqlCommand Comm)
        {
            conn = new SqlConnection();
            conn = AbreBanco();
            Comm.Connection = conn;
            try
            {
                Comm.Connection.Open();
                Comm.ExecuteNonQuery();
            }
            catch
            {
                throw;
            }
            finally
            {
                FechaBanco(conn);
                conn.Dispose();
            }
        }
    }
}

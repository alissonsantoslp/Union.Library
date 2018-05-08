using System;
using SAPbobsCOM;

namespace Union.Library
{
    /// <summary>
    /// Api SAP
    /// </summary>
    public class DiCompany
    {
        /// <summary>
        /// Metodo utilizado para Validação de usuário logado no SAP.
        /// </summary>
        /// <param name="sUsuario">Informar Usuário do SAP</param>
        /// <param name="sSenha">Informar Senha do usuário que está informado no parametro sUsuario</param>
        /// <example>
        ///     <code> 
        ///         SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company()    
        /// 
        ///         Union.Library obj = new Union.Library()
        ///         oCompany = obj.ConectarUsuarioEspecifico("Usuario","Senha")
        ///     </code>
        /// </example>
        /// <returns>Company</returns>
        /// <remarks>Metodo utilizado para conectar com um usuario especifico no SAP e também para validar usuário que está logando.</remarks>
        /// <exception cref="SystemException">Falha! Erro: Descrição da Exception, Descricao B1: Descrição do erro do SAP, Codigo Erro B1: Código do Erro do SAP</exception>
        public Company ConectarUsuarioEspecifico(string sUsuario, string sSenha)
        {
            XmlHelper xml = new XmlHelper();
            Company oCompany = new Company();
            int iRetCode = -1;

            try
            {
                if (!oCompany.Connected)
                {
                    oCompany.Server = xml.ReadSettingAppConfig("Servidor");
                    oCompany.DbUserName = xml.ReadSettingAppConfig("UsuarioSQL");
                    oCompany.DbPassword = xml.ReadSettingAppConfig("SenhaSQL");
                    oCompany.CompanyDB = xml.ReadSettingAppConfig("BancoDadosSAP");
                    oCompany.UserName = sUsuario.ToString().Trim();
                    oCompany.Password = sSenha.ToString().Trim();
                    oCompany.LicenseServer = xml.ReadSettingAppConfig("ServidorLicenca");
                    oCompany.language = BoSuppLangs.ln_Portuguese_Br;

                    int iDbServerType = Convert.ToInt32(xml.ReadSettingAppConfig("VersaoBancoDados"));

                    switch (iDbServerType)
                    {
                        case 1:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_MSSQL;
                            }
                            break;
                        case 2:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_DB_2;
                            }
                            break;
                        case 3:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_SYBASE;
                            }
                            break;
                        case 4:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2005;
                            }
                            break;
                        case 5:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_MAXDB;
                            }
                            break;
                        case 6:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2008;
                            }
                            break;
                        case 7:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2012;
                            }
                            break;
                        case 8:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2014;
                            }
                            break;
                        case 9:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_HANADB;
                            }
                            break;
                    }

                    iRetCode = oCompany.Connect();

                    if (iRetCode != 0)
                    {
                        string ErrorDescription = oCompany.GetLastErrorDescription();
                        return oCompany;
                    }
                    else
                    {
                        return oCompany;
                    }
                }
                else
                {
                    return oCompany;
                }
            }
            catch (Exception ex)
            {
                return oCompany;
                throw new Exception(string.Format("Falha! \nErro: {0}\nDescricao B1: {1}\nCodigo Erro B1: {2}", ex.Message, oCompany.GetLastErrorDescription(), iRetCode.ToString()));
            }
        }

        #region Conectar
        public Company Conectar()
        {
            XmlHelper xml = new XmlHelper();
            Company oCompany = new Company();
            int iRetCode = -1;
            
            try
            {
                if (!oCompany.Connected)
                {
                    oCompany.Server = xml.ReadSettingAppConfig("Servidor");
                    oCompany.DbUserName = xml.ReadSettingAppConfig("UsuarioSQL");
                    oCompany.DbPassword = xml.ReadSettingAppConfig("SenhaSQL");
                    oCompany.CompanyDB = xml.ReadSettingAppConfig("BancoDadosSAP");
                    oCompany.UserName = xml.ReadSettingAppConfig("UsuarioSAP");
                    oCompany.Password = xml.ReadSettingAppConfig("SenhaSAP");
                    oCompany.LicenseServer = xml.ReadSettingAppConfig("ServidorLicenca");
                    oCompany.language = BoSuppLangs.ln_Portuguese_Br;

                    int iDbServerType = Convert.ToInt32(xml.ReadSettingAppConfig("VersaoBancoDados"));

                    switch(iDbServerType)
                    {
                        case 1:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_MSSQL;
                            }
                            break;
                        case 2:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_DB_2;
                            }
                            break;
                        case 3:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_SYBASE;
                            }
                            break;
                        case 4:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2005;
                            }
                            break;
                        case 5:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_MAXDB;
                            }
                            break;
                        case 6:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2008;
                            }
                            break;
                        case 7:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2012;
                            }
                            break;
                        case 8:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2014;
                            }
                            break;
                        case 9:
                            {
                                oCompany.DbServerType = BoDataServerTypes.dst_HANADB;
                            }
                            break;
                    }

                    iRetCode = oCompany.Connect();

                    if (iRetCode != 0)
                    {
                        string ErrorDescription = oCompany.GetLastErrorDescription();
                        throw new Exception("Falha! \nDescricao B1: " + ErrorDescription + "\nCodigo Erro B1: " + iRetCode.ToString());
                    }
                    else
                    {
                        return oCompany;
                    }
                }
                else
                {
                    return oCompany;
                }
            }
            catch (Exception ex)
            {
                string ErrorDescription = oCompany.GetLastErrorDescription();
                throw new Exception("Falha! \nErro: " + ex.Message + "\nDescricao B1: " + ErrorDescription + "\nCodigo Erro B1: " + iRetCode.ToString());
            }
        }
        #endregion

        #region Desconectar
        public void Desconectar(Company oCompany)
        {
            try
            {
                if (oCompany.Connected)
                {
                    oCompany.Disconnect();
                }
            }
            catch (Exception ex)
            {
                string ErrorDescription = oCompany.GetLastErrorDescription();
                throw new Exception("Falha! \nErro ao desconectar da company: " + ex.Message + "\nDescricao B1: " + ErrorDescription);
            }
        }
        #endregion

        #region Verifica Conexão
        public bool VerificaConexao(Company oCompany)
        {
            try
            {
                return oCompany.Connected;
            }
            catch(Exception ex)
            {
                string ErrorDescription = oCompany.GetLastErrorDescription();
                throw new Exception("Falha! \nErro ao desconectar da company: " + ex.Message + "\nDescricao B1: " + ErrorDescription);
            }
        }
        #endregion

        #region Rotina para a criação de campos de usuário
        public void CriaCampo(Company oCompany, string sTabela, string sCampo, string sDescricao, BoFieldTypes Tipo, BoFldSubTypes SubTipo, int iComprimento, BoYesNoEnum Obrigatorio, bool bValorPadrao, string sValorPadrao, bool bValoresValidos, string[,] sValoresValidos, out string sMensagem)
        {
            sMensagem = string.Empty;

            try
            {
                if (!VerificaExistenciaCampo(oCompany, sTabela, sCampo))
                {
                    int lErrCode;
                    string lErrStr;
                    int lRetVal;

                    UserFieldsMD oUserFieldMD = null;

                    //Limpa a memória
                    oUserFieldMD = (UserFieldsMD)oCompany.GetBusinessObject(BoObjectTypes.oUserFields);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD);
                    oUserFieldMD = null;
                    GC.Collect();
                    oUserFieldMD = (UserFieldsMD)oCompany.GetBusinessObject(BoObjectTypes.oUserFields);
                    oUserFieldMD.TableName = sTabela;
                    oUserFieldMD.Name = sCampo;
                    oUserFieldMD.Description = sDescricao;
                    oUserFieldMD.Type = Tipo;
                    oUserFieldMD.SubType = SubTipo;
                    oUserFieldMD.Mandatory = Obrigatorio;

                    if (bValoresValidos)
                    {
                        for (int i = 0; i < sValoresValidos.GetLength(0); i++)
                        {
                            oUserFieldMD.ValidValues.Value = sValoresValidos[i, 0];
                            oUserFieldMD.ValidValues.Description = sValoresValidos[i, 1];
                            oUserFieldMD.ValidValues.Add();
                        }
                    }

                    if (bValorPadrao)
                        oUserFieldMD.DefaultValue = sValorPadrao;

                    if (iComprimento > 0)
                        oUserFieldMD.Size = iComprimento;

                    lRetVal = oUserFieldMD.Add();
                    if (lRetVal != 0)
                    {
                        oCompany.GetLastError(out lErrCode, out lErrStr);
                        sMensagem = string.Format("Falha! Descricao B1: {0} - Codigo Erro B1: {1}", lErrStr , lErrCode.ToString());
                    }
                    else
                    {
                        sMensagem = string.Format("Campo {0} - {1}, criado com sucesso..", sCampo, sDescricao);
                    }

                    //Limpa a Memória
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD);
                    oUserFieldMD = null;
                    GC.Collect();
                }
            }
            catch(Exception ex)
            {
                sMensagem = string.Format("Erro ao criar Tabela/Campos SAP, Erro: {0}", ex.Message);
            }
        }
        #endregion

        #region Verifica se o Campo já existe na tabela - Caso não exista o mesmo cria.
        public bool VerificaExistenciaCampo(Company oCompany, string sTabela, string sCampo)
        {
            Recordset oRecordset = null;
            oRecordset = (Recordset)(oCompany.GetBusinessObject(BoObjectTypes.BoRecordset));

            string sSQL = " SELECT " +
                            "   COUNT(*) " +
                            " FROM CUFD " +
                            " WHERE TableID  = '" + sTabela + "' " +
                            " AND AliasID = '" + sCampo + "'";

            oRecordset.DoQuery(sSQL);

            if (Convert.ToInt64(oRecordset.Fields.Item(0).Value.ToString()) > 0)
                return true;
            else
                return false;
        }
        #endregion

        #region Rotina para adicionar tabelas de usuário AddUserTable
        public void AddUserTable(Company oCompany, string Name, string Description, BoUTBTableType Type)
        {
            try
            {
                int lErrCode = 0;
                string lErrStr = string.Empty;
                string LastMsg = string.Empty;

                UserTablesMD oUserTablesMD = null;

                oUserTablesMD = (UserTablesMD)oCompany.GetBusinessObject(BoObjectTypes.oUserTables);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                oUserTablesMD = null;
                GC.Collect();
                oUserTablesMD = (UserTablesMD)oCompany.GetBusinessObject(BoObjectTypes.oUserTables);
                lErrCode = 0;

                if (oUserTablesMD.GetByKey(Name) == false)
                {
                    oUserTablesMD.TableName = Name;
                    oUserTablesMD.TableDescription = Description;
                    oUserTablesMD.TableType = Type;

                    if ((lErrCode = oUserTablesMD.Add()) != 0)
                    {
                        oCompany.GetLastError(out lErrCode, out lErrStr);
                        throw new Exception("Falha! \nDescricao B1: " + lErrStr + "\nCodigo Erro B1: " + lErrCode.ToString());
                    }

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                    oUserTablesMD = null;
                    GC.Collect();

                    if (lErrCode != 0 && lErrCode != -2035)
                    {
                        oCompany.GetLastError(out lErrCode, out lErrStr);
                        throw new Exception("Falha! \nDescricao B1: " + lErrStr + "\nCodigo Erro B1: " + lErrCode.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Erro ao criar tabela de usuário: " + ex.Message);
            }
        }
        #endregion

        #region Registrar o UDO
        public void addUDO(Company oCompany, string sObjeto, string sCode, string sName, string sTableName, BoUDOObjType TipoObjeto,
                            BoYesNoEnum CanFind, BoYesNoEnum CanDelete, BoYesNoEnum CanCancel, BoYesNoEnum CanYearTransfer, BoYesNoEnum CanLog, 
                            string sLogTableName)
        {
            UserObjectsMD oUserObjectMD = null;
            oUserObjectMD = (UserObjectsMD)oCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD);

            int lRetVal;
            string lErrMsg = string.Empty;

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
            oUserObjectMD = null;
            GC.Collect();
            oUserObjectMD = (UserObjectsMD)oCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD);

            if (oUserObjectMD.GetByKey(sObjeto) == false)
            {
                oUserObjectMD.Code = sCode;
                oUserObjectMD.Name = sName;
                oUserObjectMD.TableName = sTableName;

                oUserObjectMD.ObjectType = TipoObjeto;

                oUserObjectMD.CanFind = CanFind;
                oUserObjectMD.CanDelete = CanDelete;
                oUserObjectMD.CanCancel = CanCancel;
                oUserObjectMD.CanYearTransfer = CanYearTransfer;
                oUserObjectMD.CanLog = CanLog;
                oUserObjectMD.LogTableName = sLogTableName;

                lRetVal = oUserObjectMD.Add();

                if (lRetVal != 0)
                {
                    oCompany.GetLastError(out lRetVal, out lErrMsg);
                }
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
            oUserObjectMD = null;
            GC.Collect();
        }
        #endregion

    }
}

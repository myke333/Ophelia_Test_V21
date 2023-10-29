using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.OracleClient;
using System.Diagnostics;

namespace Ophelia_Test_V21.AFuncionesGlobales
{
    class SqlAdapter
    {
        static public void InsertExecutionOrder(string plan, string suite, string order)
        {
            try
            {
                string ConnectionString = ConfigurationManager.ConnectionStrings["SA"].ConnectionString;

                SqlConnection cnn = new SqlConnection(ConnectionString);
                SqlCommand cmd = new SqlCommand("InsertOrderExecution", cnn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@plans", SqlDbType.Char, 50);
                cmd.Parameters.Add("@suites", SqlDbType.Char, 50);
                cmd.Parameters.Add("@orders", SqlDbType.Char, 50);
                cmd.Parameters["@plans"].Value = plan;
                cmd.Parameters["@suites"].Value = suite;
                cmd.Parameters["@orders"].Value = order;
                cnn.Open();
                cmd.ExecuteNonQuery();
                cnn.Close();
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        static public void UdpOPSessio(string user, string Ip, string Enc)
        {
            try
            {
                string ConnectionString = ConfigurationManager.ConnectionStrings[user].ConnectionString;

                SqlConnection cnn = new SqlConnection();
                cnn = new SqlConnection(ConnectionString);

                SqlCommand cmd = new SqlCommand("UdpOPSessio", cnn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@Enc", SqlDbType.Char, 128);
                cmd.Parameters.Add("@Ip", SqlDbType.Char, 50);

                cmd.Parameters["@Enc"].Value = Enc;
                cmd.Parameters["@Ip"].Value = Ip;

                cnn.Open();
                cmd.ExecuteNonQuery();
                cnn.Close();
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        static public void UdpOPSessioV2(string user, string Ip, string Enc, string database)
        {
            try
            {
                
                string ConnectionString2 = ConfigurationManager.ConnectionStrings[user].ConnectionString;
                if (database == "SQL")
                {
                    SqlConnection cnn = new SqlConnection(ConnectionString2);
                    string sql = string.Format("UPDATE OP_SESIO SET SES_TPRG='{0}' WHERE SES_IPAD = '{1}'", Enc, Ip);
                    cnn.Open();
                    SqlCommand cmd = new SqlCommand(sql, cnn);
                    cmd.CommandType = CommandType.Text;
                    cmd.ExecuteNonQuery();
                    cnn.Close();
                }
                else if (database == "ORA")
                {
                    OracleConnection cnn = new OracleConnection(ConnectionString2);
                    cnn.Open();
                    OracleCommand command = cnn.CreateCommand();
                    OracleTransaction transaction;
                    transaction = cnn.BeginTransaction(IsolationLevel.ReadCommitted);
                    command.Transaction = transaction;
                    try
                    {
                        string sql = string.Format("UPDATE OP_SESIO SET SES_TPRG='{0}' WHERE SES_IPAD = '{1}'", Enc, Ip);
                        command.CommandText = sql;
                        command.ExecuteNonQuery();
                        transaction.Commit();
                        cnn.Close();
                    }
                    catch (Exception e)
                    {
                        transaction.Rollback();
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        static public DataSet SelectExecutionOrder(string p, string plan = null, string suite = null, string CountDes = null)
        {
            try
            {
                string ConnectionString = ConfigurationManager.ConnectionStrings["SA"].ConnectionString;
                DataSet Data = new DataSet();
                SqlConnection cnn = new SqlConnection(ConnectionString);
                cnn.Open();
                SqlCommand cmd = new SqlCommand("SelecOrderExecution", cnn);

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@p", SqlDbType.Char, 50);
                cmd.Parameters.Add("@Suite", SqlDbType.Char, 50);
                cmd.Parameters.Add("@Plan", SqlDbType.Char, 50);
                cmd.Parameters.Add("@CountDes", SqlDbType.Char, 50);
                cmd.Parameters["@p"].Value = p;
                cmd.Parameters["@Suite"].Value = suite;
                cmd.Parameters["@Plan"].Value = plan;
                cmd.Parameters["@CountDes"].Value = CountDes;


                DataTable dTable = new DataTable("dTable");
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                adapter.Fill(dTable);

                Data.Tables.Add(dTable);
                cnn.Close();
                return Data;
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return null;
            }
        }

        static public DataSet SelectExecutionOrderV2(string p, string plan = null, string suite = null, string CountDes = null, string CaseID = null)
        {
            try
            {
                string ConnectionString = ConfigurationManager.ConnectionStrings["SA"].ConnectionString;
                DataSet Data = new DataSet();
                SqlConnection cnn = new SqlConnection(ConnectionString);
                SqlCommand cmd = new SqlCommand("SelecOrderExecutionV2", cnn);

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@p", SqlDbType.Char, 50);
                cmd.Parameters.Add("@Suite", SqlDbType.Char, 50);
                cmd.Parameters.Add("@Plan", SqlDbType.Char, 50);
                cmd.Parameters.Add("@CaseID", SqlDbType.Char, 50);
                cmd.Parameters.Add("@CountDes", SqlDbType.Char, 50);
                cmd.Parameters["@p"].Value = p;
                cmd.Parameters["@Suite"].Value = suite;
                cmd.Parameters["@Plan"].Value = plan;
                cmd.Parameters["@CaseID"].Value = CaseID;
                cmd.Parameters["@CountDes"].Value = CountDes;


                DataTable dTable = new DataTable("dTable");
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                adapter.Fill(dTable);

                Data.Tables.Add(dTable);
                return Data;
            }

            catch (Exception ex)
            {
                return null;
                Console.WriteLine(ex);
            }
        }

        public static DataSet ConsultaSqlOra(string tabla, string condicion, string database, string user, string campo = null, string condicion_2 = null)
        {
            string ConnectionString2 = ConfigurationManager.ConnectionStrings[user].ConnectionString;

            if (database == "SQL")
            {
                #region Declaracion_Variables
                string sql, sql2, numero = "";
                sql2 = string.Format("Select DATA_TYPE from information_schema.columns WHERE TABLE_NAME='{0}' AND COLUMN_NAME='{1}'", tabla, campo);
                SqlConnection cnn = new SqlConnection(ConnectionString2);
                cnn.Open();
                DataSet DataSql;
                SqlCommand cmd;
                DataTable dTable;
                SqlDataAdapter adapter;
                #endregion

                if (campo != "" && campo != null && condicion_2 != "" && condicion_2 != null)
                {
                    DataSql = new DataSet();
                    cmd = new SqlCommand(sql2, cnn);
                    cmd.CommandType = CommandType.Text;
                    dTable = new DataTable("dTable");
                    adapter = new SqlDataAdapter(cmd);
                    adapter.Fill(dTable);
                    DataSql.Tables.Add(dTable);

                    foreach (DataRow fila in DataSql.Tables[0].Rows)
                    {
                        numero = fila["DATA_TYPE"].ToString();
                    }

                    switch (numero)
                    {
                        case "text":
                        case "varchar":
                        case "char":
                        case "nchar ":
                        case "nvarchar":
                        case "ntext":
                        case "binary":
                        case "varbinary":
                        case "image":
                        case "time":
                        case "datetime":
                        case "smalldatetime":
                        case "date":
                        case "timestamp":
                        case "datetimeoffset":
                        case "sql_variant":
                            condicion_2 = "'" + condicion_2 + "'";
                            break;
                    }

                    if (tabla.ToUpper() == "GN_EMPRE")
                    {
                        sql = string.Format("SELECT * FROM {0} WHERE {1} = '{2}' ", tabla,campo,condicion_2);
                    } else
                    {
                        sql = string.Format("SELECT * FROM {0} WHERE ACT_USUA = '{1}' and {2} = {3}", tabla, condicion, campo, condicion_2);
                    }
                    
                }
                else
                {
                    //A la consulta general se le está añadiendo la comprobacion de la hora para evitar confuciones en las tablas de detalle
                    sql = string.Format("SELECT * FROM {0} WHERE ACT_USUA = '{1}'", tabla, condicion);

                    if (tabla.ToUpper() == "GN_EMPRE")
                    {
                        sql = string.Format("SELECT * FROM {0}", tabla);
                    }

                }

                DataSql = new DataSet();
                cmd = new SqlCommand(sql, cnn);
                cmd.CommandType = CommandType.Text;
                dTable = new DataTable("dTable");
                adapter = new SqlDataAdapter(cmd);
                adapter.Fill(dTable);
                DataSql.Tables.Add(dTable);
                cnn.Close();
                return DataSql;
            }
            else if (database == "ORA")
            {
                OracleConnection cnn = new OracleConnection(ConnectionString2);
                string sql, sql2, numero = "";
                sql2 = string.Format("Select DATA_TYPE from all_tab_columns WHERE TABLE_NAME='{0}' AND COLUMN_NAME='{1}'", tabla, campo);
                cnn.Open();
                DataSet DataOra = new DataSet();
                DataTable dTable = new DataTable();
                if (campo != "" && campo != null && condicion_2 != "" && condicion_2 != null)
                {
                    DataOra = new DataSet();
                    OracleDataAdapter adapter = new OracleDataAdapter(sql2, cnn);
                    dTable = new DataTable("dTable");
                    adapter.Fill(dTable);
                    DataOra.Tables.Add(dTable);

                    foreach (DataRow fila in DataOra.Tables[0].Rows)
                    {
                        numero = fila["DATA_TYPE"].ToString();
                    }

                    switch (numero)
                    {
                        case "TEXT":
                        case "VARCHAR":
                        case "CHAR":
                        case "NCHAR ":
                        case "NVARCHAR":
                        case "NTEXT":
                        case "BINARY":
                        case "VARBINARY":
                        case "IMAGE":
                        case "TIME":
                        case "DATATIME":
                        case "SMALLDATETIME":
                        case "DATE":
                        case "TIMESTAMP":
                        case "DATATIMEOFFSET":
                        case "VARCHAR2":
                            condicion_2 = "'" + condicion_2 + "'";
                            break;
                    }
                }
                if (tabla.ToUpper() == "GN_DIVIP")
                {
                    sql = string.Format("SELECT * from GN_DIVIP where ACT_USUA='TNATALIA' and COD_DPTO=123");
                }else if (tabla.ToUpper() == "GN_EMPRE")
                {
                    sql = string.Format("SELECT * FROM {0} WHERE {1} = {2} ", tabla, campo, condicion_2);
                }
                else if (tabla.ToUpper() == "GN_LOGTW")
                {
                    sql = string.Format("SELECT * FROM GN_LOGTW WHERE ACT_USUA = '19301797'");
                }else
                {
                    sql = string.Format("SELECT * FROM {0} WHERE ACT_USUA = '{1}' and {2} = {3}", tabla, condicion,campo, condicion_2);
                }
                OracleDataAdapter adaptera = new OracleDataAdapter(sql, cnn);
                DataOra = new DataSet();
                dTable = new DataTable("dTable");
                adaptera.Fill(dTable);
                DataOra.Tables.Add(dTable);
                cnn.Close();
                return DataOra;
            }
            return null;
        }

        public static DataSet ConsultaSqlOraAuditoria(string tabla, string database, string user, string campo, string valcamp)
        {
            //string {0}tabla, string {1}condicion, string {2}database, {3}string user, string {4}campo = null, string {5}condicion_2 = null
            string ConnectionString2 = ConfigurationManager.ConnectionStrings[user].ConnectionString;

            if (database == "SQL")
            {
                #region Declaracion_Variables
                string sql, sql2, numero = "";
                sql2 = string.Format("Select DATA_TYPE from information_schema.columns WHERE TABLE_NAME='{0}' AND COLUMN_NAME='{1}'", tabla, campo);
                SqlConnection cnn = new SqlConnection(ConnectionString2);
                cnn.Open();
                DataSet DataSql;
                SqlCommand cmd;
                DataTable dTable;
                SqlDataAdapter adapter;
                #endregion
                if (campo != "" && campo != null && valcamp != "" && valcamp != null)
                {
                    DataSql = new DataSet();
                    cmd = new SqlCommand(sql2, cnn);
                    cmd.CommandType = CommandType.Text;
                    dTable = new DataTable("dTable");
                    adapter = new SqlDataAdapter(cmd);
                    adapter.Fill(dTable);
                    DataSql.Tables.Add(dTable);

                    foreach (DataRow fila in DataSql.Tables[0].Rows)
                    {
                        numero = fila["DATA_TYPE"].ToString();
                    }

                    switch (numero)
                    {
                        case "text":
                        case "varchar":
                        case "char":
                        case "nchar ":
                        case "nvarchar":
                        case "ntext":
                        case "binary":
                        case "varbinary":
                        case "image":
                        case "time":
                        case "datetime":
                        case "smalldatetime":
                        case "date":
                        case "timestamp":
                        case "datetimeoffset":
                        case "sql_variant":
                            valcamp = "'" + valcamp + "'";
                            break;
                    }
                    //sql = string.Format("SELECT * FROM {0} WHERE ACT_USUA = '{1}' and {2} = {3}", tabla, condicion, campo, condicion_2);
                    
                }
                sql = string.Format("SELECT * FROM GN_LOGBO WHERE LOG_FECH BETWEEN dateadd(second, -1,(SELECT act_hora FROM  {0} WHERE act_usua = '{1}' and {2}={3})) and dateadd(second, 90, (SELECT act_hora FROM {0} WHERE act_usua = '{1}' and {2} = {3})) AND LOG_TABL = '{0}';", tabla, user, campo, valcamp);

                DataSql = new DataSet();
                cmd = new SqlCommand(sql, cnn);
                cmd.CommandType = CommandType.Text;
                dTable = new DataTable("dTable");
                adapter = new SqlDataAdapter(cmd);
                adapter.Fill(dTable);
                DataSql.Tables.Add(dTable);
                cnn.Close();
                return DataSql;
            }
            else if (database == "ORA")
            {

                OracleConnection cnn = new OracleConnection(ConnectionString2);
                cnn.Open();

                string sql, sql2, numero = "";
                sql2 = string.Format("Select DATA_TYPE from all_tab_columns WHERE TABLE_NAME='{0}' AND COLUMN_NAME='{1}'", tabla, campo);

                DataSet DataOra = new DataSet();
                DataTable dTable = new DataTable();
                if (campo != "" && campo != null && valcamp != "" && valcamp != null)
                {
                    DataOra = new DataSet();
                    OracleDataAdapter adaptera = new OracleDataAdapter(sql2, cnn);
                    dTable = new DataTable("dTable");
                    adaptera.Fill(dTable);
                    DataOra.Tables.Add(dTable);

                    foreach (DataRow fila in DataOra.Tables[0].Rows)
                    {
                        numero = fila["DATA_TYPE"].ToString();
                    }

                    switch (numero)
                    {
                        case "TEXT":
                        case "VARCHAR":
                        case "CHAR":
                        case "NCHAR ":
                        case "NVARCHAR":
                        case "NTEXT":
                        case "BINARY":
                        case "VARBINARY":
                        case "IMAGE":
                        case "TIME":
                        case "DATATIME":
                        case "SMALLDATETIME":
                        case "DATE":
                        case "TIMESTAMP":
                        case "DATATIMEOFFSET":
                        case "VARCHAR2":
                            valcamp = "'" + valcamp + "'";
                            break;
                    }
                }

                //sql = string.Format("SELECT * FROM {0} WHERE ACT_USUA = '{1}'", tabla, condicion);
                sql = string.Format("SELECT * FROM GN_LOGBO WHERE LOG_FECH BETWEEN (SELECT CAST(act_hora - 30 / 86400 as TIMESTAMP) FROM {0} WHERE {1}={2} AND ACT_USUA = '{3}' ) AND (SELECT CAST(act_hora + 240 / 86400 as TIMESTAMP) FROM {0} WHERE {1}= {2} AND ACT_USUA = '{3}') AND LOG_TABL='{0}'", tabla, campo, valcamp, user);
                OracleDataAdapter adapter = new OracleDataAdapter(sql, cnn);
                DataOra = new DataSet();
                dTable = new DataTable("dTable");
                adapter.Fill(dTable);
                DataOra.Tables.Add(dTable);
                cnn.Close();
                return DataOra;
            }
            return null;
        }

        public static DataSet ConsultaSqlOraAuditoriaDelete(string database, string user, string tabla)
        {
            //string {0}tabla, string {1}condicion, string {2}database, {3}string user, string {4}campo = null, string {5}condicion_2 = null
            string ConnectionString2 = ConfigurationManager.ConnectionStrings[user].ConnectionString;

            if (database == "SQL")
            {
                #region Declaracion_Variables
                string sql, sql2, numero = "";
                sql2 = string.Format("Select DATA_TYPE from information_schema.columns WHERE TABLE_NAME='GN_LOGBO' AND COLUMN_NAME= 'LOG_TIPO'");
                SqlConnection cnn = new SqlConnection(ConnectionString2);
                cnn.Open();
                DataSet DataSql;
                SqlCommand cmd;
                DataTable dTable;
                SqlDataAdapter adapter;
                #endregion

                //sql = string.Format("SELECT * FROM {0} WHERE ACT_USUA = '{1}' and {2} = {3}", tabla, condicion, campo, condicion_2);
                sql = string.Format("SELECT LOG_TIPO from GN_LOGBO where LOG_USUA='{0}' and LOG_TABL = '{1}' and LOG_FECH between dateadd(second, -20,SYSDATETIME()) and SYSDATETIME();", user, tabla);


                DataSql = new DataSet();
                cmd = new SqlCommand(sql, cnn);
                cmd.CommandType = CommandType.Text;
                dTable = new DataTable("dTable");
                adapter = new SqlDataAdapter(cmd);
                adapter.Fill(dTable);
                DataSql.Tables.Add(dTable);
                cnn.Close();
                return DataSql;
            }
            else if (database == "ORA")
            {
                OracleConnection cnn = new OracleConnection(ConnectionString2);
                cnn.Open();

                DataSet DataOra = new DataSet();
                DataTable dTable = new DataTable();

                string sql;

                //sql = string.Format("SELECT * FROM {0} WHERE ACT_USUA = '{1}'", tabla, condicion);
                sql = string.Format("SELECT LOG_TIPO FROM GN_LOGBO WHERE LOG_USUA= '{0}' and LOG_TABL = '{1}' AND LOG_FECH BETWEEN SYSDATE - 60/86400 and SYSDATE", user, tabla);
                OracleDataAdapter adapter = new OracleDataAdapter(sql, cnn);
                adapter.Fill(dTable);
                DataOra.Tables.Add(dTable);
                cnn.Close();
                return DataOra;
            }
            return null;
        }

        public static void ActualizarDigiFlag_Dia31(string DigiFlag, string Dia31, string IDUsuario, string user, string database)
        {
            try
            {
                string ConnectionString2 = ConfigurationManager.ConnectionStrings[user].ConnectionString;
                if (database == "SQL")
                {
                    string sql = string.Format("declare{0}" +
                                                "@codgpro smallint;{1}" +
                                                "select top 1 @codgpro=cod_gpro from nm_contr where cod_empr=557 and cod_empl={2} order by fec_ingr desc;{3}" +
                                                "UPDATE GN_DIGIF SET VAL_VARI='{4}' WHERE COD_VARI='K0000120';{5}" +
                                                "UPDATE NM_PASSO SET DES_DI31='{6}' WHERE COD_GPRO= @codgpro;", Environment.NewLine,
                                                Environment.NewLine, IDUsuario, Environment.NewLine, DigiFlag, Environment.NewLine, Dia31);

                    SqlConnection cnn = new SqlConnection(ConnectionString2);
                    cnn.Open();
                    SqlCommand command = cnn.CreateCommand();
                    SqlTransaction transaction;
                    transaction = cnn.BeginTransaction("Update DigiFlag y Dia 31");
                    command.Connection = cnn;
                    command.Transaction = transaction;
                    try
                    {
                        command.CommandText = sql;
                        command.ExecuteNonQuery();
                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        Console.WriteLine(ex.ToString());
                    }
                    cnn.Close();
                }
                else if (database == "ORA")
                {
                    string sql = string.Format("declare{0}" +
                                                "codgpro number(5); {1}" +
                                                "begin{2}" +
                                                "SELECT COD_GPRO into codgpro FROM NM_CONTR WHERE ROWNUM <= 1 and cod_empr=557 and cod_empl={3} order by fec_ingr desc; {4}" +
                                                "UPDATE GN_DIGIF SET VAL_VARI='{5}' WHERE COD_VARI='K0000120'; {6}" +
                                                "UPDATE NM_PASSO SET DES_DI31='{7}' WHERE COD_GPRO=codgpro; {8}" +
                                                "end;", Environment.NewLine, Environment.NewLine, Environment.NewLine, IDUsuario,
                                                Environment.NewLine, DigiFlag, Environment.NewLine, Dia31, Environment.NewLine);

                    OracleConnection cnn = new OracleConnection(ConnectionString2);
                    cnn.Open();
                    OracleCommand command = cnn.CreateCommand();
                    OracleTransaction transaction;
                    transaction = cnn.BeginTransaction(IsolationLevel.ReadCommitted);
                    command.Transaction = transaction;
                    try
                    {
                        command.CommandText = sql;
                        command.ExecuteNonQuery();
                        transaction.Commit();
                    }
                    catch (Exception e)
                    {
                        transaction.Rollback();
                        Console.WriteLine(e.ToString());
                    }
                    cnn.Close();
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }


        static public DataSet SelectOrderExecution(string Parameter, string Table, string plan = null, string suite = null, string CaseID = null, string CountDes = null)
        {
            string ConnectionString2 = ConfigurationManager.ConnectionStrings["SA"].ConnectionString;
            try
            {
               
                string SentenciaSQL = null;
                DataSet DataSql = new DataSet();
                SqlConnection cnn = new SqlConnection(ConnectionString2);

                if (Parameter == "T")
                {
                    SentenciaSQL = string.Format("SELECT * FROM {0}", Table);
                }
                else if (Parameter == "P")
                {
                    SentenciaSQL = string.Format("SELECT * FROM {0} WHERE plans={1} AND suite={2} AND CaseID={3}", Table, plan, suite, CaseID);
                }
                else if (Parameter == "U")
                {
                    SentenciaSQL = string.Format("DELETE FROM {0} WHERE plans = {1} AND suite = {2} AND CaseID={3}", Table, plan, suite, CaseID);
                }
                else if (Parameter == "UP")
                {
                    SentenciaSQL = string.Format("UPDATE {0} SET CountDes={1} FROM {2} WHERE plans = {3} AND suite = {4} AND CaseID={5}", Table, CountDes, Table, plan, suite, CaseID);
                }

                SqlCommand cmd = new SqlCommand(SentenciaSQL, cnn);
                cmd.CommandType = CommandType.Text;
                DataTable dTable = new DataTable("dTable");
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                adapter.Fill(dTable);
                DataSql.Tables.Add(dTable);
                return DataSql;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return null;
            }
        }
        static public DataSet SelectOrderExecutionV2(string Parameter, string Table, string plan = null, string suite = null, string CaseID = null, string CountDes = null)
        {
            try
            {
                string SentenciaSQL = null;
                DataSet DataSql = new DataSet();
                string ConnectionString = ConfigurationManager.ConnectionStrings["SA"].ConnectionString;
                SqlConnection cnn = new SqlConnection(ConnectionString);
                cnn.Open();
                if (Parameter == "T")
                {
                    SentenciaSQL = string.Format("SELECT * FROM {0}", Table);
                }
                else if (Parameter == "P")
                {
                    SentenciaSQL = string.Format("SELECT * FROM {0} WHERE plans={1} AND suite={2} AND CaseID={3}", Table, plan, suite, CaseID);
                }
                else if (Parameter == "U")
                {
                    SentenciaSQL = string.Format("DELETE FROM {0} WHERE plans = {1} AND suite = {2} AND CaseID={3}", Table, plan, suite, CaseID);
                }
                else if (Parameter == "UP")
                {
                    SentenciaSQL = string.Format("UPDATE {0} SET CountDes={1} WHERE plans = {2} AND suite = {3} AND CaseID={4}", Table, CountDes, plan, suite, CaseID);
                }
                SqlCommand cmd = new SqlCommand(SentenciaSQL, cnn);
                cmd.CommandType = CommandType.Text;
                DataTable dTable = new DataTable("dTable");
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                adapter.Fill(dTable);
                DataSql.Tables.Add(dTable);
                cnn.Close();
                return DataSql;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return null;
            }
        }

        static public void ModificaGNPAKAC(string TipoAlmacenamiento, string Database, string User, string CodEmpresa, string RutaFTP, string UserFTP, string PassFTP, string RutaCP, string PesoArch, string PortFTP, string ServerFTP)
        {
            string ConnectionString = ConfigurationManager.ConnectionStrings[User].ConnectionString;
            string Sentencia = "";
            string getdate = "";
            if (Database == "SQL")
            {
                getdate = "GETDATE()";
            }
            else if (Database == "ORA")
            {
                getdate = "sysdate";
            }
            switch (TipoAlmacenamiento.ToUpper())
            {
                case "F":
                    Sentencia = $@"UPDATE GN_PAKAC SET ACT_USUA='{User}', ACT_HORA={getdate}, ACT_ESTA='M', COD_EMPR={CodEmpresa}, USA_RFTP='S', FTP_RUTA='{RutaFTP}', 
                                FTP_USUA='DIGITALWARE\{UserFTP}', FTP_PASS='{PassFTP}', ARC_PHAT=NULL, PES_ARCH={PesoArch},
                                FTP_PUER={PortFTP}, FTP_HOST='{ServerFTP}', USA_BASD='N', CAS_CONT=NULL WHERE COD_EMPR={CodEmpresa}";
                    break;
                case "B":
                    Sentencia = $@"UPDATE GN_PAKAC SET ACT_USUA='{User}', ACT_HORA={getdate}, ACT_ESTA='M', COD_EMPR={CodEmpresa}, USA_RFTP='N', FTP_RUTA=NULL, 
                                FTP_USUA = NULL, FTP_PASS = NULL, ARC_PHAT=NULL, PES_ARCH={PesoArch},
                                FTP_PUER = NULL, FTP_HOST = NULL, USA_BASD = 'S', CAS_CONT = NULL WHERE COD_EMPR={CodEmpresa}";
                    break;
                case "C":
                    Sentencia = $@"UPDATE GN_PAKAC SET ACT_USUA='{User}', ACT_HORA={getdate}, ACT_ESTA='M', COD_EMPR={CodEmpresa}, USA_RFTP='N', FTP_RUTA=NULL, 
                                FTP_USUA=NULL, FTP_PASS=NULL, ARC_PHAT='{RutaCP}', PES_ARCH={PesoArch},
                                FTP_PUER=NULL, FTP_HOST=NULL, USA_BASD='N', CAS_CONT=NULL WHERE COD_EMPR={CodEmpresa}";
                    break;
                default:
                    break;
            }
            switch (Database.ToUpper())
            {
                case "SQL":

                    SqlConnection sqlConnection = new SqlConnection(ConnectionString);
                    sqlConnection.Open();
                    SqlCommand sqlCommand = sqlConnection.CreateCommand();
                    SqlTransaction sqlTransaction;
                    sqlTransaction = sqlConnection.BeginTransaction();
                    sqlCommand.Connection = sqlConnection;
                    sqlCommand.Transaction = sqlTransaction;
                    try
                    {
                        sqlCommand.CommandText = Sentencia;
                        sqlCommand.ExecuteNonQuery();
                        sqlTransaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        sqlTransaction.Rollback();
                        Console.WriteLine(ex.ToString());
                    }
                    sqlConnection.Close();
                    break;

                case "ORA":

                    OracleConnection oracleConnection = new OracleConnection(ConnectionString);
                    oracleConnection.Open();
                    OracleCommand oracleCommand = oracleConnection.CreateCommand();
                    OracleTransaction oracleTransaction;
                    oracleTransaction = oracleConnection.BeginTransaction();
                    oracleCommand.Connection = oracleConnection;
                    oracleCommand.Transaction = oracleTransaction;
                    try
                    {
                        oracleCommand.CommandText = Sentencia;
                        oracleCommand.ExecuteNonQuery();
                        oracleTransaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        oracleTransaction.Rollback();
                        Console.WriteLine(ex.ToString());
                    }
                    oracleConnection.Close();
                    break;
                default:
                    break;
            }



        }
        public static void DeleteSqlOra(string tabla, string condicion, string database, string user, string campo = null, string condicion_2 = null)
        {
            string ConnectionString2 = ConfigurationManager.ConnectionStrings[user].ConnectionString;
            

            if( database== "SQL")
            {
                string sql2, numero = "";
                sql2 = string.Format("Select DATA_TYPE from information_schema.columns WHERE TABLE_NAME='{0}' AND COLUMN_NAME='{1}'", tabla, campo);
                SqlConnection cnn = new SqlConnection(ConnectionString2);
                cnn.Open();
                DataSet DataSql;
                SqlCommand cmd;
                DataTable dTable;
                SqlDataAdapter adapter;
                if (campo != "" && campo != null && condicion_2 != "" && condicion_2 != null)
                {
                    DataSql = new DataSet();
                    cmd = new SqlCommand(sql2, cnn);
                    cmd.CommandType = CommandType.Text;
                    dTable = new DataTable("dTable");
                    adapter = new SqlDataAdapter(cmd);
                    adapter.Fill(dTable);
                    DataSql.Tables.Add(dTable);


                    foreach (DataRow fila in DataSql.Tables[0].Rows)
                    {
                        numero = fila["DATA_TYPE"].ToString();
                    }

                    switch (numero)
                    {
                        case "text":
                        case "varchar":
                        case "char":
                        case "nchar ":
                        case "nvarchar":
                        case "ntext":
                        case "binary":
                        case "varbinary":
                        case "image":
                        case "time":
                        case "datetime":
                        case "smalldatetime":
                        case "date":
                        case "timestamp":
                        case "datetimeoffset":
                        case "sql_variant":
                            condicion_2 = "'" + condicion_2 + "'";
                            break;
                    }
                    string sql = "";
                    sql = string.Format("DELETE FROM {0} WHERE ACT_USUA = '{1}' and {2} = {3}", tabla, condicion, campo, condicion_2);
                    SqlCommand sqlCommand = cnn.CreateCommand();
                    SqlTransaction sqlTransaction;
                    sqlTransaction = cnn.BeginTransaction();
                    sqlCommand.Connection = cnn;
                    sqlCommand.Transaction = sqlTransaction;
                    try
                    {
                        sqlCommand.CommandText = sql;
                        sqlCommand.ExecuteNonQuery();
                        sqlTransaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        sqlTransaction.Rollback();
                        Console.WriteLine(ex.ToString());
                    }
                    cnn.Close();
                }
                   
            }
            else
            {
                OracleConnection oracleConnection = new OracleConnection(ConnectionString2);
                oracleConnection.Open();
                OracleCommand oracleCommand = oracleConnection.CreateCommand();
                OracleTransaction oracleTransaction;
                oracleTransaction = oracleConnection.BeginTransaction();
                oracleCommand.Connection = oracleConnection;
                oracleCommand.Transaction = oracleTransaction;
                string sql = "";
                sql = string.Format("DELETE FROM {0} WHERE ACT_USUA = '{1}' and {2} = {3}", tabla, condicion, campo, condicion_2);
                try
                {
                    oracleCommand.CommandText = sql;
                    oracleCommand.ExecuteNonQuery();
                    oracleTransaction.Commit();
                }
                catch (Exception ex)
                {
                    oracleTransaction.Rollback();
                    Console.WriteLine(ex.ToString());
                }
                oracleConnection.Close();
            }
        }
    }
}

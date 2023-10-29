using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Ophelia_Test_V21.AFuncionesGlobales;
using System.Diagnostics;
using System.Configuration;
using System.Data.SqlClient;
using System.Threading;
using System.Data.OracleClient;

namespace Ophelia_Test_V21.AFuncionesGlobales.ValidacionQueries
{
    class ValidacionesQueries : FuncionesVitales
    {
        public static List<string> ValidacionCabecera(string Tabla, string user, string motor, string Campo, string condicion2, string value, string nameColumn, DataSet[] c, List<string> ErrorValidacion, string message, int flag, string EliminarRegistros = null) {
            //Eliminar registros pegados
            if (EliminarRegistros == "1")
            {
                //Obtener datos de la database
                string ConnectionString2 = ConfigurationManager.ConnectionStrings[user].ConnectionString;
                string sql,sqlDelete,ora,oracleDelete;
                bool sentencia = false;
                if (motor == "SQL")
                {
                    using (SqlConnection connection = new SqlConnection(ConnectionString2))
                    {
                        try
                        {
                            //Abre la conexión
                            connection.Open();
                            //Crea la consulta
                            sql = string.Format("SELECT * FROM {0} WHERE ACT_USUA = '{1}' and {2} = '{3}'", Tabla, user, Campo, condicion2);
                            //Ejecutar sentencia
                            Thread.Sleep(1000);
                            SqlCommand comando = new SqlCommand(sql, connection);
                            SqlDataReader registro = comando.ExecuteReader();
                            Thread.Sleep(2000);
                            //Validamos que exista la consulta
                            if (registro.Read())
                            {
                                sentencia = true;
                                //Se cierra la conexión
                                connection.Close();
                            }
                            Thread.Sleep(1000);
                            if (sentencia)
                            {
                                //Abre la conexión
                                connection.Open();
                                //Crea la sentencia
                                sqlDelete = string.Format("DELETE FROM {0} WHERE ACT_USUA = '{1}' AND {2} = '{3}'", Tabla, user, Campo, condicion2);
                                //Ejecutar sentencia
                                Thread.Sleep(2000);
                                SqlCommand comandoDelete = new SqlCommand(sqlDelete, connection);
                                int cant = 0;
                                Thread.Sleep(2000);
                                cant = comandoDelete.ExecuteNonQuery();
                                if (cant > 0)
                                {
                                    Console.WriteLine("Se eliminaron " + cant + " registros correctamente");
                                }
                                else
                                {
                                    Console.WriteLine("No se elimino el registro correctamente");
                                }
                            }
                            connection.Close();

                        }
                        catch (Exception)
                        {
                            Console.WriteLine("::::ERROR:::: No se pudo realizar la conexión con la base de datos SQL Server");
                        }
                    }
                }
                else if(motor == "ORA")
                {
                    using (OracleConnection connection = new OracleConnection(ConnectionString2))
                    {
                        try
                        {
                            //Abre la conexión
                            connection.Open();
                            //Crea la consulta
                            ora = string.Format("SELECT * FROM {0} WHERE ACT_USUA = '{1}' and {2} = '{3}'", Tabla, user, Campo, condicion2);
                            //Ejecutar sentencia
                            Thread.Sleep(1000);
                            OracleCommand comando = new OracleCommand(ora, connection);
                            OracleDataReader registro = comando.ExecuteReader();
                            Thread.Sleep(2000);
                            //Validamos que exista la consulta
                            if (registro.Read())
                            {
                                sentencia = true;
                                //Se cierra la conexión
                                connection.Close();
                            }
                            Thread.Sleep(1000);
                            if (sentencia)
                            {
                                //Abre la conexión
                                connection.Open();
                                //Crea la sentencia
                                oracleDelete = string.Format("DELETE FROM {0} WHERE ACT_USUA = '{1}' AND {2} = '{3}'", Tabla, user, Campo, condicion2);
                                //Ejecutar sentencia
                                Thread.Sleep(1000);
                                OracleCommand comandoDelete = new OracleCommand(oracleDelete, connection);
                                int cant = 0;
                                Thread.Sleep(1000);
                                cant = comandoDelete.ExecuteNonQuery();
                                if (cant > 0)
                                {
                                    Console.WriteLine("Se eliminaron " + cant + " registros correctamente");
                                }
                                else
                                {
                                    Console.WriteLine("No se elimino el registro correctamente");
                                }
                            }
                            connection.Close();

                        }
                        catch (Exception)
                        {
                            Console.WriteLine("::::ERROR:::: No se pudo realizar la conexión con la base de datos Oracle");
                        }
                    }
                }
            }
            else
            {
                //0 agregar
                //1 editar
                //2eliminar
                c[0] = SqlAdapter.ConsultaSqlOra(Tabla, user, motor, user, Campo, condicion2);

                string[] variable = new string[60];
                foreach (DataRow fila in c[0].Tables[0].Rows)
                {
                    variable[0] = fila[nameColumn].ToString();

                }

                if (flag == 0)
                {
                    variable[0] = variable[0].Replace(" 12:00:00 AM", string.Empty);
                    variable[0] = variable[0].Replace(" 00:00:00 ", string.Empty);
                    variable[0] = variable[0].Replace(" ", string.Empty);
                    //variable[0] = variable[0].Replace("0", string.Empty);
                    if (value != variable[0]) ErrorValidacion.Add(string.Format(message));

                }
                else
                {
                    if (flag == 1)
                    {
                        variable[0] = variable[0].Replace(".000000", string.Empty);
                        variable[0] = variable[0].Replace(".00000", string.Empty);
                        variable[0] = variable[0].Replace(",00000", string.Empty);
                        variable[0] = variable[0].Replace(",000000", string.Empty);
                        variable[0] = variable[0].Replace(".00", string.Empty);
                        variable[0] = variable[0].Replace(",00", string.Empty);
                        variable[0] = variable[0].Replace(" 12:00:00 AM", string.Empty);
                        variable[0] = variable[0].Replace(" 00:00:00 ", string.Empty);
                        variable[0] = variable[0].Replace("0:00:00", string.Empty);
                        variable[0] = variable[0].Replace(" ", string.Empty);
                        if (variable[0].Replace(" ", string.Empty) != value.Replace(" ", string.Empty)) ErrorValidacion.Add(string.Format(message));

                    }
                    else
                    {
                        if (flag == 2)
                        {
                            if (value != "") ErrorValidacion.Add(string.Format(message));
                        }
                    }
                }
            }
            return ErrorValidacion;
        }

        public static string CodEmpresa(string Tabla, string user, string motor, string Campo, string condicion2, DataSet[] c)
        {
            c[0] = SqlAdapter.ConsultaSqlOra(Tabla, user, motor, user, Campo, condicion2);
            string cod_empr = "";
            foreach (DataRow fila in c[0].Tables[0].Rows)
            {
                cod_empr = fila["COD_EMPR"].ToString();
            }
            return cod_empr;
        }

        public static string NombreEmpresa (string user, string motor, string COD_EMPRAUX)
        {
            string cod_empr = "";
            string nom_empr = "";
            string campo1 = "COD_EMPR";
            DataSet empresa = SqlAdapter.ConsultaSqlOra("GN_EMPRE", user, motor, user,campo1,COD_EMPRAUX);
            foreach (DataRow fila in empresa.Tables[0].Rows)
            {
                nom_empr = fila["nom_empr"].ToString();
                cod_empr = fila["COD_EMPR"].ToString();
                if (cod_empr == COD_EMPRAUX)
                    break;
            }
            return nom_empr;
        }

        public static List<string> ValidacionLogbo(string Tabla, string motor, string user, string Campo, string valCamp, string tipo, DataSet[] c, List<string> ErrorValidacion) {
            string variable = "";
            if (tipo == "E")
            {
                c[0] = SqlAdapter.ConsultaSqlOraAuditoria(Tabla, motor, user, Campo, valCamp);
                if (c[0].Tables[0].Rows.Count != 0)
                {
                    foreach (DataRow fila in c[0].Tables[0].Rows)
                    {
                        variable = fila["LOG_TIPO"].ToString();
                    }
                    if ((variable.Replace(" ", string.Empty) != tipo))
                    {
                        //ErrorValidacion.Add(string.Format("EL tipo de LOG encontrado no corresponde, el esperado es {0} y el encontrado es {1}", tipo, variable));
                    }
                }
                else
                {
                    //ErrorValidacion.Add(string.Format("No se registra el log de auditoria en la edición de la cabecera, el resultado esperado era {0}", tipo));

                }
            }
            else if (tipo == "D") {
                
                c[0] = SqlAdapter.ConsultaSqlOraAuditoriaDelete(motor, user, Tabla);
                if (c[0].Tables[0].Rows.Count != 0)
                {
                    foreach (DataRow fila in c[0].Tables[0].Rows)
                    {
                        variable = fila["LOG_TIPO"].ToString();
                    }
                    if ((variable.Replace(" ", string.Empty) != tipo))
                    {
                        //ErrorValidacion.Add(string.Format("No se registra el log de auditoria en DELETE de Cabecera, el resultado esperado era {0}", tipo));
                    }
                }
                else
                {
                    //ErrorValidacion.Add(string.Format("El DELETE del registro no está generando inserción en KGn_LogBo {0}", Tabla));
                }
            }


            return ErrorValidacion;
        }

        /*
         //CONSULTA DE DATOS DE AUDITORIA GN_LOGBO (EDITAR CABECERA) 
         c[2] = SqlAdapter.ConsultaSqlOraAuditoria(Tabla, motor, user, Campo, identificacion);
          string tipo = "E";
         if (c[2].Tables[0].Rows.Count != 0)
        {
            foreach (DataRow fila in c[2].Tables[0].Rows)
            {
                VARIABLE[2] = fila["LOG_TIPO"].ToString();
            }
            if ((VARIABLE[2].Replace(" ", string.Empty) != tipo))
            {
                ErrorValidacion.Add(string.Format("EL tipo de LOG encontrado no corresponde, el esperado es {0} y el encontrado es {1}", tipo, VARIABLE[2]));
            }
        }
        else
        {
            ErrorValidacion.Add(string.Format("No se registra el log de auditoria en la edición de la cabecera, el resultado esperado era {0}", tipo));

        }
         

         //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
        string tipo2 = "D";
        c[4] = SqlAdapter.ConsultaSqlOraAuditoriaDelete(motor, user);
        if (c[4].Tables[0].Rows.Count != 0)
        {
            foreach (DataRow fila in c[4].Tables[0].Rows)
            {
                VARIABLE[4] = fila["LOG_TIPO"].ToString();
            }
            if ((VARIABLE[4].Replace(" ", string.Empty) != tipo2))
            {
                ErrorValidacion.Add(string.Format("No se registra el log de auditoria en DELETE de Cabecera, el resultado esperado era {0}", tipo2));
            }
        }
        else
        {
            ErrorValidacion.Add(string.Format("El DELETE del registro no está generando inserción en KGn_LogBo {0}", Tabla));
        }



         
         */



    }
}

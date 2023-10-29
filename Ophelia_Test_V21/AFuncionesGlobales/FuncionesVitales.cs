using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.Data.OracleClient;
using System.IO;
using System.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Management;
using System.Threading;
using System.Net;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Ophelia_Test_V21.AFuncionesGlobales
{
    public class FuncionesVitales
    {
        #region Declaracion de Variables
        protected string Mensaje;
        protected List<string> listaEsperada = new List<string>();
        protected int constante;
        protected string Modulo;
        protected static int num = 60;
        protected string[] VARIABLE = new string[num];
        protected DataSet[] c = new DataSet[num];
        protected List<string> pdf = new List<string>();
        protected List<string> word = new List<string>();
        protected List<string> excel = new List<string>();
        protected List<string> texto = new List<string>();
        #endregion


        public FuncionesVitales()
        {
            Mensaje = "Error en el registro: ";
            constante = 150;
        }

        static public void LimpiarProcesos()
        {
            Process[] processes = Process.GetProcessesByName("AcroRd32");
            if (processes.Length > 0)
            {
                for (int i = 0; i < processes.Length; i++)
                {
                    try
                    {
                        processes[i].Kill();
                    }
                    catch { }
                }
            }
            Process[] processes1 = Process.GetProcessesByName("EXCEL");
            if (processes1.Length > 0)
            {
                for (int i = 0; i < processes1.Length; i++)
                {
                    try
                    {
                        processes1[i].Kill();
                    }
                    catch { }
                }
            }
            Process[] processes2 = Process.GetProcessesByName("WINWORD");
            if (processes2.Length > 0)
            {
                for (int i = 0; i < processes2.Length; i++)
                {
                    try
                    {
                        processes2[i].Kill();
                    }
                    catch { }
                }
            }

            Process[] processes3 = Process.GetProcessesByName("notepad");
            if (processes3.Length > 0)
            {
                for (int i = 0; i < processes3.Length; i++)
                {
                    try
                    {
                        processes3[i].Kill();
                    }
                    catch { }
                }
            }

        }

        static public DataTable SelectData(string Sentencia, string user, string database)
        {
            DataSet dataSet = new DataSet();
            DataTable dTable = new DataTable("dTable");
            string ConnectionString2 = ConfigurationManager.ConnectionStrings[user].ConnectionString;

            if (database == "SQL")
            {
                SqlConnection cnn = new SqlConnection(ConnectionString2);
                cnn.Open();
                SqlDataAdapter adapter = new SqlDataAdapter(Sentencia, cnn);
                adapter.Fill(dTable);
                dataSet.Tables.Add(dTable);
                cnn.Close();
                return dTable;
            }
            else if (database == "ORA")
            {
                OracleConnection cnn = new OracleConnection(ConnectionString2);
                cnn.Open();
                OracleDataAdapter adapter = new OracleDataAdapter(Sentencia, cnn);
                adapter.Fill(dTable);
                dataSet.Tables.Add(dTable);
                cnn.Close();
                return dTable;
            }
            return null;

        }
        public static void UpdateDeleteInsert(string Sentencia, string Database, string User)
        {
            string ConnectionString = ConfigurationManager.ConnectionStrings[User].ConnectionString;
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



        static public string Hora()
        {
            DateTime dateAndTime = DateTime.Now;
            string fecha = dateAndTime.ToString("ddMMyyyy_HHmmss");
            return fecha;
        }

        static public string ValidacionesSalario(int bandera, string motor, string user, string id, string empresa, string concepto)
        {
            
                //salario base
                string salariobase = string.Empty;
                string sentencia = "select sue_basi from nm_contr where cod_empl="+ "'" + id + "'" + "and cod_empr="+"'"+empresa+"'";
                DataTable resultado = FuncionesVitales.SelectData(sentencia, user, motor);
                if (resultado.Rows.Count > 0)
                {
                    foreach (DataRow rw in resultado.Rows)
                    {
                        salariobase = rw["sue_basi"].ToString();

                    }
                }
                //cantidad dias
                string cantdias = string.Empty;
                string sentencia2 = "select can_acum from nm_preno where cod_empl=" + "'" + id + "'" + " and cod_conc in ("+concepto+")";
                DataTable resultado2 = FuncionesVitales.SelectData(sentencia2, user, motor);
                if (resultado2.Rows.Count > 0)
                {
                    foreach (DataRow rw in resultado2.Rows)
                    {
                        cantdias = rw["can_acum"].ToString();

                    }
                }
                //formulas
                string[] basesal = salariobase.Split('.',',');
                string[] diassal = cantdias.Split('.',',');
                int dias = Int32.Parse(diassal[0]);
                int sbase = Int32.Parse(basesal[0]);

                int resultadoformula = sbase / 30 * dias;
                int rangomenor = resultadoformula - 100;
                int rangomayor = resultadoformula + 100;

                //comparacion

                string kactus = string.Empty;
                string sentencia3 = "select val_acum from nm_preno where cod_empl=" + "'" + id + "'" + " and cod_conc in (" + concepto + ")";
                DataTable resultado3 = FuncionesVitales.SelectData(sentencia3, user, motor);
                if (resultado3.Rows.Count > 0)
                {
                    foreach (DataRow rw in resultado3.Rows)
                    {
                        kactus = rw["val_acum"].ToString();

                    }
                }
                string[] kactusvt = kactus.Split('.',',');
                int kactusr = Int32.Parse(kactusvt[0]);

                string resultadoformulas = resultadoformula.ToString();

                if (kactusr >= rangomenor && kactusr <= rangomayor)
                {

                    return "Validacion correcta, los valores concuerdan" + kactusvt[0] + "  " + resultadoformulas;
                }
                else
                {
                    return "Validacion de salario erronea, se esperaba" + resultadoformulas + "se obtuvo " + kactusvt[0];

                }
        }

        static public string ValidacionesIncapEmpresa(int bandera, string motor, string user, string id, string empresa, string concepto, string fechaacumu)
        {

            //salario base
            string salariobase = string.Empty;
            string sentencia = "select sue_basi from nm_contr where cod_empl="+"'"+id+"'"+"and cod_empr="+"'"+empresa+"'";
            DataTable resultado = FuncionesVitales.SelectData(sentencia, user, motor);
            if (resultado.Rows.Count > 0)
            {
                foreach (DataRow rw in resultado.Rows)
                {
                    salariobase = rw["sue_basi"].ToString();

                }
            }
            var dateTime = DateTime.ParseExact(fechaacumu, "dd/MM/yyyy", null);
            string fecha = dateTime.ToString("yyyyMMdd");
            //cantidad dias
            string cantdias = string.Empty;
            string sentencia2 = "select can_acum from nm_preno where cod_empl=" + "'" + id + "'" + "and cod_conc in ("+concepto+")"+"AND fec_acum ="+"'"+fecha+"'";
            DataTable resultado2 = FuncionesVitales.SelectData(sentencia2, user, motor);
            if (resultado2.Rows.Count > 0)
            {
                foreach (DataRow rw in resultado2.Rows)
                {
                    cantdias = rw["can_acum"].ToString();

                }
            }
            //formulas
            string[] basesal = salariobase.Split('.',',');
            string[] diassal = cantdias.Split('.', ',');
            int dias = Int32.Parse(diassal[0]);
            int sbase = Int32.Parse(basesal[0]);

            int resultadoformula = sbase / 30 * dias;
            int rangomenor = resultadoformula - 100;
            int rangomayor = resultadoformula + 100;
            

            //comparacion

            string kactus = string.Empty;
            string sentencia3 = "select val_acum from nm_preno where cod_empl="+"'"+id+"'"+ "and cod_conc in ("+concepto+")"+ "AND fec_acum ="+"'"+fecha+"'";
            DataTable resultado3 = FuncionesVitales.SelectData(sentencia3, user, motor);
            if (resultado3.Rows.Count > 0)
            {
                foreach (DataRow rw in resultado3.Rows)
                {
                    kactus = rw["val_acum"].ToString();

                }
            }
            string[] kactusvt = kactus.Split('.',',');
            int kactusr = Int32.Parse(kactusvt[0]);
            string resultadoformulas = resultadoformula.ToString();

            if (kactusr >= rangomenor && kactusr <= rangomayor)
            {
                
                return "Validacion correcta, los valores concuerdan" + kactusvt[0] + "  " + resultadoformulas;
            }
            else
            {
                return "Validacion de 2 dias incapacidad erronea, se esperaba" + resultadoformulas + "se obtuvo " + kactusvt[0];

            }

        }

        static public string ValidacionesIncapacidad(int bandera, string motor, string user, string id, string concepto, string fechaincap, string fechaacumu)
        {

            //base liquidacion
            string baseliquidacion = string.Empty;
            string sentencia = "select bas_liqu from nm_incap where cod_empl="+"'"+id+"'"+" and fec_desd="+"'"+fechaincap+"'";
            DataTable resultado = FuncionesVitales.SelectData(sentencia, user, motor);
            if (resultado.Rows.Count > 0)
            {
                foreach (DataRow rw in resultado.Rows)
                {
                    baseliquidacion = rw["bas_liqu"].ToString();

                }
            }

            //formulas
            string[] basesal = baseliquidacion.Split('.',',');
            int sbase = Int32.Parse(basesal[0]);

            double resultadoformula = (sbase / 30 * 2) * (66.667 / 100);
            string resultadoformulas = resultadoformula.ToString();
            string[] formula = resultadoformulas.Split('.',',');
            int valor = Int32.Parse(formula[0]);

            int rangomenor = valor - 100;
            int rangomayor = valor + 100;

            //comparacion
            var dateTime = DateTime.ParseExact(fechaacumu, "dd/MM/yyyy", null);
            string fecha = dateTime.ToString("yyyyMMdd");
            string kactus = string.Empty;

            string sentencia3 = "select val_acum from nm_preno where cod_empl="+"'"+id+"'"+ "and cod_conc in ("+concepto+") AND fec_acum ="+"'"+fecha+"'";
            DataTable resultado3 = FuncionesVitales.SelectData(sentencia3, user, motor);
            if (resultado3.Rows.Count > 0)
            {
                foreach (DataRow rw in resultado3.Rows)
                {
                    kactus = rw["val_acum"].ToString();

                }
            }
            string[] kactusvt = kactus.Split('.',',');
            int kactusr = Int32.Parse(kactusvt[0]);

            if (kactusr >= rangomenor && kactusr <= rangomayor)
            {
                return "Validacion correcta, los valores concuerdan" + kactusvt[0] +"  "+formula[0];
            }
            else
            {
                return "Validacion de incapacidad erronea, se esperaba" + formula[0] + "se obtuvo "+kactusvt[0];

            }

        }
        static public string ValidacionesSalud(int bandera, string motor, string user, string id, string concepto, string concepto1, string concepto2, string fechaacumu)
        {
            string tipsala = string.Empty;
            string sentencia3 = "select tip_sala from nm_contr where cod_empl=" + "'" + id + "'";
            DataTable resultado3 = FuncionesVitales.SelectData(sentencia3, user, motor);
            if (resultado3.Rows.Count > 0)
            {
                foreach (DataRow rw in resultado3.Rows)
                {
                    tipsala = rw["tip_sala"].ToString();

                }
            }

           


            string valor = string.Empty;
            var dateTime = DateTime.ParseExact(fechaacumu, "dd/MM/yyyy", null);
            string fecha = dateTime.ToString("yyyyMMdd");
            string sentencia = "select val_acum from nm_preno where cod_empl=" + "'" + id + "'" + "and cod_conc in (" + concepto + ") AND fec_acum =" + "'" + fecha + "'";
            DataTable resultado = FuncionesVitales.SelectData(sentencia, user, motor);
            if (resultado.Rows.Count > 0)
            {
                foreach (DataRow rw in resultado.Rows)
                {
                    valor = rw["val_acum"].ToString();

                }
            }
            string valor1 = string.Empty;
            string sentencia1 = "select val_acum from nm_preno where cod_empl=" + "'" + id + "'" + "and cod_conc in (" + concepto1 + ") AND fec_acum =" + "'" + fecha + "'";
            DataTable resultado1 = FuncionesVitales.SelectData(sentencia1, user, motor);
            if (resultado1.Rows.Count > 0)
            {
                foreach (DataRow rw in resultado1.Rows)
                {
                    valor1 = rw["val_acum"].ToString();

                }
            }
            string valor2 = string.Empty;
            string sentencia2 = "select val_acum from nm_preno where cod_empl=" + "'" + id + "'" + "and cod_conc in (" + concepto2 + ") AND fec_acum =" + "'" + fecha + "'";
            DataTable resultado2 = FuncionesVitales.SelectData(sentencia2, user, motor);
            if (resultado2.Rows.Count > 0)
            {
                foreach (DataRow rw in resultado2.Rows)
                {
                    valor2 = rw["val_acum"].ToString();

                }
            }



            string[] valorr = valor.Split('.', ',');
            int valori = Int32.Parse(valorr[0]);

            string[] valorr1 = valor1.Split('.', ',');
            int valori1 = Int32.Parse(valorr1[0]);

            string[] valorr2 = valor2.Split('.', ',');
            int valori2 = Int32.Parse(valorr2[0]);

            double sumaconceptos = valori + valori1 + valori2;

            string salud = string.Empty;
            string sentenciasa = "select val_acum from nm_preno where cod_empl=" + "'" + id + "'" + "and cod_conc in ('5050') AND fec_acum =" + "'" + fecha + "'";
            DataTable resultadosa = FuncionesVitales.SelectData(sentenciasa, user, motor);
            if (resultadosa.Rows.Count > 0)
            {
                foreach (DataRow rw in resultadosa.Rows)
                {
                    salud = rw["val_acum"].ToString();

                }
            }
            string[] saluds = salud.Split('.', ',');
            int salidi = Int32.Parse(saluds[0]);

            int resultadoi = 0;

            if (tipsala.Contains("F"))
            {
               
                double resultadoformula = sumaconceptos * 0.04;
                double resultadoro=Math.Ceiling(resultadoformula);
                resultadoi = (int)resultadoro;



            }
            else if (tipsala.Contains("I"))
            {
                double resultadoformula = sumaconceptos * 0.7;
                double resultadorformula = Math.Ceiling(resultadoformula);
                double resultadoformuladef = resultadorformula * 0.04;

                double resultadodo = Math.Ceiling(resultadoformuladef);

                resultadoi = (int)resultadodo;

            }

            if (salidi >= resultadoi - 100 && salidi <= resultadoi + 100)
            {
                return "Validacion correcta, los valores se encuentran dentro del rango correcto   " + " Kactus: " + salidi.ToString() + " Formula: " + resultadoi.ToString();
            }
            else
            {
                return "Validacion erronea " + "Kactus: " + salidi.ToString() + "Formula: " + resultadoi.ToString();

            }

        }

        static public string ValidacionesPension(int bandera, string motor, string user, string id, string concepto, string concepto1, string concepto2, string fechaacumu)
        {
            string tipsala = string.Empty;
            string sentencia3 = "select tip_sala from nm_contr where cod_empl=" + "'" + id + "'";
            DataTable resultado3 = FuncionesVitales.SelectData(sentencia3, user, motor);
            if (resultado3.Rows.Count > 0)
            {
                foreach (DataRow rw in resultado3.Rows)
                {
                    tipsala = rw["tip_sala"].ToString();

                }
            }




            string valor = string.Empty;
            var dateTime = DateTime.ParseExact(fechaacumu, "dd/MM/yyyy", null);
            string fecha = dateTime.ToString("yyyyMMdd");
            string sentencia = "select val_acum from nm_preno where cod_empl=" + "'" + id + "'" + "and cod_conc in (" + concepto + ") AND fec_acum =" + "'" + fecha + "'";
            DataTable resultado = FuncionesVitales.SelectData(sentencia, user, motor);
            if (resultado.Rows.Count > 0)
            {
                foreach (DataRow rw in resultado.Rows)
                {
                    valor = rw["val_acum"].ToString();

                }
            }
            string valor1 = string.Empty;
            string sentencia1 = "select val_acum from nm_preno where cod_empl=" + "'" + id + "'" + "and cod_conc in (" + concepto1 + ") AND fec_acum =" + "'" + fecha + "'";
            DataTable resultado1 = FuncionesVitales.SelectData(sentencia1, user, motor);
            if (resultado1.Rows.Count > 0)
            {
                foreach (DataRow rw in resultado1.Rows)
                {
                    valor1 = rw["val_acum"].ToString();

                }
            }
            string valor2 = string.Empty;
            string sentencia2 = "select val_acum from nm_preno where cod_empl=" + "'" + id + "'" + "and cod_conc in (" + concepto2 + ") AND fec_acum =" + "'" + fecha + "'";
            DataTable resultado2 = FuncionesVitales.SelectData(sentencia2, user, motor);
            if (resultado2.Rows.Count > 0)
            {
                foreach (DataRow rw in resultado2.Rows)
                {
                    valor2 = rw["val_acum"].ToString();

                }
            }



            string[] valorr = valor.Split('.', ',');
            int valori = Int32.Parse(valorr[0]);

            string[] valorr1 = valor1.Split('.', ',');
            int valori1 = Int32.Parse(valorr1[0]);

            string[] valorr2 = valor2.Split('.', ',');
            int valori2 = Int32.Parse(valorr2[0]);

            double sumaconceptos = valori + valori1 + valori2;

            string pension = string.Empty;
            string sentenciasa = "select val_acum from nm_preno where cod_empl=" + "'" + id + "'" + "and cod_conc in ('5052') AND fec_acum =" + "'" + fecha + "'";
            DataTable resultadosa = FuncionesVitales.SelectData(sentenciasa, user, motor);
            if (resultadosa.Rows.Count > 0)
            {
                foreach (DataRow rw in resultadosa.Rows)
                {
                    pension = rw["val_acum"].ToString();

                }
            }
            string[] pensions = pension.Split('.', ',');
            int pensioni = Int32.Parse(pensions[0]);

            int resultadoi = 0;

            if (tipsala.Contains("F"))
            {

                double resultadoformula = sumaconceptos * 0.04;
                double resultadoro = Math.Ceiling(resultadoformula);
                resultadoi = (int)resultadoro;



            }
            else if (tipsala.Contains("I"))
            {
                double resultadoformula = sumaconceptos * 0.7;
                double resultadorformula = Math.Ceiling(resultadoformula);
                double resultadoformuladef = resultadorformula * 0.04;

                double resultadodo = Math.Ceiling(resultadoformuladef);

                resultadoi = (int)resultadodo;

            }

            if (pensioni >= resultadoi - 100 && pensioni <= resultadoi + 100)
            {
                return "Validacion correcta, los valores se encuentran dentro del rango correcto   " + " Kactus: " + pensioni.ToString() + " Formula: " + resultadoi.ToString();
            }
            else
            {
                return "Validacion erronea " + "Kactus: " + pensioni.ToString() + "Formula: " + resultadoi.ToString();

            }

        }
        static public string CrearDocumentoWordDinamico(string CodPrograma, string modulo, string database)
        {
            string fecha = Hora();
            string[] Database = database.Split('_');
            if (Database.Length > 1)
            {
                database = Database[Database.Length - 1];
            }
            string Path = string.Format(@"C:\Reportes\ReportesFinales{0}\Reportes_{1}\", modulo, database);
            if (!Directory.Exists(Path))
            {
                Directory.CreateDirectory(Path);
            }
            string file = string.Format(Path + "Reporte{0}_{1}.docx", CodPrograma, Database[0] + "_" + fecha);
            using (WordprocessingDocument wordDocument =
            WordprocessingDocument.Create(file, type: WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Prueba de ejecución realizada: " + fecha));
            }
            return file;
        }

        static public void Eliminar(string user, string tabla)
        {
            string ConnectionString = ConfigurationManager.ConnectionStrings[user].ConnectionString;
            SqlConnection conexion = new SqlConnection(ConnectionString);
            conexion.Open();
            SqlCommand cmd = new SqlCommand(string.Format("DELETE FROM {0} WHERE ACT_USUA='{1}'", tabla, user), conexion);
            int filasafectadas = cmd.ExecuteNonQuery();
            conexion.Close();
        }

        public static void Screenshot(string maestro, bool bandera, string file, WindowsDriver<WindowsElement> session)
        {
            //Image MyImage = null;
            //string UrlImage = @"C:\Reportes\" + maestro + ".bmp";
            //MyImage = ((ITakesScreenshot)session).GetScreenshot();
            Screenshot image = ((ITakesScreenshot)session).GetScreenshot();
            string name = maestro;
            string path = @"C:\Reportes\" + maestro + ".bmp";
            Image imgSource;
            image.SaveAsFile(string.Format("C:\\Reportes\\{0}.bmp", name), ScreenshotImageFormat.Bmp);
            //image.Save(path, System.Drawing.Imaging.ImageFormat.Bmp);
            InsertAPicture(file, path, maestro, bandera);
        }

        public static void InsertAPicture(string document, string UrlImage, string maestro, bool bandera)
        {
            using (WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(document, true))
            {
                MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

                using (FileStream stream = new FileStream(UrlImage, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart), maestro, bandera, UrlImage);
            }
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId, string maestro, bool bandera, string UrlImage)
        {
            int iWidth = 0;
            int iHeight = 0;
            using (System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(UrlImage))
            {
                iWidth = bmp.Width;
                iHeight = bmp.Height;
            }
            iWidth = (int)Math.Round((decimal)iWidth * 4000);
            iHeight = (int)Math.Round((decimal)iHeight * 4000);

            var element =

                new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = iWidth, Cy = iHeight },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.bmp"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.None
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 10000L, Y = 10000L },
                                             new A.Extents() { Cx = iWidth, Cy = iHeight }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            if (bandera)
            {
                wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(new Text(maestro))));
            }
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));

        }

        static public List<string> Celda_E(string pathexcel)
        {
            List<string> val2 = new List<string>();
            List<Regex> Reg = new List<Regex>();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(pathexcel);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            val2.Clear(); Reg.Clear();
            string val = "";
            bool bandera = true;
            Reg.Add(new Regex(@"(?i)[_]"));
            Reg.Add(new Regex(@"\bFRM"));
            Reg.Add(new Regex(@"\bDTFNMB"));
            string[] res = new string[Reg.Count()];
            int i = 1;
            for (int j = 1; j <= colCount; j++)
            {
                val = (xlRange.Cells[i, j]).Text;

                for (int l = 0; l < Reg.Count(); l++)
                {
                    res[l] = Reg[l].IsMatch(val) ? "No Está de Forma Correcta" : "Está de Forma Correcta";
                    if (res[l] == "No Está de Forma Correcta")
                    {
                        val2.Add(string.Format("Se encontro la expresion '{0}' en el campo {1} en el reporte " +
                                                "Exportacion del Excel", Reg[l].Match(val).Value.ToString(), val));
                    }
                    if (val.ToLower() == "Nombre Empresa".ToLower() || val.ToLower() == "COD_EMPR".ToLower())
                    {
                        string nom_empraux = (xlRange.Cells[i + 1, j]).Text;
                        //if (nom_empr != nom_empraux)
                        //{
                        //    if (bandera)
                        //    {
                        //        val2.Add(string.Format("El nombre de la empresa en el reporte no coincide. " +
                        //        "El nombre de la empresa es '{0}' y el encontrado en el reporte es '{1}'", nom_empr, nom_empraux));
                        //        bandera = false;
                        //    }
                        //}
                    }
                    res[l] = "";
                }
            }
            xlApp.Quit();
            return val2;
        }

       
        static public List<string> Celda_Excel(string pathexcel, string nom_empr)
        {
            List<string> val2 = new List<string>();
            List<Regex> Reg = new List<Regex>();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(pathexcel);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            val2.Clear(); Reg.Clear();
            string val = "";
            bool bandera = true;
            Reg.Add(new Regex(@"(?i)[_]"));
            Reg.Add(new Regex(@"\bFRM"));
            Reg.Add(new Regex(@"\bDtfNmb"));
            string[] res = new string[Reg.Count()];
            int i = 1;
            for (int j = 1; j <= colCount; j++)
            {
                val = (xlRange.Cells[i, j]).Text;

                for (int l = 0; l < Reg.Count(); l++)
                {
                    res[l] = Reg[l].IsMatch(val) ? "No Está de Forma Correcta" : "Está de Forma Correcta";
                    if (res[l] == "No Está de Forma Correcta")
                    {
                        val2.Add(string.Format("Se encontro la expresion '{0}' en el campo {1} en el reporte " +
                                                "Exportacion del Excel", Reg[l].Match(val).Value.ToString(), val));
                    }
                    if (val.ToLower() == "Nombre Empresa".ToLower() || val.ToLower() == "NOM_EMPR".ToLower())
                    {
                        string nom_empraux = (xlRange.Cells[i + 1, j]).Text;
                        if (nom_empr.Replace(" ", string.Empty) != nom_empraux.Replace(" ", string.Empty))
                        {
                            if (bandera)
                            {
                                val2.Add(string.Format("El nombre de la empresa en el reporte no coincide. " +
                                "El nombre de la empresa esperado '{0}' y el encontrado en el reporte Exportar a Excel es es '{1}'", nom_empr, nom_empraux));
                                bandera = false;
                            }
                        }
                    }
                    res[l] = "";
                }
            }
            xlApp.Quit();
            return val2;
        }

        public List<string> Textopdf(string path, string codprog, string User)
        {
            List<Regex> ExpReg = new List<Regex>();
            List<string> Errores = new List<string>();
            Errores.Clear();
            ExpReg.Add(new Regex(@"\bKactuS:\s+" + codprog, RegexOptions.IgnoreCase));
            ExpReg.Add(new Regex(@"\bUsuario:\s+" + User, RegexOptions.IgnoreCase));
            ExpReg.Add(new Regex(@"\bKactus\s+by\s+Ophelia", RegexOptions.IgnoreCase));
            ExpReg.Add(new Regex(@"\bDigital\s+Ware", RegexOptions.IgnoreCase));
            string[] busqueda = new string[ExpReg.Count()];
            busqueda[0] = "KactuS: " + codprog;
            busqueda[1] = "Usuario: " + User;
            busqueda[2] = "Kactus by Ophelia";
            busqueda[3] = "Digital Ware";
            string[] res = new string[ExpReg.Count()];
            string texto;
            path = path + ".pdf";

            try
            {
                if (File.Exists(path))
                {
                    PdfReader r = new PdfReader(path);
                    int intPageNum = r.NumberOfPages;
                    for (int i = 1; i <= intPageNum; i++)
                    {
                        texto = PdfTextExtractor.GetTextFromPage(r, i, new LocationTextExtractionStrategy());

                        for (int l = 0; l < ExpReg.Count(); l++)
                        {
                            res[l] = ExpReg[l].Match(texto).Value.ToString();
                            if (res[l] == "")
                            {
                                Errores.Add(string.Format("La expresion '{0}' NO se encontró en el Reporte PDF.", busqueda[l]));
                            }
                        }
                    }
                }
                else
                {
                    Errores.Add("El Reporte PDF no se Generó, Puede que haya Cambiado la Ubicación o se Generó un Error al Guardarlo");
                }
            }
            catch (Exception e)
            {
                Errores.Add(e.ToString());
            }
            return Errores;
        }

        public static void ReporteErrores(List<string> Error, string CodPrograma, string database, string modulo)
        {
            string Ruta = string.Format(@"C:\Reportes\ReportesFinales{0}\Reportes_{1}\Errores\", modulo, database);
            if (!Directory.Exists(Ruta))
            {
                Directory.CreateDirectory(Ruta);
            }
            int b = 0;
            StreamWriter sw = new StreamWriter(string.Format(Ruta + "Reporte_Error{0}_{1}_{2}.txt", CodPrograma,database, Hora()));
            for (int i = 0; i < (Error.ToArray()).Length; i++)
            {
                b++;
                sw.WriteLine(b + "-" + Error[i]);
            }
            sw.Close();
        }

        public static List<string> ValidaAdjuntos(string CodEmpleado, string Database, string User, string tabla, string CodEmpresa, List<string> Documents, string Almacenamiento, string RutaCP, string RutaFTP, int NumRows = 0, string delete = null, string Sentencia = null)
        {
            if (Sentencia == null)
            {
                Sentencia = string.Format("SELECT ADJ_GUID FROM {0} WHERE COD_EMPL={1} AND COD_EMPR={2}", tabla, CodEmpleado, CodEmpresa);
            }
            string ConnectionString = ConfigurationManager.ConnectionStrings[User].ConnectionString;
            List<string> Errores = new List<string>();
            List<string> AdjCont = new List<string>();
            List<string> AdjGuid = new List<string>();
            List<string> AdjNomb = new List<string>();
            DataSet DataSql = new DataSet();
            if (Database == "SQL")
            {
                SqlDataReader Leer;
                DataTable Out = new DataTable(); DataTable Out1 = new DataTable();
                DataTable Out2 = new DataTable(); DataTable Out3 = new DataTable();
                SqlCommand comando = new SqlCommand();
                SqlConnection ConexionDB = new SqlConnection(ConnectionString);
                ConexionDB.Open();
                comando.Connection = ConexionDB;
                comando.CommandText = Sentencia;
                Leer = comando.ExecuteReader();
                Out.Load(Leer);
                string adj_guid = "";
                foreach (DataRow Aux in Out.Rows)
                {
                    adj_guid = Aux["ADJ_GUID"].ToString();
                }
                if ((adj_guid == null || adj_guid == "") && NumRows == 0)
                {
                    if (delete == null)
                    {
                        Errores.Add(string.Format("El ADJ_GUID asociado al codigo de empleado {0} se encuentra vacio", CodEmpleado));
                    }
                }

                comando.CommandText = string.Format("SELECT RAD_CONT FROM GN_RADJU WHERE ADJ_GUID='{0}'", adj_guid);
                Leer = comando.ExecuteReader();
                Out1.Load(Leer);
                string RAD_CONT = "";
                foreach (DataRow Aux in Out1.Rows)
                {
                    RAD_CONT = Aux["RAD_CONT"].ToString();
                }
                if ((RAD_CONT == null || RAD_CONT == "") && (NumRows == 0))
                {
                    if (delete == null)
                    {
                        Errores.Add(string.Format("El código RAD_CONT asociado al ADJ_GUID = '{0}' se encuentra vacio", adj_guid));
                    }
                }

                comando.CommandText = string.Format("SELECT ADJ_CONT, ADJ_GUID, ADJ_NOMB FROM GN_ADJUN WHERE RAD_CONT='{0}'", RAD_CONT);
                Leer = comando.ExecuteReader();
                Out2.Load(Leer);
                int i = 0;
                string ADJ_NOMB = "";
                string doc;
                foreach (DataRow Aux in Out2.Rows)
                {
                    AdjCont.Add(Aux["ADJ_CONT"].ToString());
                    AdjGuid.Add(Aux["ADJ_GUID"].ToString());
                    AdjNomb.Add(Aux["ADJ_NOMB"].ToString());
                    string[] split = Documents[i].Split('\\');
                    doc = split[split.Length - 1];
                    ADJ_NOMB = Aux["ADJ_NOMB"].ToString();
                    if ((ADJ_NOMB != doc) && (NumRows == 0))
                    {
                        if (delete == null)
                        {
                            Errores.Add(string.Format("El nombre del documento esperado es: {0} y el encontrado es: {1} ", doc, ADJ_NOMB));
                        }
                    }
                    i++;
                }

                switch (Almacenamiento.ToUpper())
                {
                    case "B":
                        comando.CommandText = string.Format("SELECT ADJ_CONT FROM GN_ADJUN WHERE RAD_CONT='{0}'", RAD_CONT);
                        Leer = comando.ExecuteReader();
                        Out3.Load(Leer);
                        int c = 0;
                        string ADJ_CONT = "";
                        foreach (DataRow Aux in Out3.Rows)
                        {
                            string adjcont = AdjCont[c];
                            ADJ_CONT = Aux["ADJ_CONT"].ToString();
                            if (ADJ_CONT != adjcont)
                            {
                                if (delete == null)
                                {
                                    Errores.Add(string.Format("El ADJ_CONT encontrado es: {0} y es esperado es: {1}", ADJ_CONT, AdjCont[c]));
                                }
                            }
                            c++;
                        }
                        if (NumRows != 0)
                        {
                            if (AdjGuid.Count != NumRows && delete == null)
                            {
                                Errores.Add(string.Format("No se eliminaron correctamente los archivos, se esperaban {0} registros y se encontraron {1}", NumRows, AdjGuid.Count));
                            }
                        }
                        break;
                    case "C":
                        string rutaArcCP = "";
                        string rutaFinalCP = RutaCP + adj_guid + "\\";
                        int j = 0;
                        foreach (var Arc in AdjGuid)
                        {
                            rutaArcCP = rutaFinalCP + Arc;
                            if (!File.Exists(rutaArcCP))
                            {
                                if (delete == null)
                                {
                                    Errores.Add(string.Format("No se encontró en la Carpeta Publica el archivo {0}", AdjNomb[j]));
                                }
                            }
                            j++;
                        }
                        if (NumRows != 0)
                        {
                            if (AdjGuid.Count != NumRows && delete == null)
                            {
                                Errores.Add(string.Format("No se eliminaron correctamente los archivos, se esperaban {0} registros y se encontraron {1}", NumRows, AdjGuid.Count));
                            }
                        }
                        break;
                    case "F":
                        string rutaArcF = "";
                        string rutaFinalFTP = RutaFTP + adj_guid + "/";
                        j = 0;
                        foreach (var arch in AdjGuid)
                        {
                            rutaArcF = rutaFinalFTP + arch;
                            var request = (FtpWebRequest)WebRequest.Create(rutaArcF);
                            request.Credentials = new NetworkCredential("KactusSCM", "TF$Kactus*");
                            request.Method = WebRequestMethods.Ftp.GetFileSize;
                            try
                            {
                                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                            }
                            catch (WebException ex)
                            {
                                FtpWebResponse response = (FtpWebResponse)ex.Response;
                                if (delete == null && NumRows == 0)
                                {
                                    Errores.Add(ex + "No fue encontrado el archivo en el FTP: " + AdjNomb[j]);
                                }
                            }
                            j++;
                        }
                        if (NumRows != 0)
                        {
                            if (AdjGuid.Count != NumRows && delete == null)
                            {
                                Errores.Add(string.Format("No se eliminaron correctamente los archivos, se esperaban {0} registros y se encontraron {1}", NumRows, AdjGuid.Count));
                            }
                        }
                        break;
                    default:
                        Errores.Add("El parametro de almacenamiento ingresado no esta contemplado en las posibilidades de almacenamiento de adjuntos" + Almacenamiento);
                        break;
                }
                if (delete != null)
                {
                    string sentencia;
                    if (tabla.ToUpper() == "ED_FOMEI" || tabla.ToUpper() == "ED_DFOME")
                    {
                        sentencia = $"DELETE FROM GN_ADJUN WHERE RAD_CONT = {RAD_CONT}{Environment.NewLine}" +
                                    $"DELETE FROM GN_ADJFI WHERE RAD_CONT = {RAD_CONT}";
                    }
                    else
                    {
                        sentencia = $"DELETE FROM GN_ADJFI WHERE RAD_CONT={RAD_CONT}{Environment.NewLine}" +
                                       $"DELETE FROM GN_ADJUN WHERE RAD_CONT={RAD_CONT}{Environment.NewLine}" +
                                       $"DELETE FROM GN_RADJU WHERE ADJ_GUID='{adj_guid}'{Environment.NewLine}" +
                                       $"DELETE FROM NM_RENOV WHERE COD_EMPL={CodEmpleado} AND COD_EMPR={CodEmpresa}";
                    }


                    SqlConnection cnn = new SqlConnection(ConnectionString);
                    cnn.Open();
                    SqlCommand command = cnn.CreateCommand();
                    SqlTransaction transaction;
                    transaction = cnn.BeginTransaction();
                    command.Connection = cnn;
                    command.Transaction = transaction;
                    try
                    {
                        command.CommandText = sentencia;
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
                ConexionDB.Close();
            }
            else if (Database == "ORA")
            {
                OracleDataReader Leer;
                DataTable Out = new DataTable(); DataTable Out1 = new DataTable();
                DataTable Out2 = new DataTable(); DataTable Out3 = new DataTable();
                OracleCommand comando = new OracleCommand();
                OracleConnection ConexionDB = new OracleConnection(ConnectionString);
                ConexionDB.Open();
                comando.Connection = ConexionDB;
                comando.CommandText = Sentencia;
                Leer = comando.ExecuteReader();
                Out.Load(Leer);
                string adj_guid = "";
                foreach (DataRow Aux in Out.Rows)
                {
                    adj_guid = Aux["ADJ_GUID"].ToString();
                }
                if ((adj_guid == null || adj_guid == "") && NumRows == 0)
                {
                    if (delete == null)
                    {
                        Errores.Add(string.Format("El ADJ_GUID asociado al codigo de empleado {0} se encuentra vacio", CodEmpleado));
                    }
                }

                comando.CommandText = string.Format("SELECT RAD_CONT FROM GN_RADJU WHERE ADJ_GUID='{0}'", adj_guid);
                Leer = comando.ExecuteReader();
                Out1.Load(Leer);
                string RAD_CONT = "";
                foreach (DataRow Aux in Out1.Rows)
                {
                    RAD_CONT = Aux["RAD_CONT"].ToString();
                }
                if ((RAD_CONT == null || RAD_CONT == "") && (NumRows == 0))
                {
                    if (delete == null)
                    {
                        Errores.Add(string.Format("El código RAD_CONT asociado al ADJ_GUID = '{0}' se encuentra vacio", adj_guid));
                    }
                }

                comando.CommandText = string.Format("SELECT ADJ_CONT, ADJ_GUID, ADJ_NOMB FROM GN_ADJUN WHERE RAD_CONT='{0}'", RAD_CONT);
                Leer = comando.ExecuteReader();
                Out2.Load(Leer);
                int i = 0;
                string ADJ_NOMB = "";
                string doc;
                foreach (DataRow Aux in Out2.Rows)
                {
                    AdjCont.Add(Aux["ADJ_CONT"].ToString());
                    AdjGuid.Add(Aux["ADJ_GUID"].ToString());
                    AdjNomb.Add(Aux["ADJ_NOMB"].ToString());
                    string[] split = Documents[i].Split('\\');
                    doc = split[split.Length - 1];
                    ADJ_NOMB = Aux["ADJ_NOMB"].ToString();
                    if ((ADJ_NOMB != doc) && (NumRows == 0))
                    {
                        if (delete == null)
                        {
                            Errores.Add(string.Format("El nombre del documento esperado es: {0} y el encontrado es: {1} ", doc, ADJ_NOMB));
                        }
                    }
                    i++;
                }

                switch (Almacenamiento.ToUpper())
                {
                    case "B":
                        comando.CommandText = string.Format("SELECT ADJ_CONT FROM GN_ADJUN WHERE RAD_CONT='{0}'", RAD_CONT);
                        Leer = comando.ExecuteReader();
                        Out3.Load(Leer);
                        int c = 0;
                        string ADJ_CONT = "";
                        foreach (DataRow Aux in Out3.Rows)
                        {
                            string adjcont = AdjCont[c];
                            ADJ_CONT = Aux["ADJ_CONT"].ToString();
                            if (ADJ_CONT != adjcont)
                            {
                                if (delete == null)
                                {
                                    Errores.Add(string.Format("El ADJ_CONT encontrado es: {0} y es esperado es: {1}", ADJ_CONT, AdjCont[c]));
                                }
                            }
                            c++;
                        }
                        if (NumRows != 0)
                        {
                            if (AdjGuid.Count != NumRows && delete == null)
                            {
                                Errores.Add(string.Format("No se eliminaron correctamente los archivos, se esperaban {0} registros y se encontraron {1}", NumRows, AdjGuid.Count));
                            }
                        }
                        break;
                    case "C":
                        string rutaArcCP = "";
                        string rutaFinalCP = RutaCP + adj_guid + "\\";
                        int j = 0;
                        foreach (var Arc in AdjGuid)
                        {
                            rutaArcCP = rutaFinalCP + Arc;
                            if (!File.Exists(rutaArcCP))
                            {
                                if (delete == null)
                                {
                                    Errores.Add(string.Format("No se encontró en la Carpeta Publica el archivo {0}", AdjNomb[j]));
                                }
                            }
                            j++;
                        }
                        if (NumRows != 0)
                        {
                            if (AdjGuid.Count != NumRows && delete == null)
                            {
                                Errores.Add(string.Format("No se eliminaron correctamente los archivos, se esperaban {0} registros y se encontraron {1}", NumRows, AdjGuid.Count));
                            }
                        }
                        break;
                    case "F":
                        string rutaArcF = "";
                        string rutaFinalFTP = RutaFTP + adj_guid + "/";
                        j = 0;
                        foreach (var arch in AdjGuid)
                        {
                            rutaArcF = rutaFinalFTP + arch;
                            var request = (FtpWebRequest)WebRequest.Create(rutaArcF);
                            request.Credentials = new NetworkCredential("KactusSCM", "TF$Kactus*");
                            request.Method = WebRequestMethods.Ftp.GetFileSize;
                            try
                            {
                                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                            }
                            catch (WebException ex)
                            {
                                FtpWebResponse response = (FtpWebResponse)ex.Response;
                                if (delete == null && NumRows == 0)
                                {
                                    Errores.Add(ex + "No fue encontrado el archivo en el FTP: " + AdjNomb[j]);
                                }
                            }
                            j++;
                        }
                        if (NumRows != 0)
                        {
                            if (AdjGuid.Count != NumRows && delete == null)
                            {
                                Errores.Add(string.Format("No se eliminaron correctamente los archivos, se esperaban {0} registros y se encontraron {1}", NumRows, AdjGuid.Count));
                            }
                        }
                        break;
                    default:
                        Errores.Add("El parametro de almacenamiento ingresado no esta contemplado en las posibilidades de almacenamiento de adjuntos" + Almacenamiento);
                        break;
                }
                if (delete != null)
                {
                    List<string> Sentencias = new List<string>();
                    string sentencia;
                    if (tabla.ToUpper() == "ED_FOMEI" || tabla.ToUpper() == "ED_DFOME")
                    {
                        Sentencias.Add($"DELETE FROM GN_ADJUN WHERE RAD_CONT={RAD_CONT}");
                        Sentencias.Add($"DELETE FROM GN_ADJFI WHERE RAD_CONT={RAD_CONT}");
                    }
                    else
                    {
                        Sentencias.Add($"DELETE FROM GN_ADJFI WHERE RAD_CONT={RAD_CONT}");
                        Sentencias.Add($"DELETE FROM GN_ADJUN WHERE RAD_CONT={RAD_CONT}");
                        Sentencias.Add($"DELETE FROM GN_RADJU WHERE ADJ_GUID='{adj_guid}'");
                        Sentencias.Add($"DELETE FROM NM_RENOV WHERE COD_EMPL={CodEmpleado} AND COD_EMPR={CodEmpresa}");
                    }

                    foreach (string sen in Sentencias)
                    {
                        OracleConnection cnn = new OracleConnection(ConnectionString);
                        cnn.Open();
                        OracleCommand command = cnn.CreateCommand();
                        OracleTransaction transaction;
                        transaction = cnn.BeginTransaction();
                        command.Connection = cnn;
                        command.Transaction = transaction;
                        try
                        {
                            command.CommandText = sen;
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
                }
                ConexionDB.Close();
            }
            return Errores;
        }

        public static List<string> ValidaDescarga(List<string> Documents, string file)
        {
            List<string> Errores = new List<string>();
            foreach (string path in Documents)
            {
                if (!File.Exists(path))
                {
                    Errores.Add(string.Format("Ocurrió un error al descargar el archivo adjunto. File: {0}", path));
                }
                else
                {
                    Process.Start(path);
                    Thread.Sleep(6000);
                    //Screenshot("Archivo Descargado MTOM", true, file);
                    Thread.Sleep(1000);
                    LimpiarProcesos();
                    Thread.Sleep(1000);
                    File.Delete(path);
                }
            }
            return Errores;
        }

        ////public static List<string> CalculaDiaIngreso(DateTime fechaProceso, int totDias, string CodEmpl, string CodEmpr, string User, string db)
        //{
        //    LiquidacionSS ObjetoLiqss = new LiquidacionSS();
        //    List<string> Errores = new List<string>();
        //    int cantDiaMes = 30;
        //    string query = $"SELECT fec_ingr FROM nm_contr WHERE cod_empl={CodEmpl} and cod_empr={CodEmpr}";
        //    var resul = ObjetoLiqss.ConsultarSqlOra(query, User, db);
        //    DateTime fechaIngreso = new DateTime();
        //    foreach (DataRow fila in resul.Tables[0].Rows)
        //    {
        //        fechaIngreso = Convert.ToDateTime(fila["fec_ingr"]);
        //    }
        //    var mesProceso = fechaProceso.Month;
        //    var anoProceso = fechaProceso.Year;
        //    var diaIngreso = fechaIngreso.Day;
        //    var mesIngreso = fechaIngreso.Month;
        //    var anoIngreso = fechaIngreso.Year;

        //    if (cantDiaMes != totDias)
        //    {
        //        if (mesProceso == mesIngreso && anoProceso == anoIngreso)
        //        {
        //            int diasTrabajaMes = cantDiaMes - (diaIngreso - 1);
        //            if (diasTrabajaMes != totDias)
        //            {
        //                Errores.Add($"EMPLEADO {CodEmpl} La cantidad de días trabajados en el mes deberia ser {diasTrabajaMes} y la sumatoria encontrada es {totDias}, Fecha de proceso: {fechaProceso}, Fecha ingreso: {fechaIngreso}");
        //            }
        //        }
        //        else
        //        {
        //            Errores.Add($"EMPLEADO {CodEmpl} La cantidad de días trabajados no debería ser diferente a {cantDiaMes}, Fecha de proceso: {fechaProceso}, Fecha ingreso: {fechaIngreso}");
        //        }
        //    }
        //    return Errores;
        //}
    }
}

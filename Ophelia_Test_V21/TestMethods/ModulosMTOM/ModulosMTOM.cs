using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21;
using Ophelia_Test_V21.AFuncionesGlobales;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.ValidacionQueries;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesMTOM;

namespace Ophelia_Test_V21.TestMethods.ModulosMTOM
{
    /// <summary>
    /// Descripción resumida de ModulosMTOM
    /// </summary>
    [TestClass]
    public class ModulosMTOM:FuncionesVitales
    {

        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;

        public ModulosMTOM(){}


        [TestMethod]
        public void abrirPrograma()
        {
            List<string> errors = new List<string>();
            string motor = "SQL";
            AbrirPrograma a = new AbrirPrograma("KBiDatad", "UCalida1");
            //string motor = "ORA";
            //AbrirPrograma a = new AbrirPrograma("KNmBasld", "ODESAR");

            

            desktopSession = a.Start2(motor, "");
            Thread.Sleep(200000);

            //string file = CrearDocumentoWordDinamico("KBiEdfor", "PruebasMTOM", motor);
            //FuncionesGlobales.QBEQry(desktopSession, "0", "7", file);

            //newFunctions_1.lapiz(desktopSession);

            //AdjuntarArchivos.flechaRoja(desktopSession);

            //AdjuntarArchivos.controlVentana(desktopSession);

            //Thread.Sleep(2000);
        }
        
        [TestMethod]
        public void MarcoMTOM()
        {
            List<string> ErrorValidacion = new List<string>();
            List<string> ValidaSubida = new List<string>();
            List<string> ValidaDescarga = new List<string>();
            List<string> DocumentsLoad = new List<string>();
            List<string> DocumentsDownload = new List<string>();
            List<string> DocumentsDownloadAll = new List<string>();
            List<string> errorMessages = new List<string>();
            List<string> listaResultado;
            List<string> Celda = new List<string>();
            bool bandera = false;
            string enviroment = (Environment.MachineName);
            string[] auxtable = enviroment.Split('-');
            string TableOrder = "";
            if (auxtable.Length > 1)
            {
                TableOrder = (enviroment.Replace("-", "_")).ToUpper();
            }
            else
            {
                TableOrder = enviroment.ToUpper();
            }
            //TableOrder = "ktest1";
            DataSet OrderExecutionCase = SqlAdapter.SelectOrderExecution("T", TableOrder);
            int NumCasAgen = OrderExecutionCase.Tables[0].Rows.Count;
            if (NumCasAgen < 1)
            {
                errorMessages.Add("No hay casos en el agendamiento");
            }

            foreach (DataRow rowsi in OrderExecutionCase.Tables[0].Rows)
            {
                string plans = rowsi["plans"].ToString();
                string suite = rowsi["suite"].ToString();
                string CaseId = rowsi["CaseId"].ToString();
                string orders = rowsi["orders"].ToString();
                string states = rowsi["states"].ToString();
                string methodname = rowsi["methodname"].ToString();
                string CountDes = rowsi["CountDes"].ToString();

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosMTOM.ModulosMTOM.MarcoMTOM")
                {
                    DataSet sta = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                    string endstatus = null;
                    foreach (DataRow rowsta in sta.Tables[0].Rows)
                    {
                        endstatus = rowsta["states"].ToString();
                    }
                    if (endstatus == "True")
                    {

                        TFSData GetCasen = new TFSData(CaseId);
                        DataSet DataCase = GetCasen.GetParams();

                        foreach (DataRow rows in DataCase.Tables[0].Rows)
                        {
                            if (
                                //Variables Ophelia
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["Database"].ToString().Length != 0 && rows["Database"].ToString() != null &&
                                rows["Table"].ToString().Length != 0 && rows["Table"].ToString() != null &&
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["CodEmpresa"].ToString().Length != 0 && rows["CodEmpresa"].ToString() != null &&
                                rows["Almacenamiento"].ToString().Length != 0 && rows["Almacenamiento"].ToString() != null &&
                                rows["RutaFTP"].ToString().Length != 0 && rows["RutaFTP"].ToString() != null &&
                                rows["UserFTP"].ToString().Length != 0 && rows["UserFTP"].ToString() != null &&
                                rows["PassFTP"].ToString().Length != 0 && rows["PassFTP"].ToString() != null &&
                                rows["RutaCP"].ToString().Length != 0 && rows["RutaCP"].ToString() != null &&
                                rows["PesoArch"].ToString().Length != 0 && rows["PesoArch"].ToString() != null &&
                                rows["PortFTP"].ToString().Length != 0 && rows["PortFTP"].ToString() != null &&
                                rows["ServerFTP"].ToString().Length != 0 && rows["ServerFTP"].ToString() != null &&
                                rows["BanderaSmPe"].ToString().Length != 0 && rows["BanderaSmPe"].ToString() != null &&

                                //Variables Self-Service   
                                rows["PassSelf"].ToString().Length != 0 && rows["PassSelf"].ToString() != null &&
                                rows["URL"].ToString().Length != 0 && rows["URL"].ToString() != null &&
                                rows["url2"].ToString().Length != 0 && rows["url2"].ToString() != null &&
                                rows["CodDocu"].ToString().Length != 0 && rows["CodDocu"].ToString() != null &&
                                rows["NumDocu"].ToString().Length != 0 && rows["NumDocu"].ToString() != null &&
                                rows["Pais"].ToString().Length != 0 && rows["Pais"].ToString() != null &&
                                rows["Departamento"].ToString().Length != 0 && rows["Departamento"].ToString() != null &&
                                rows["Municipio"].ToString().Length != 0 && rows["Municipio"].ToString() != null &&
                                rows["Localidad"].ToString().Length != 0 && rows["Localidad"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["TipDocu1"].ToString().Length != 0 && rows["TipDocu1"].ToString() != null &&
                                rows["TipDocu2"].ToString().Length != 0 && rows["TipDocu2"].ToString() != null &&
                                rows["TipDoc"].ToString().Length != 0 && rows["TipDoc"].ToString() != null

                                )
                            {
                                string User = rows["User"].ToString();
                                string CodPrograma = rows["CodProgram"].ToString();
                                string Database = rows["Database"].ToString();
                                string Table = rows["Table"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string CodEmpresa = rows["CodEmpresa"].ToString();
                                string Almacenamiento = rows["Almacenamiento"].ToString();
                                string RutaFTP = rows["RutaFTP"].ToString();
                                string UserFTP = rows["UserFTP"].ToString();
                                string PassFTP = rows["PassFTP"].ToString();
                                string RutaCP = rows["RutaCP"].ToString();
                                string PesoArch = rows["PesoArch"].ToString();
                                string PortFTP = rows["PortFTP"].ToString();
                                string ServerFTP = rows["ServerFTP"].ToString();
                                string BanderaSmPe = rows["BanderaSmPe"].ToString();

                                string PassSelf = rows["PassSelf"].ToString();
                                string URL = rows["URL"].ToString();
                                string url2 = rows["url2"].ToString();
                                string CodDocu = rows["CodDocu"].ToString();
                                string NumDocu = rows["NumDocu"].ToString();
                                string Pais = rows["Pais"].ToString();
                                string Departamento = rows["Departamento"].ToString();
                                string Municipio = rows["Municipio"].ToString();
                                string Localidad = rows["Localidad"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string TipDocu1 = rows["TipDocu1"].ToString();
                                string TipDocu2 = rows["TipDocu2"].ToString();
                                string TipDoc = rows["TipDoc"].ToString();


                                //CREAR EL DOCUMENTO DINÄMICO
                                string file = CrearDocumentoWordDinamico(CodPrograma, "MTOM", Database);
                                try
                                {
                                    List<string> errors = new List<string>();

                                    ////////ACTUALIZA GNPAKAC//////
                                    //SqlAdapter.ModificaGNPAKAC(Almacenamiento, Database, User, CodEmpresa, RutaFTP, UserFTP, PassFTP, RutaCP, PesoArch, PortFTP, ServerFTP);

                                    ////INICIAR EL PROGRAMA OPHELIA////
                                    AbrirPrograma a = new AbrirPrograma(CodPrograma, User);
                                    desktopSession = a.Start2(Database, "");
                                    bool leeprograma = true;

                                    if (desktopSession != null)
                                    {
                                        string str = desktopSession.PageSource;
                                        if (str != null)
                                        {
                                            if (!str.Contains("TPanel"))
                                            {
                                                leeprograma = false;
                                            }
                                        }
                                        else
                                        {
                                            leeprograma = false;
                                        }
                                    }
                                    else
                                    {
                                        leeprograma = false;
                                    }

                                    if (leeprograma)
                                    {
                                        newFunctions_4.ScreenshotNuevo("Abrir Programa", true, file);
                                        
                                        string RutaFinalFTP = string.Format("ftp://{0}/{1}/", ServerFTP, RutaFTP);
                                        int NumRows = 0;

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                       
                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        
                                        //PASAR A LA VENTANA PRINCIPAL
                                        newFunctions_1.lapiz(desktopSession);

                                        //FUNCION ENCONRAR ICONO DE DOCUMENTOS MTOM
                                        AdjuntarArchivos.flechaRoja(desktopSession);

                                        ////ACCIONES DE CARGA MTOM/////

                                        //Adjuntar archivos
                                        AdjuntarArchivos.controlVentana(desktopSession, "1", QbeFilter, Database, User, Table, CodEmpresa, Almacenamiento, RutaCP, RutaFinalFTP, ValidaSubida, ErrorValidacion, ValidaDescarga, file, BanderaSmPe, URL, PassSelf, url2);
                                        //Descargar un archivo
                                        AdjuntarArchivos.controlVentana(desktopSession, "2", QbeFilter, Database, User, Table, CodEmpresa, Almacenamiento, RutaCP, RutaFinalFTP, ValidaSubida, ErrorValidacion, ValidaDescarga, file, BanderaSmPe, URL, PassSelf, url2);
                                        //Descargar Todo
                                        AdjuntarArchivos.controlVentana(desktopSession, "3", QbeFilter, Database, User, Table, CodEmpresa, Almacenamiento, RutaCP, RutaFinalFTP, ValidaSubida, ErrorValidacion, ValidaDescarga, file, BanderaSmPe, URL, PassSelf, url2);
                                        //Abrir un archivo
                                        AdjuntarArchivos.controlVentana(desktopSession, "4", QbeFilter, Database, User, Table, CodEmpresa, Almacenamiento, RutaCP, RutaFinalFTP, ValidaSubida, ErrorValidacion, ValidaDescarga, file, BanderaSmPe, URL, PassSelf, url2);
                                        //Elimnar un archivo
                                        AdjuntarArchivos.controlVentana(desktopSession, "5", QbeFilter, Database, User, Table, CodEmpresa, Almacenamiento, RutaCP, RutaFinalFTP, ValidaSubida, ErrorValidacion, ValidaDescarga, file, BanderaSmPe, URL, PassSelf, url2);
                                        //Eliminar Todos
                                        AdjuntarArchivos.controlVentana(desktopSession, "6", QbeFilter, Database, User, Table, CodEmpresa, Almacenamiento, RutaCP, RutaFinalFTP, ValidaSubida, ErrorValidacion, ValidaDescarga, file, BanderaSmPe, URL, PassSelf, url2);
                                        //Cerrar ventana MTOM
                                        AdjuntarArchivos.controlVentana(desktopSession, "7", QbeFilter, Database, User, Table, CodEmpresa, Almacenamiento, RutaCP, RutaFinalFTP, ValidaSubida, ErrorValidacion, ValidaDescarga, file, BanderaSmPe, URL, PassSelf, url2);

                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {CodPrograma}");
                                    }
                                    //CONTADOR DE ERRORES
                                    if (ErrorValidacion.Count > 0)
                                    {
                                        ReporteErrores(ErrorValidacion, CodPrograma, Database, Modulo);
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, ErrorValidacion);
                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}", Environment.NewLine, errorMessageString));
                                    }

                                    Thread.Sleep(4000);

                                    bandera = true;

                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();
                                        int StCount = Int32.Parse(SthCount);

                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    break;

                                }
                                catch (Exception e)
                                {
                                    Thread.Sleep(500);
                                    bandera = true;
                                    DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
                                    string SthCount = null;
                                    foreach (DataRow rowsta in Sth.Tables[0].Rows)
                                    {
                                        SthCount = rowsta["CountDes"].ToString();

                                        int StCount = Int32.Parse(SthCount);
                                        if (StCount > 0)
                                        {
                                            int NewCount = StCount - 1;
                                            DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
                                            if (NewCount == 0)
                                            {
                                                DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
                                                break;
                                            }
                                        }
                                    }
                                    Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    break;
                                }
                            }
							else
							{
								errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Al validar variables, no tiene registros o no existe");
							}
                        }
                        break;
                    }
                }
                else
                {
                    errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: El nombre de la automatizacion no corresponde");
                }
            }
            if (bandera == false)
            {
                if (errorMessages.Count > 0)
                {
                    var separator = string.Format("{0}{0}", Environment.NewLine);
                    string errorMessageString = string.Join(separator, errorMessages);

                    Assert.Inconclusive(string.Format("Las condiciones de ejecucion del caso son:{0}{1}",
                                 Environment.NewLine, errorMessageString));
                }
            }
        }

        [TestCleanup]
        public void Limpiar()
        {
            AbrirPrograma a = new AbrirPrograma();
            if (desktopSession != null)
            {
                desktopSession.Close();
                desktopSession.Dispose();
            }
            a.Stop();
        }

        private TestContext testContextInstance;

        /// <summary>
        ///Obtiene o establece el contexto de las pruebas que proporciona
        ///información y funcionalidad para la serie de pruebas actual.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Atributos de prueba adicionales
        //
        // Puede usar los siguientes atributos adicionales conforme escribe las pruebas:
        //
        // Use ClassInitialize para ejecutar el código antes de ejecutar la primera prueba en la clase
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup para ejecutar el código una vez ejecutadas todas las pruebas en una clase
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Usar TestInitialize para ejecutar el código antes de ejecutar cada prueba 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup para ejecutar el código una vez ejecutadas todas las pruebas
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        
    }
}

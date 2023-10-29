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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesoKGnAlert;

namespace Ophelia_Test_V21.TestMethods.ModulosGN
{
    /// <summary>
    /// Descripción resumida de ModulosBi_9
    /// </summary>
    [TestClass]
    public class ModulosProcesoKGnAlert : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();
        public ModulosProcesoKGnAlert()
        {
            //
            // TODO: Agregar aquí la lógica del constructor
            //
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

        //[TestMethod]
        //public void ModulosAlertasVencimientos()
        //{
        //    List<string> errorMessages = new List<string>();
        //    bool bandera = false;
        //    string enviroment = (Environment.MachineName);
        //    string[] auxtable = enviroment.Split('-');
        //    string TableOrder = "";
        //    if (auxtable.Length > 1)
        //    {
        //        TableOrder = (enviroment.Replace("-", "_")).ToUpper();
        //    }
        //    else
        //    {
        //        TableOrder = enviroment.ToUpper();
        //    }
        //    ////TableOrder = "ktest1";
        //    DataSet OrderExecutionCase = SqlAdapter.SelectOrderExecution("T", TableOrder);
        //    int NumCasAgen = OrderExecutionCase.Tables[0].Rows.Count;
        //    if (NumCasAgen < 1)
        //    {
        //        errorMessages.Add("No hay casos en el agendamiento");
        //    }

        //    foreach (DataRow rowsi in OrderExecutionCase.Tables[0].Rows)
        //    {
        //        string plans = rowsi["plans"].ToString();
        //        string suite = rowsi["suite"].ToString();
        //        string CaseId = rowsi["CaseId"].ToString();
        //        string orders = rowsi["orders"].ToString();
        //        string states = rowsi["states"].ToString();
        //        string methodname = rowsi["methodname"].ToString();
        //        string CountDes = rowsi["CountDes"].ToString();

        //        if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesoKGnAlert.ModulosProcesoKGnAlert.ModulosAlertasVencimientos")
        //        {
        //            DataSet sta = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
        //            string endstatus = null;
        //            foreach (DataRow rowsta in sta.Tables[0].Rows)
        //            {
        //                endstatus = rowsta["states"].ToString();
        //            }
        //            if (endstatus == "True")
        //            {

        //                TFSData GetCasen = new TFSData(CaseId);
        //                DataSet DataCase = GetCasen.GetParams();

        //                foreach (DataRow rows in DataCase.Tables[0].Rows)
        //                {

        //                    if (

        //                        //Datos Obligatorios
        //                        rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
        //                        rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
        //                        rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
        //                        rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&

        //                        //Datos Generales
        //                        rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
        //                        rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
        //                        rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
        //                        rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
        //                        rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
        //                        rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&

        //                        //Datos Crud Principal
        //                        rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
        //                        rows["Notifi"].ToString().Length != 0 && rows["Notifi"].ToString() != null &&
        //                        rows["Usua"].ToString().Length != 0 && rows["Usua"].ToString() != null &&

        //                        //Variables Validaciones Crud principal
        //                        rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
        //                        rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
        //                        rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null)

        //                    {
        //                        //Variables obligatorias
        //                        string user = rows["User"].ToString();
        //                        string motor = rows["Motor"].ToString();
        //                        string codProgram = rows["CodProgram"].ToString();
        //                        string nomProgram = rows["NomProgram"].ToString();

        //                        //Variables generales
        //                        string TipoQbe = rows["TipoQbe"].ToString();
        //                        string QbeFilter = rows["QbeFilter"].ToString();
        //                        string banExcel = rows["banExcel"].ToString();
        //                        string BanderaLupa = rows["BanderaLupa"].ToString();
        //                        string BanderaPreli = rows["BanderaPreli"].ToString();
        //                        string BanderaSum = rows["BanderaSum"].ToString();

        //                        //Variables validacion  crud Principal
        //                        string Tabla = rows["Tabla"].ToString();
        //                        string Campo = rows["Campo"].ToString();
        //                        string EditCampo = rows["EditCampo"].ToString();

        //                        //Variables crud principal
        //                        string Nombre = rows["Nombre"].ToString();
        //                        string Notifi = rows["Notifi"].ToString();
        //                        string Usua = rows["Usua"].ToString();

        //                        String campo = "1";

        //                        {

        //                            //LISTA CRUD PRINCIPAL
        //                            List<string> crudPrincipal = new List<string>() { Nombre, Notifi, Usua };

        //                            //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
        //                            string file = CrearDocumentoWordDinamico(codProgram, "ProcesooKGnAlert", motor);

        //                            try
        //                            {

        //                                List<string> ErrorValidacion = new List<string>();
        //                                List<string> errors = new List<string>();

        //                                AbrirPrograma a = new AbrirPrograma(codProgram, user);
        //                                desktopSession = a.Start2(motor, "");
        //                                ////Debugger.Launch();
        //                                bool leeprograma = true;

        //                                if (desktopSession != null)
        //                                {
        //                                    string str = desktopSession.PageSource;
        //                                    if (str != null)
        //                                    {
        //                                        if (!str.Contains("TPanel"))
        //                                        {
        //                                            leeprograma = false;
        //                                        }
        //                                    }
        //                                    else
        //                                    {
        //                                        leeprograma = false;
        //                                    }
        //                                }
        //                                else
        //                                {
        //                                    leeprograma = false;
        //                                }

        //                                if (leeprograma)
        //                                {
        //                                    //newFunctions_4.ScreenshotNuevo("Abrir Programa", true, file);
        //                                    ////VALIDACION CODIGO DEL PROGRAMA
        //                                    //errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
        //                                    //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


        //                                    ////VALIDACION NOMBRE DEL PROGRAMA
        //                                    //errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
        //                                    //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


        //                                    ////VERSION
        //                                    //FuncionesGlobales.GetVersion(desktopSession, file);
        //                                    //Thread.Sleep(1000);
        //                                    ////NOTAS
        //                                    //newFunctions_4.openInnerNote(desktopSession, file);
        //                                    //Thread.Sleep(1000);

        //                                    //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Nom, Nom, Campo, c, ErrorValidacion, "No se agregó correctamente", 0, "1");


        //                                    //MODIFICAMOS
        //                                    //AGREGAR REGISTRO Principal
        //                                    Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
        //                                    botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
        //                                    var ElementList = PruebaCRUD.NavClass(desktopSession);

        //                                    //QBE
        //                                    errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
        //                                    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                    newFunctions_1.lapiz(desktopSession);

        //                                    Thread.Sleep(1000);
        //                                    ////ELIMINAR REGISTRO
        //                                    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
        //                                    desktopSession.Mouse.Click(null);
        //                                    Thread.Sleep(500);
        //                                    newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
        //                                    PruebaCRUD.cerrarBorrar(desktopSession);
        //                                    Thread.Sleep(500);
        //                                    newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
        //                                    Thread.Sleep(1000);

        //                                    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
        //                                    desktopSession.Mouse.Click(null);

        //                                    FuncionesKGnAlert.AgregarCrud(desktopSession, Notifi, crudPrincipal, file);
        //                                    newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

        //                                    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
        //                                    desktopSession.Mouse.Click(null);
        //                                    newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);
        //                                    Thread.Sleep(1500);


        //                                }
        //                                else
        //                                {
        //                                    ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
        //                                }

        //                                if (ErrorValidacion.Count > 0)
        //                                {
        //                                    ReporteErrores(ErrorValidacion, codProgram, motor, Modulo);
        //                                    var separator = string.Format("{0}{0}", Environment.NewLine);
        //                                    string errorMessageString = string.Join(separator, ErrorValidacion);
        //                                    Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}", Environment.NewLine, errorMessageString));
        //                                }

        //                                bandera = true;

        //                                DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
        //                                string SthCount = null;
        //                                foreach (DataRow rowsta in Sth.Tables[0].Rows)
        //                                {
        //                                    SthCount = rowsta["CountDes"].ToString();
        //                                    int StCount = Int32.Parse(SthCount);

        //                                    if (StCount > 0)
        //                                    {
        //                                        int NewCount = StCount - 1;
        //                                        DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
        //                                        if (NewCount == 0)
        //                                        {
        //                                            DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
        //                                            break;
        //                                        }
        //                                    }
        //                                }
        //                            }
        //                            catch (Exception e)
        //                            {
        //                                Thread.Sleep(500);
        //                                bandera = true;
        //                                DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
        //                                string SthCount = null;
        //                                foreach (DataRow rowsta in Sth.Tables[0].Rows)
        //                                {
        //                                    SthCount = rowsta["CountDes"].ToString();

        //                                    int StCount = Int32.Parse(SthCount);
        //                                    if (StCount > 0)
        //                                    {
        //                                        int NewCount = StCount - 1;
        //                                        DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
        //                                        if (NewCount == 0)
        //                                        {
        //                                            DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
        //                                            //break;
        //                                        }
        //                                    }
        //                                }
        //                                AbrirPrograma a = new AbrirPrograma();
        //                                a.Stop();
        //                                //break;
        //                            }
        //                        }

        //                    }
        //                }
        //            }
        //            else
        //            {
        //                errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: El nombre de la automatizacion no corresponde");
        //            }
        //        }
        //        if (bandera == false)
        //        {
        //            if (errorMessages.Count > 0)
        //            {
        //                var separator = string.Format("{0}{0}", Environment.NewLine);
        //                string errorMessageString = string.Join(separator, errorMessages);

        //                Assert.Inconclusive(string.Format("Las condiciones de ejecucion del caso son:{0}{1}",
        //                             Environment.NewLine, errorMessageString));
        //            }
        //        }

        //    }
        //}

        [TestMethod]
        public void AbrirPrograma()
        {

            string codProgram = "KNmComps";
            string user = "UNatalia";
            string motor = "SQL";

            AbrirPrograma a = new AbrirPrograma(codProgram, user);
            desktopSession = a.Start2(motor, "");


        }

        [TestMethod]
        public void AbrirPrograma2()
        {

            string codProgram = "KNmComps";
            string user = "ONATALIA";
            string motor = "ORA";

            AbrirPrograma a = new AbrirPrograma(codProgram, user);
            desktopSession = a.Start2(motor, "");


        }

        [TestMethod]
        public void ModulosAlertasVencimientos()
        {

            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            string motor = string.Empty;
            string codProgram = string.Empty;

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
            //TableOrder = "DW_A1705";
            DataSet OrderExecutionCase = SqlAdapter.SelectOrderExecution("T", TableOrder);
            int NumCasAgen = OrderExecutionCase.Tables[0].Rows.Count;
            if (NumCasAgen < 1)
            {
                errorMessages.Add("No hay casos en el agendamiento");
            }
            //Debugger.Launch();
            foreach (DataRow rowsi in OrderExecutionCase.Tables[0].Rows)
            {
                string plans = rowsi["plans"].ToString();
                string suite = rowsi["suite"].ToString();
                string CaseId = rowsi["CaseId"].ToString();
                string orders = rowsi["orders"].ToString();
                string states = rowsi["states"].ToString();
                string methodname = rowsi["methodname"].ToString();
                string CountDes = rowsi["CountDes"].ToString();

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesoKGnAlert.ModulosProcesoKGnAlert.ModulosAlertasVencimientos")
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
                            //int velocidad = 10;

                            //Playback.PlaybackSettings.DelayBetweenActions = velocidad;

                            //@User @Motor @CodProgram @NomProgram @TipoQbe @QbeFilter @banExcel @BanderaLupa @BanderaSum @BanderaPreli
                            //@Identificacion @Concepto @Ausencia @Observaciones

                            if (
                                //Datos Obligatorios
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&

                                //Datos Generales
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&

                                //Datos Crud Principal
                                rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
                                rows["Notifi"].ToString().Length != 0 && rows["Notifi"].ToString() != null &&
                                rows["Usua"].ToString().Length != 0 && rows["Usua"].ToString() != null &&

                                //Variables Validaciones Crud principal
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null)
                            {

                                //Variables obligatorias
                                string user = rows["User"].ToString();
                                motor = rows["Motor"].ToString();
                                codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();

                                //Variables generales
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();

                                //Variables validacion  crud Principal
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string EditCampo = rows["EditCampo"].ToString();

                                //Variables crud principal
                                string Nombre = rows["Nombre"].ToString();
                                string Notifi = rows["Notifi"].ToString();
                                string Usua = rows["Usua"].ToString();

                                String campo = "1";

                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { Nombre, Notifi, Usua };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                //Debugger.Launch();
                                string nomCapture = codProgram + motor + Hora();
                                string file = CrearDocumentoWordDinamico(nomCapture, "Alertas_KGnAlert", motor);

                                try
                                {
                                    ErrorValidacion.Clear();
                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");


                                    //Thread.Sleep(15000);
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
                                        ////VALIDACION CODIGO DEL PROGRAMA
                                        //errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        ////VALIDACION NOMBRE DEL PROGRAMA
                                        //errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Nom, Nom, Campo, c, ErrorValidacion, "No se agregó correctamente", 0, "1");


                                        //MODIFICAMOS
                                        //AGREGAR REGISTRO Principal
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_1.lapiz(desktopSession);

                                        Thread.Sleep(1000);
                                        ////ELIMINAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
                                        PruebaCRUD.cerrarBorrar(desktopSession);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
                                        Thread.Sleep(1000);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        FuncionesKGnAlert.AgregarCrud(desktopSession, Notifi, crudPrincipal, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);
                                        Thread.Sleep(1500);

                                        FuncionesKGnAlert.BotonEjecutar(desktopSession, file);
                                        Thread.Sleep(1500);

                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }
                                    if (ErrorValidacion.Count > 0)
                                    {
                                        ReporteErrores(ErrorValidacion, codProgram, motor, "ST_SP_107");
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, ErrorValidacion);
                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba {2} son:{0}{1}", Environment.NewLine, errorMessageString, codProgram));
                                    }
                                    a.Stop();
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
                                                //break;
                                            }
                                        }
                                    }
                                    //AbrirPrograma a = new AbrirPrograma();
                                    //a.Stop();
                                    Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    break;
                                }
                            }
                            else
                            {
                                errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Valide los datos de entrada");
                            }
                        }

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

    }
}
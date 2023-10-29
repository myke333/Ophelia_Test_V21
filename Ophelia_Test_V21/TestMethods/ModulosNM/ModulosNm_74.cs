using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_9;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.ValidacionQueries;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Text;
using System.Threading;

namespace Ophelia_Test_V21.TestMethods.ModulosNM
{
    /// <summary>
    /// Summary description for ModulosNm_74
    /// </summary>
    [TestClass]
    public class ModulosNm_74 : FuncionesVitales
    { 
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        protected static WindowsDriver<WindowsElement> desktopSession;
        protected static WindowsDriver<WindowsElement> rootSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();
        public static WindowsDriver<WindowsElement> RootSession()
        {
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }
        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string Clase)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(Clase);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }
        public ModulosNm_74()
        {
        }
        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
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

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion
        [TestMethod]
        public void AbrirPrograma()
        {

            string programa = "KNmRacio";
            AbrirPrograma a = new AbrirPrograma(programa, "UCalida1");
            desktopSession = a.Start2("SQL", "");
            //Thread.Sleep(2000000);
            //List<string> crudData = new List<string>() { };
            string file = CrearDocumentoWordDinamico("prueba", "Pruebas", "SQL");
            Thread.Sleep(200000);

            //////AGREGAR REGISTRO Principal
            Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
            botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
            var ElementList = PruebaCRUD.NavClass(desktopSession);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            desktopSession.Mouse.Click(null);

            //CrudKNmrproa.AgregarCrud(desktopSession);
            Thread.Sleep(2000000);

            //CrudKnmCaaut.AgregarCrud(desktopSession, 0, crudPrincipal);
            //Thread.Sleep(2000000);

            //newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);
            ////newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

            ////Thread.Sleep(1500);

            ////////Validacion Agregar
            //////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, RolNum, RolNum, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

            ////////DESCARTAR EDICION
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            //desktopSession.Mouse.Click(null);
            ////newFunctions_4.ScreenshotNuevo("Editar", true, file);

            //CrudKnmvarig.AgregarCrud(desktopSession, 1, crudPrincipal);

            ////CrudKed3crol.CRUD(desktopSession, 2, crudVariables, file);
            ////newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            //desktopSession.Mouse.Click(null);
            ////newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

            //Thread.Sleep(1000);

            //////Validacion Editar Descartar
            //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, RolNum, RolPal, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

            //////ACEPTAR EDICION
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            //desktopSession.Mouse.Click(null);
            ////newFunctions_4.ScreenshotNuevo("Editar", true, file);

            //CrudKnmvarig.AgregarCrud(desktopSession, 1, crudPrincipal);

            ////CrudKed3crol.CRUD(desktopSession, 2, crudVariables, file);
            //////newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);
            ////newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

            ///*Thread.Sleep(1000);
            ////validacion error al editar
            //bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
            //if (valEditOra == true)
            //{
            //    ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
            //    //LUPA
            //    errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
            //    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            //    //ACEPTAR EDICION
            //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            //    desktopSession.Mouse.Click(null);
            //    newFunctions_4.ScreenshotNuevo("Editar", true, file);

            //    CrudKed3crol.CRUD(desktopSession, 2, crudVariables, file);
            //    newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

            //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            //    desktopSession.Mouse.Click(null);
            //    Thread.Sleep(3000);
            //}
            //newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

            //////Validacion Editar
            //Thread.Sleep(1500);
            //// //Debugger.Launch();
            //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, RolNum, "E", c, ErrorValidacion);
            //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, RolNum, EditarRol, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);*/

            //Thread.Sleep(1000);

            //ELIMINAR REGISTRO CRUD

            //PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
            //newFunctions_4.ScreenshotNuevo("Datos Eliminados", true, file);

            //VALIDACIÓN ELIMINAR REGISTRO
            //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla1, motor, user, Campo1, Codigo, "D", c, ErrorValidacion);
            //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Codigo, "", Campo1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);
        }
        [TestMethod]
        public void KNmRproAChecklist()
        {
            List<string> errorMessages = new List<string>();
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
            ////TableOrder = "ktest1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_74.KNmRproAChecklist")
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
                                //Datos Obligatorios
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&

                                //Datos Generales
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&

                                //Datos Cruds
                                rows["FechaIni"].ToString().Length != 0 && rows["FechaIni"].ToString() != null &&
                                rows["FechaFin"].ToString().Length != 0 && rows["FechaFin"].ToString() != null)
                            {
                                //Variables obligatorias
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();

                                //Variables generales
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();

                                //Variables crud principal
                                string FechaIni = rows["FechaIni"].ToString();
                                string FechaFin = rows["FechaFin"].ToString();

                                //string Tabla = rows["Tabla"].ToString();
                                //string Campo = rows["Campo"].ToString();
                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { FechaIni, FechaFin };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);

                                try
                                {
                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();

                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
                                    ////Debugger.Launch();
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
                                        //VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR REGISTRO PEGADO
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, CampoV, FechaIni, FechaFin, CampoV, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        ////VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //Modificamos

                                        //////AGREGAR REGISTRO
                                        //Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        //botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        //var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        //desktopSession.Mouse.Click(null);

                                        CrudKNmrproa.AgregarCrud(desktopSession, 0, crudPrincipal);
                                        Thread.Sleep(2000000);

                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        //Thread.Sleep(1500);
                                        //Debugger.Launch();
                                        //////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, CampoV, NombreRubro, NombreRubro, CampoV, c, ErrorValidacion, "No se agregó correctamente", 0);

         

                                        //Thread.Sleep(1000);
                                        //validacion error al editar
                                        //bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        //if (valEditOra == true)
                                        //{
                                        //    ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
                                        //    //LUPA
                                        //    errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //    //////ACEPTAR EDICION
                                        //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        //    CrudKnmvarig.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        //    newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    Thread.Sleep(3000);
                                        //}
                                        //newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //REPORTE PRELIMINAR
                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        //ELIMINAR REGISTRO CRUD

                                        ////LUPA
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        //newFunctions_4.ScreenshotNuevo("Datos Eliminados", true, file);
                                        //////VALIDACIÓN ELIMINAR REGISTRO
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, CampoV, NombreRubro, "D", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, CampoV, NombreRubro, "", CampoV, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);


                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }

                                    if (ErrorValidacion.Count > 0)
                                    {
                                        ReporteErrores(ErrorValidacion, codProgram, motor, Modulo);
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

        [TestMethod]
        public void KNmPrsruChecklist()
        {
            List<string> errorMessages = new List<string>();
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
            ////TableOrder = "ktest1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_74.KNmPrsruChecklist")
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

                                //Datos Cruds
                                rows["NombreRubro"].ToString().Length != 0 && rows["NombreRubro"].ToString() != null &&
                                rows["CenCosto"].ToString().Length != 0 && rows["CenCosto"].ToString() != null &&
                                rows["PresuUno"].ToString().Length != 0 && rows["PresuUno"].ToString() != null &&
                                rows["PresuDos"].ToString().Length != 0 && rows["PresuDos"].ToString() != null &&
                                rows["PresuTres"].ToString().Length != 0 && rows["PresuTres"].ToString() != null &&
                                rows["PresuCuatro"].ToString().Length != 0 && rows["PresuCuatro"].ToString() != null &&
                                rows["ActPresuCuatro"].ToString().Length != 0 && rows["ActPresuCuatro"].ToString() != null &&

                                //Datos Validacion
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["CampoV"].ToString().Length != 0 && rows["CampoV"].ToString() != null &&
                                rows["ActCampo"].ToString().Length != 0 && rows["ActCampo"].ToString() != null)
                            {
                                //Variables obligatorias
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();

                                //Variables generales
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();

                                //Variables crud principal
                                string NombreRubro = rows["NombreRubro"].ToString();
                                string CenCosto = rows["CenCosto"].ToString();
                                string PresuUno = rows["PresuUno"].ToString();
                                string PresuDos = rows["PresuDos"].ToString();
                                string PresuTres = rows["PresuTres"].ToString();
                                string PresuCuatro = rows["PresuCuatro"].ToString();
                                string ActPresuCuatro = rows["ActPresuCuatro"].ToString();

                                //Variables validacion
                                string Tabla = rows["Tabla"].ToString();
                                string CampoV = rows["CampoV"].ToString();
                                string ActCampo = rows["ActCampo"].ToString();

                                //string Tabla = rows["Tabla"].ToString();
                                //string Campo = rows["Campo"].ToString();
                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { NombreRubro, CenCosto,  PresuUno, PresuTres, PresuCuatro, ActPresuCuatro };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);

                                try
                                {
                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();

                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
                                    ////Debugger.Launch();
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
                                        //VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR REGISTRO PEGADO
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, CampoV, NombreRubro, PresuCuatro, CampoV, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        ////VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //Modificamos

                                        ////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKNmPrsru.AgregarCrud(desktopSession, 0, crudPrincipal);
                                        //Thread.Sleep(2000000);

                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);
                                        //Debugger.Launch();
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, CampoV, NombreRubro, NombreRubro, CampoV, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKNmPrsru.AgregarCrud(desktopSession, 1, crudPrincipal);

                                        //CrudKNmPrsru.CRUD(desktopSession, 2, crudVariables, file);
                                        //newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        //Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, CampoV, NombreRubro, PresuCuatro, ActCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        //////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKNmPrsru.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);


                                        Thread.Sleep(1000);
                                        //validacion error al editar
                                        //bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        //if (valEditOra == true)
                                        //{
                                        //    ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
                                        //    //LUPA
                                        //    errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //    //////ACEPTAR EDICION
                                        //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        //    CrudKnmvarig.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        //    newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    Thread.Sleep(3000);
                                        //}
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        //Validacion Editar
                                        Thread.Sleep(1500);
                                        //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, CampoV, NombreRubro, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, CampoV, NombreRubro, ActPresuCuatro, ActCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //Thread.Sleep(1000);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_1.lapiz(desktopSession);

                                        //REPORTE DINAMICO
                                        string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
                                        errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF DINAMICO
                                        listaResultado = Textopdf(pdf, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //REPORTE PRELIMINAR
                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR REGISTRO CRUD

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados", true, file);
                                        ////VALIDACIÓN ELIMINAR REGISTRO
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, CampoV, NombreRubro, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, CampoV, NombreRubro, "", CampoV, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);


                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }

                                    if (ErrorValidacion.Count > 0)
                                    {
                                        ReporteErrores(ErrorValidacion, codProgram, motor, Modulo);
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


        //[TestMethod]
        //public void KNmPrtpuChecklist()
        //{
        //    //Debugger.Launch();
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

        //        if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_68.KNmPhoexChecklist")
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

        //                        //Variables Validaciones
        //                        rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
        //                        rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
        //                        rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&

        //                        //Variables Validaciones detalle 1
        //                        rows["TablaDetalle1"].ToString().Length != 0 && rows["TablaDetalle1"].ToString() != null &&
        //                        rows["CampoDetalle1"].ToString().Length != 0 && rows["CampoDetalle1"].ToString() != null &&
        //                        rows["EditCampoDetalle1"].ToString().Length != 0 && rows["EditCampoDetalle1"].ToString() != null &&

        //                        //Variables Validaciones detalle 2
        //                        rows["TablaDetalle2"].ToString().Length != 0 && rows["TablaDetalle2"].ToString() != null &&
        //                        rows["CampoDetalle2"].ToString().Length != 0 && rows["CampoDetalle2"].ToString() != null &&
        //                        rows["EditCampoDetalle2"].ToString().Length != 0 && rows["EditCampoDetalle2"].ToString() != null &&

        //                        //Variables Validaciones detalle 3
        //                        rows["TablaDetalle3"].ToString().Length != 0 && rows["TablaDetalle3"].ToString() != null &&
        //                        rows["CampoDetalle3"].ToString().Length != 0 && rows["CampoDetalle3"].ToString() != null &&
        //                        rows["EditCampoDetalle3"].ToString().Length != 0 && rows["EditCampoDetalle3"].ToString() != null &&

        //                        //Datos Cruds
        //                        rows["Codigo"].ToString().Length != 0 && rows["Codigo"].ToString() != null &&
        //                        rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
        //                        rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
        //                        rows["ActNombre"].ToString().Length != 0 && rows["ActNombre"].ToString() != null &&

        //                        rows["Cargo"].ToString().Length != 0 && rows["Cargo"].ToString() != null &&
        //                        rows["Dependencia"].ToString().Length != 0 && rows["Dependencia"].ToString() != null &&
        //                        rows["Horas"].ToString().Length != 0 && rows["Horas"].ToString() != null &&
        //                        rows["ActHoras"].ToString().Length != 0 && rows["ActHoras"].ToString() != null &&

        //                        rows["CargoDet2"].ToString().Length != 0 && rows["CargoDet2"].ToString() != null &&
        //                        rows["HorasMin"].ToString().Length != 0 && rows["HorasMin"].ToString() != null &&
        //                        rows["HorasMax"].ToString().Length != 0 && rows["HorasMax"].ToString() != null &&
        //                        rows["ActHorasMax"].ToString().Length != 0 && rows["ActHorasMax"].ToString() != null &&

        //                        rows["CargoDet3"].ToString().Length != 0 && rows["CargoDet3"].ToString() != null &&
        //                        rows["DependenciaDet3"].ToString().Length != 0 && rows["DependenciaDet3"].ToString() != null &&
        //                        rows["ActDependenciaDet3"].ToString().Length != 0 && rows["ActDependenciaDet3"].ToString() != null)

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

        //                        //Variables validacion
        //                        string Tabla = rows["Tabla"].ToString();
        //                        string Campo = rows["Campo"].ToString();
        //                        string EditCampo = rows["EditCampo"].ToString();

        //                        //Variables validacion DETALLE 1
        //                        string TablaDetalle1 = rows["TablaDetalle1"].ToString();
        //                        string CampoDetalle1 = rows["CampoDetalle1"].ToString();
        //                        string EditCampoDetalle1 = rows["EditCampoDetalle1"].ToString();

        //                        //Variables validacion DETALLE 2
        //                        string TablaDetalle2 = rows["TablaDetalle2"].ToString();
        //                        string CampoDetalle2 = rows["CampoDetalle2"].ToString();
        //                        string EditCampoDetalle2 = rows["EditCampoDetalle2"].ToString();

        //                        //Variables validacion DETALLE 3
        //                        string TablaDetalle3 = rows["TablaDetalle3"].ToString();
        //                        string CampoDetalle3 = rows["CampoDetalle3"].ToString();
        //                        string EditCampoDetalle3 = rows["EditCampoDetalle3"].ToString();

        //                        //Variables crud principal
        //                        string Codigo = rows["Codigo"].ToString();
        //                        string Nombre = rows["Nombre"].ToString();
        //                        string Fecha = rows["Fecha"].ToString();
        //                        string ActNombre = rows["ActNombre"].ToString();

        //                        //Variables Crud Det 1
        //                        string Cargo = rows["Cargo"].ToString();
        //                        string Dependencia = rows["Dependencia"].ToString();
        //                        string Horas = rows["Horas"].ToString();
        //                        string ActHoras = rows["ActHoras"].ToString();

        //                        //Variables Crud Det 2
        //                        string CargoDet2 = rows["CargoDet2"].ToString();
        //                        string HorasMin = rows["HorasMin"].ToString();
        //                        string HorasMax = rows["HorasMax"].ToString();
        //                        string ActHorasMax = rows["ActHorasMax"].ToString();

        //                        //Variables Crud Det 3
        //                        string CargoDet3 = rows["CargoDet3"].ToString();
        //                        string DependenciaDet3 = rows["DependenciaDet3"].ToString();
        //                        string ActDependenciaDet3 = rows["ActDependenciaDet3"].ToString();

        //                        //string Tabla = rows["Tabla"].ToString();
        //                        //string Campo = rows["Campo"].ToString();


        //                        //LISTA CRUD PRINCIPAL
        //                        List<string> crudPrincipal = new List<string>() { Codigo, Nombre, Fecha, ActNombre };
        //                        //LISTA DETALLE 1
        //                        List<string> crudDet1 = new List<string>() { Cargo, Dependencia, Horas, ActHoras };
        //                        //LISTA DETALLE 2
        //                        List<string> crudDet2 = new List<string>() { CargoDet2, HorasMin, HorasMax, ActHorasMax };
        //                        //LISTA DETALLE 3
        //                        List<string> crudDet3 = new List<string>() { CargoDet3, DependenciaDet3, ActDependenciaDet3 };

        //                        //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
        //                        string file = CrearDocumentoWordDinamico(codProgram, "ED", motor);

        //                        try
        //                        {
        //                            List<string> ErrorValidacion = new List<string>();
        //                            List<string> errors = new List<string>();

        //                            AbrirPrograma a = new AbrirPrograma(codProgram, user);
        //                            desktopSession = a.Start2(motor, "");
        //                            ////Debugger.Launch();
        //                            bool leeprograma = true;

        //                            if (desktopSession != null)
        //                            {
        //                                string str = desktopSession.PageSource;
        //                                if (str != null)
        //                                {
        //                                    if (!str.Contains("TPanel"))
        //                                    {
        //                                        leeprograma = false;
        //                                    }
        //                                }
        //                                else
        //                                {
        //                                    leeprograma = false;
        //                                }
        //                            }
        //                            else
        //                            {
        //                                leeprograma = false;
        //                            }

        //                            if (leeprograma)
        //                            {
        //                                newFunctions_4.ScreenshotNuevo("Abrir Programa", true, file);
        //                                //VALIDACION CODIGO DEL PROGRAMA
        //                                errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                //VALIDACION NOMBRE DEL PROGRAMA
        //                                errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                //VERSION
        //                                FuncionesGlobales.GetVersion(desktopSession, file);
        //                                Thread.Sleep(1000);
        //                                //NOTAS
        //                                newFunctions_4.openInnerNote(desktopSession, file);
        //                                Thread.Sleep(1000);

        //                                //MODIFICAMOS
        //                                ////AGREGAR REGISTRO Principal
        //                                Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
        //                                botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
        //                                var ElementList = PruebaCRUD.NavClass(desktopSession);
        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
        //                                desktopSession.Mouse.Click(null);

        //                                CrudKnmphoex.AgregarCrud(desktopSession, 0, crudPrincipal);
        //                                newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

        //                                Thread.Sleep(1500);

        //                                ////Validacion Agregar
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

        //                                ////DESCARTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                CrudKnmphoex.AgregarCrud(desktopSession, 1, crudPrincipal);
        //                                newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

        //                                Thread.Sleep(1000);

        //                                ////Validacion Editar Descartar
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

        //                                ////ACEPTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                CrudKnmphoex.AgregarCrud(desktopSession, 1, crudPrincipal);
        //                                newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

        //                                Thread.Sleep(3000);
        //                                ////validacion error al editar
        //                                bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
        //                                //Debugger.Launch();
        //                                if (valEditOra == true)
        //                                {
        //                                    ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
        //                                    //LUPA
        //                                    errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                    newFunctions_3.LupaAud(desktopSession, "1", file);
        //                                    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                    //ACEPTAR EDICION
        //                                    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
        //                                    desktopSession.Mouse.Click(null);
        //                                    newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                    CrudKnmphoex.AgregarCrud(desktopSession, 1, crudPrincipal);
        //                                    newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
        //                                    desktopSession.Mouse.Click(null);
        //                                    Thread.Sleep(3000);
        //                                }
        //                                newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

        //                                //////Validacion Editar
        //                                Thread.Sleep(1500);
        //                                ////// //Debugger.Launch();
        //                                ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "E", c, ErrorValidacion);
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, ActNombre, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

        //                                Thread.Sleep(1000);

        //                                //LUPA
        //                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                ////AGREGAR REGISTRO DETALLE 1                                        
        //                                Dictionary<string, Point> botonesNavegador2 = new Dictionary<string, Point>();
        //                                botonesNavegador2 = PruebaCRUD.findname(desktopSession, 28, 1);
        //                                var ElementList2 = PruebaCRUD.NavClass(desktopSession);
        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Insertar"].X, botonesNavegador2["Insertar"].Y);
        //                                desktopSession.Mouse.Click(null);

        //                                CrudKnmphoex.AgregarCrud1(desktopSession, 0, crudDet1);
        //                                newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

        //                                Thread.Sleep(1500);

        //                                ////Validacion Agregar
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

        //                                ////DESCARTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                CrudKnmphoex.AgregarCrud1(desktopSession, 1, crudDet1);
        //                                newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Cancelar"].X, botonesNavegador2["Cancelar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

        //                                Thread.Sleep(1000);

        //                                ////Validacion Editar Descartar
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

        //                                ////ACEPTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                CrudKnmphoex.AgregarCrud1(desktopSession, 1, crudDet1);
        //                                newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

        //                                Thread.Sleep(3000);
        //                                //validacion error al editar
        //                                bool valEditOra1 = PruebaCRUD.ValEditOra(desktopSession);
        //                                if (valEditOra1 == true)
        //                                {
        //                                    ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
        //                                    //LUPA
        //                                    errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                    newFunctions_3.LupaAud(desktopSession, "1", file);
        //                                    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                    //ACEPTAR EDICION
        //                                    desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
        //                                    desktopSession.Mouse.Click(null);
        //                                    newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                    CrudKnmphoex.AgregarCrud1(desktopSession, 1, crudDet1);
        //                                    newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                    desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
        //                                    desktopSession.Mouse.Click(null);
        //                                    Thread.Sleep(3000);
        //                                }
        //                                newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

        //                                ////Validacion Editar
        //                                Thread.Sleep(1500);
        //                                //Debugger.Launch();
        //                                ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "E", c, ErrorValidacion);
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, ActNombre, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

        //                                Thread.Sleep(2000);

        //                                CrudKnmphoex.ClickButton(desktopSession, 0);

        //                                Thread.Sleep(2000);

        //                                ////AGREGAR REGISTRO DETALLE 2

        //                                Dictionary<string, Point> botonesNavegador3 = new Dictionary<string, Point>();
        //                                botonesNavegador3 = PruebaCRUD.findname(desktopSession, 25, 1);
        //                                var ElementList3 = PruebaCRUD.NavClass(desktopSession);
        //                                desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Insertar"].X, botonesNavegador3["Insertar"].Y);
        //                                desktopSession.Mouse.Click(null);

        //                                CrudKnmphoex.AgregarCrud2(desktopSession, 0, crudDet2);
        //                                newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Aplicar"].X, botonesNavegador3["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

        //                                Thread.Sleep(1500);

        //                                ////Validacion Agregar
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

        //                                ////DESCARTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Editar"].X, botonesNavegador3["Editar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                CrudKnmphoex.AgregarCrud2(desktopSession, 1, crudDet2);
        //                                newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Cancelar"].X, botonesNavegador3["Cancelar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

        //                                Thread.Sleep(2000);

        //                                ////Validacion Editar Descartar
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

        //                                ////ACEPTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Editar"].X, botonesNavegador3["Editar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                CrudKnmphoex.AgregarCrud2(desktopSession, 1, crudDet2);
        //                                newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                Thread.Sleep(2000);
        //                                desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Aplicar"].X, botonesNavegador3["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);

        //                                Thread.Sleep(3000);
        //                                //validacion error al editar
        //                                bool valEditOra2 = PruebaCRUD.ValEditOra(desktopSession);
        //                                if (valEditOra2 == true)
        //                                {
        //                                    ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
        //                                    //LUPA
        //                                    errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                    ////ACEPTAR EDICION
        //                                    desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Editar"].X, botonesNavegador3["Editar"].Y);
        //                                    desktopSession.Mouse.Click(null);
        //                                    newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                    CrudKnmphoex.AgregarCrud2(desktopSession, 1, crudDet2);
        //                                    newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                    desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Aplicar"].X, botonesNavegador3["Aplicar"].Y);
        //                                    desktopSession.Mouse.Click(null);
        //                                    Thread.Sleep(3000);
        //                                }
        //                                newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

        //                                ////Validacion Editar
        //                                Thread.Sleep(1500);
        //                                // //Debugger.Launch();
        //                                ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "E", c, ErrorValidacion);
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

        //                                Thread.Sleep(2000);

        //                                CrudKnmphoex.ClickButton(desktopSession, 1);

        //                                Thread.Sleep(2000);

        //                                ////AGREGAR REGISTRO DETALLE 3

        //                                Dictionary<string, Point> botonesNavegador4 = new Dictionary<string, Point>();
        //                                botonesNavegador4 = PruebaCRUD.findname(desktopSession, 26, 1);
        //                                var ElementList4 = PruebaCRUD.NavClass(desktopSession);
        //                                desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Insertar"].X, botonesNavegador4["Insertar"].Y);
        //                                desktopSession.Mouse.Click(null);

        //                                CrudKnmphoex.AgregarCrud3(desktopSession, 0, crudDet3);
        //                                newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Aplicar"].X, botonesNavegador4["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

        //                                Thread.Sleep(1500);

        //                                ////Validacion Agregar
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

        //                                ////DESCARTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Editar"].X, botonesNavegador4["Editar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                CrudKnmphoex.AgregarCrud3(desktopSession, 1, crudDet3);
        //                                newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Cancelar"].X, botonesNavegador4["Cancelar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

        //                                Thread.Sleep(1000);

        //                                ////Validacion Editar Descartar
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

        //                                ////ACEPTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Editar"].X, botonesNavegador4["Editar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                CrudKnmphoex.AgregarCrud3(desktopSession, 1, crudDet3);
        //                                newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Aplicar"].X, botonesNavegador4["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

        //                                Thread.Sleep(3000);
        //                                //validacion error al editar
        //                                bool valEditOra3 = PruebaCRUD.ValEditOra(desktopSession);
        //                                if (valEditOra3 == true)
        //                                {
        //                                    ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
        //                                    //LUPA
        //                                    errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                    ////ACEPTAR EDICION
        //                                    desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Editar"].X, botonesNavegador4["Editar"].Y);
        //                                    desktopSession.Mouse.Click(null);
        //                                    newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                    CrudKnmphoex.AgregarCrud3(desktopSession, 1, crudDet3);
        //                                    newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                    desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Aplicar"].X, botonesNavegador4["Aplicar"].Y);
        //                                    desktopSession.Mouse.Click(null);
        //                                    Thread.Sleep(3000);
        //                                }
        //                                newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

        //                                ////Validacion Editar
        //                                Thread.Sleep(1500);
        //                                // //Debugger.Launch();
        //                                ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "E", c, ErrorValidacion);
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, ActNombre, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

        //                                //Thread.Sleep(1000);

        //                                //ClasePrueba.Clickbutton(desktopSession, 2);

        //                                //QBE
        //                                errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                //SUMATORIA
        //                                errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                newFunctions_1.lapiz(desktopSession);

        //                                //REPORTE DINAMICO
        //                                string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
        //                                errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                //Rectificar Pie Pagina PDF DINAMICO
        //                                listaResultado = Textopdf(pdf, codProgram, user);
        //                                listaResultado.ForEach(e => ErrorValidacion.Add(e));

        //                                //REPORTE PRELIMINAR
        //                                string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
        //                                errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                //Rectificar Pie Pagina PDF PRELIMINAR
        //                                listaResultado = Textopdf(pdf1, codProgram, user);
        //                                listaResultado.ForEach(e => ErrorValidacion.Add(e));

        //                                //EXPORTAR EXCEL
        //                                string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
        //                                errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                //RECTIFICACION DE CAMPOS DE EXCEL       
        //                                /*ring COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, tipo, c);
        //                                string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
        //                                Celda = Celda_Excel(ruta, nom_empr);
        //                                Celda.ForEach(u => ErrorValidacion.Add(u));*/

        //                                Thread.Sleep(2000);

        //                                CrudKnmphoex.ClickButton(desktopSession, 1);
        //                                //LUPA
        //                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                PruebaCRUD.EliminarRegistro(desktopSession, ElementList4[1], botonesNavegador4, file);
        //                                newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 3", true, file);
        //                                //VALIDACIÓN ELIMINAR REGISTRO
        //                                ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle3, motor, user, CampoDetalle3, Codigo, "D", c, ErrorValidacion);
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle3, user, motor, CampoDetalle3, Codigo, "", CampoDetalle3, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

        //                                Thread.Sleep(3000);

        //                                CrudKnmphoex.ClickButton(desktopSession, 0);

        //                                Thread.Sleep(3000);

        //                                //LUPA
        //                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                PruebaCRUD.EliminarRegistro(desktopSession, ElementList3[1], botonesNavegador3, file);
        //                                newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 2", true, file);
        //                                //VALIDACIÓN ELIMINAR REGISTRO
        //                                ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle2, motor, user, CampoDetalle2, Codigo, "D", c, ErrorValidacion);
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle2, user, motor, CampoDetalle2, Codigo, "", CampoDetalle2, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

        //                                Thread.Sleep(2000);

        //                                CrudKnmphoex.ClickButton(desktopSession, 2);

        //                                Thread.Sleep(2000);

        //                                //LUPA
        //                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                PruebaCRUD.EliminarRegistro(desktopSession, ElementList2[1], botonesNavegador2, file);
        //                                newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 1", true, file);
        //                                //VALIDACIÓN ELIMINAR REGISTRO
        //                                ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle1, motor, user, CampoDetalle1, Codigo, "D", c, ErrorValidacion);
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, "", CampoDetalle1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

        //                                Thread.Sleep(2000);

        //                                //LUPA
        //                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
        //                                newFunctions_4.ScreenshotNuevo("Datos Eliminados Principal", true, file);
        //                                //VALIDACIÓN ELIMINAR REGISTRO
        //                                ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "D", c, ErrorValidacion);
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);


        //                            }
        //                            else
        //                            {
        //                                ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
        //                            }

        //                            if (ErrorValidacion.Count > 0)
        //                            {
        //                                ReporteErrores(ErrorValidacion, codProgram, motor, Modulo);
        //                                var separator = string.Format("{0}{0}", Environment.NewLine);
        //                                string errorMessageString = string.Join(separator, ErrorValidacion);
        //                                Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}", Environment.NewLine, errorMessageString));
        //                            }

        //                            Thread.Sleep(4000);

        //                            bandera = true;

        //                            DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
        //                            string SthCount = null;
        //                            foreach (DataRow rowsta in Sth.Tables[0].Rows)
        //                            {
        //                                SthCount = rowsta["CountDes"].ToString();
        //                                int StCount = Int32.Parse(SthCount);

        //                                if (StCount > 0)
        //                                {
        //                                    int NewCount = StCount - 1;
        //                                    DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
        //                                    if (NewCount == 0)
        //                                    {
        //                                        DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
        //                                        break;
        //                                    }
        //                                }
        //                            }
        //                            break;

        //                        }
        //                        catch (Exception e)
        //                        {
        //                            Thread.Sleep(500);
        //                            bandera = true;
        //                            DataSet Sth = SqlAdapter.SelectOrderExecution("P", TableOrder, plans, suite, CaseId);
        //                            string SthCount = null;
        //                            foreach (DataRow rowsta in Sth.Tables[0].Rows)
        //                            {
        //                                SthCount = rowsta["CountDes"].ToString();

        //                                int StCount = Int32.Parse(SthCount);
        //                                if (StCount > 0)
        //                                {
        //                                    int NewCount = StCount - 1;
        //                                    DataSet DisCount = SqlAdapter.SelectOrderExecution("UP", TableOrder, plans, suite, CaseId, NewCount.ToString());
        //                                    if (NewCount == 0)
        //                                    {
        //                                        DataSet UdpRes = SqlAdapter.SelectOrderExecution("U", TableOrder, plans, suite, CaseId);
        //                                        break;
        //                                    }
        //                                }
        //                            }
        //                            Assert.Fail(CaseId + " ::::::" + e.ToString());
        //                            break;
        //                        }
        //                    }
        //                }
        //                break;
        //            }
        //        }
        //        else
        //        {
        //            errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: El nombre de la automatizacion no corresponde");
        //        }
        //    }
        //    if (bandera == false)
        //    {
        //        if (errorMessages.Count > 0)
        //        {
        //            var separator = string.Format("{0}{0}", Environment.NewLine);
        //            string errorMessageString = string.Join(separator, errorMessages);

        //            Assert.Inconclusive(string.Format("Las condiciones de ejecucion del caso son:{0}{1}",
        //                         Environment.NewLine, errorMessageString));
        //        }
        //    }

        //}
        [TestMethod]
        public void KNmRpnnoChecklist()
        {
            //Debugger.Launch();
            List<string> errorMessages = new List<string>();
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
            ////TableOrder = "ktest1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_68.KNmRpnnoChecklist")
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
                                //Datos Obligatorios
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null )

                            {
                                //Variables obligatorias
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "ED", motor);

                                try
                                {
                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();

                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
                                    ////Debugger.Launch();
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
                                        //VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        ////QBE
                                        //errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //CrudKnmdhono.AgregarCrud(desktopSession);
                                        //newFunctions_4.ScreenshotNuevo("Iniciar Proceso de Novedades", true, file);

                                        //CrudKnmdhono.AgregarCrud1(desktopSession);
                                        //newFunctions_4.ScreenshotNuevo("Tiempo Total Liquidacion", true, file);

                                        //CrudKnmdhono.AgregarCrud2(desktopSession);
                                        //newFunctions_4.ScreenshotNuevo("Finalizado", true, file);

                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }

                                    if (ErrorValidacion.Count > 0)
                                    {
                                        ReporteErrores(ErrorValidacion, codProgram, motor, Modulo);
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
    }
}

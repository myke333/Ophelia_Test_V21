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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo;
using OpenQA.Selenium.Interactions;

namespace Ophelia_Test_V21.TestMethods.ModulosSO
{
    /// <summary>
    /// Descripción resumida de ModulosSO_5
    /// </summary>
    [TestClass]
    public class ModulosSO_5 : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();
        public ModulosSO_5()
        {}

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

        [TestMethod]
        public void KSoMaobsAbrirPrograma()
        {
            string file = CrearDocumentoWordDinamico("KSoMaobs", "SO", "SQL");
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();

            AbrirPrograma a = new AbrirPrograma("KSoMaobs", "UCalida1");
            desktopSession = a.Start2("SQL", "");


            string BanderaLupa = "1";
            string TipoQbe = "0";
            string QbeFilter = "1";
            string BanderaPreli = "1";
            Thread.Sleep(1000);

            //QBE
            errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            Thread.Sleep(1000);
            //REPORTE PRELIMINAR
            string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
            errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
            CrudKsomaobs.PreliKSoMaobs(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli);

            string CentroTrab = "1";
            string Area = "1";
            string Observador = "1";
            string Colaborador = "1";
            string GestionEje = "4";
            string FechaObs = "02/12/2020";
            string AreaSitio = "Prueba1";
            string TareaObs = "Prueba2";
            string DescriHalla = "Prueba3";
            string AccionToma = "Prueba4";
            string AccionTomaEditar = "Prueba4Editar";

            List<string> crudVariables = new List<string>() { CentroTrab, Area, Observador, Colaborador,
                GestionEje, FechaObs, AreaSitio, TareaObs, DescriHalla, AccionToma, AccionTomaEditar
            };

            string FechaI = "7/10/2020";
            string TipoAccion = "C";
            string FechaC = "22/11/2020";
            string FechaIEditar = "18/08/2020";

            List<string> crudVariablesdet1 = new List<string>() { FechaI, TipoAccion, FechaC, FechaIEditar };

            ////LUPA
            //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
            //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

            ////PASAR A ACCIONES
            var TPageControl = desktopSession.FindElementsByClassName("TPageControl");
            desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 45, 15);
            desktopSession.Mouse.Click(null);
            for (int i = 1; i <= 3; i++)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
            }

            ////////AGREGAR REGISTRO
            Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
            botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
            var ElementList = PruebaCRUD.NavClass(desktopSession);

            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsomaobs.CRUDKSoMaobs(desktopSession, 1, crudVariables, file);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(200);
            newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

            ////Validacion Agregar
            //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsomaobs.CRUDKSoMaobs(desktopSession, 2, crudVariables, file);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

            ////Validacion Editar Descartar
            ////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, "OBS_ERVA", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsomaobs.CRUDKSoMaobs(desktopSession, 2, crudVariables, file);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

            //PASAR A DETALLE 1
            desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 45, 10);
            desktopSession.Mouse.Click(null);
            desktopSession.Mouse.Click(null);
            for (int i = 1; i <= 4; i++)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
            }

            ////////AGREGAR DETALLE 1
            botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 1);
            var ElementList1 = PruebaCRUD.NavClass(desktopSession);

            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsomaobs.CRUDDetalle1KSoMaobs(desktopSession, 1, crudVariablesdet1, file);
            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(200);
            newFunctions_4.ScreenshotNuevo("Agregar Detalle 1", true, file);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsomaobs.CRUDDetalle1KSoMaobs(desktopSession, 2, crudVariablesdet1, file);
            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 1", true, file);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsomaobs.CRUDDetalle1KSoMaobs(desktopSession, 2, crudVariablesdet1, file);
            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 1", true, file);

            //ELIMINAR DETALLE 1
            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[1], botonesNavegador, file);
            Thread.Sleep(1000);

            //ELIMINAR REGISTRO
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
            Thread.Sleep(1000);
        }

        [TestMethod]
        public void KSoMaobsCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_5.KSoMaobsCheckList")
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
                                //Datos Login
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&

                                rows["CentroTrab"].ToString().Length != 0 && rows["CentroTrab"].ToString() != null &&
                                rows["Area"].ToString().Length != 0 && rows["Area"].ToString() != null &&
                                rows["Observador"].ToString().Length != 0 && rows["Observador"].ToString() != null &&
                                rows["Colaborador"].ToString().Length != 0 && rows["Colaborador"].ToString() != null &&
                                rows["GestionEje"].ToString().Length != 0 && rows["GestionEje"].ToString() != null &&
                                rows["FechaObs"].ToString().Length != 0 && rows["FechaObs"].ToString() != null &&
                                rows["AreaSitio"].ToString().Length != 0 && rows["AreaSitio"].ToString() != null &&
                                rows["TareaObs"].ToString().Length != 0 && rows["TareaObs"].ToString() != null &&
                                rows["DescriHalla"].ToString().Length != 0 && rows["DescriHalla"].ToString() != null &&
                                rows["AccionToma"].ToString().Length != 0 && rows["AccionToma"].ToString() != null && 
                                rows["AccionTomaEditar"].ToString().Length != 0 && rows["AccionTomaEditar"].ToString() != null &&

                                rows["FechaI"].ToString().Length != 0 && rows["FechaI"].ToString() != null &&
                                rows["TipoAccion"].ToString().Length != 0 && rows["TipoAccion"].ToString() != null &&
                                rows["FechaC"].ToString().Length != 0 && rows["FechaC"].ToString() != null &&
                                rows["FechaIEditar"].ToString().Length != 0 && rows["FechaIEditar"].ToString() != null
                                )
                            {
                                //VARIABLES GENERALES
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();
                                //VARIABLES CRUD
                                string CentroTrab = rows["CentroTrab"].ToString();
                                string Area = rows["Area"].ToString();
                                string Observador = rows["Observador"].ToString();
                                string Colaborador = rows["Colaborador"].ToString();
                                string GestionEje = rows["GestionEje"].ToString();
                                string FechaObs = rows["FechaObs"].ToString();
                                string AreaSitio = rows["AreaSitio"].ToString();
                                string TareaObs = rows["TareaObs"].ToString();
                                string DescriHalla = rows["DescriHalla"].ToString();
                                string AccionToma = rows["AccionToma"].ToString();
                                string AccionTomaEditar = rows["AccionTomaEditar"].ToString();
                                //VARIABLES DETALLE 1
                                string FechaI = rows["FechaI"].ToString();
                                string TipoAccion = rows["TipoAccion"].ToString();
                                string FechaC = rows["FechaC"].ToString();
                                string FechaIEditar = rows["FechaIEditar"].ToString();

                                List<string> crudVariables = new List<string>() { CentroTrab, Area, Observador, Colaborador,
                                GestionEje, FechaObs, AreaSitio, TareaObs, DescriHalla, AccionToma, AccionTomaEditar };

                                List<string> crudVariablesdet1 = new List<string>() { FechaI, TipoAccion, FechaC, FechaIEditar };


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);

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
                                        Screenshot("Abrir Programa", true, file, desktopSession);
                                        //VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Observador, Observador, "COD_OBSR", c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        ////PASAR A ACCIONES
                                        var TPageControl = desktopSession.FindElementsByClassName("TPageControl");
                                        desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 45, 15);
                                        desktopSession.Mouse.Click(null);
                                        for (int i = 1; i <= 3; i++)
                                        {
                                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        }

                                        ////////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(50000);
                                        
                                        CrudKsomaobs.CRUDKSoMaobs(desktopSession, 1, crudVariables, file);
                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

                                        //Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Observador, Observador, "COD_OBSR", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        CrudKsomaobs.CRUDKSoMaobs(desktopSession, 2, crudVariables, file);
                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        //Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Observador, AccionToma, "ACI_COIN", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        CrudKsomaobs.CRUDKSoMaobs(desktopSession, 2, crudVariables, file);
                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Observador, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Observador, AccionTomaEditar, "COD_CONN", c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //PASAR A DETALLE 1
                                        desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 45, 10);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Mouse.Click(null);
                                        for (int i = 1; i <= 4; i++)
                                        {
                                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        }

                                        ////////AGREGAR DETALLE 1
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 1);
                                        var ElementList1 = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomaobs.CRUDDetalle1KSoMaobs(desktopSession, 1, crudVariablesdet1, file);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 1", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomaobs.CRUDDetalle1KSoMaobs(desktopSession, 2, crudVariablesdet1, file);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 1", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomaobs.CRUDDetalle1KSoMaobs(desktopSession, 2, crudVariablesdet1, file);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 1", true, file);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

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
                                        CrudKsomaobs.PreliKSoMaobs(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //RECTIFICACION DE CAMPOS DE EXCEL
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, Observador, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));
                                        newFunctions_1.lapiz(desktopSession);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR DETALLE 1
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[1], botonesNavegador, file);
                                        Thread.Sleep(1000);

                                        //ELIMINAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        Thread.Sleep(1000);

                                        //Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Observador, "", "COD_OBSR", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Observador, "D", c, ErrorValidacion);

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

        [TestMethod]
        public void KSoMasimCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_5.KSoMasimCheckList")
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
                                //Datos Login
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&

                                rows["Area"].ToString().Length != 0 && rows["Area"].ToString() != null &&
                                rows["CentroTrabajo"].ToString().Length != 0 && rows["CentroTrabajo"].ToString() != null &&
                                rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
                                rows["PoblacionObj"].ToString().Length != 0 && rows["PoblacionObj"].ToString() != null &&
                                rows["PoblFija"].ToString().Length != 0 && rows["PoblFija"].ToString() != null &&
                                rows["OtrosPoblF"].ToString().Length != 0 && rows["OtrosPoblF"].ToString() != null &&
                                rows["PoblFlotante"].ToString().Length != 0 && rows["PoblFlotante"].ToString() != null &&
                                rows["PersonasAusen"].ToString().Length != 0 && rows["PersonasAusen"].ToString() != null &&
                                rows["NPersonasEva"].ToString().Length != 0 && rows["NPersonasEva"].ToString() != null &&
                                rows["AnchoSalida"].ToString().Length != 0 && rows["AnchoSalida"].ToString() != null &&
                                rows["ConstExp"].ToString().Length != 0 && rows["ConstExp"].ToString() != null &&
                                rows["DistanciaTot"].ToString().Length != 0 && rows["DistanciaTot"].ToString() != null &&
                                rows["VeloDespl"].ToString().Length != 0 && rows["VeloDespl"].ToString() != null &&
                                rows["TiempoEjer"].ToString().Length != 0 && rows["TiempoEjer"].ToString() != null &&
                                rows["PoblObjEditar"].ToString().Length != 0 && rows["PoblObjEditar"].ToString() != null &&

                                rows["Complejidad"].ToString().Length != 0 && rows["Complejidad"].ToString() != null &&
                                rows["ComplejidadEditar"].ToString().Length != 0 && rows["ComplejidadEditar"].ToString() != null &&

                                rows["TipoInte"].ToString().Length != 0 && rows["TipoInte"].ToString() != null &&
                                rows["IdentiDet3"].ToString().Length != 0 && rows["IdentiDet3"].ToString() != null &&
                                rows["ReaccionAdec"].ToString().Length != 0 && rows["ReaccionAdec"].ToString() != null &&
                                rows["TipoInteEditar"].ToString().Length != 0 && rows["TipoInteEditar"].ToString() != null &&

                                rows["ComplejidadD2"].ToString().Length != 0 && rows["ComplejidadD2"].ToString() != null &&
                                rows["ComplejidadEditarD2"].ToString().Length != 0 && rows["ComplejidadEditarD2"].ToString() != null &&

                                rows["NombreAct"].ToString().Length != 0 && rows["NombreAct"].ToString() != null &&
                                rows["NombreActEditar"].ToString().Length != 0 && rows["NombreActEditar"].ToString() != null &&

                                rows["IdentiDet5"].ToString().Length != 0 && rows["IdentiDet5"].ToString() != null &&
                                rows["IdentiDet5Editar"].ToString().Length != 0 && rows["IdentiDet5Editar"].ToString() != null 
                                )
                            {
                                //VARIABLES GENERALES
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();
                                //VARIABLES CRUD
                                string Area = rows["Area"].ToString();
                                string CentroTrabajo = rows["CentroTrabajo"].ToString();
                                string ClaseSimu = rows["ClaseSimu"].ToString();
                                string Fecha = rows["Fecha"].ToString();
                                string PoblacionObj = rows["PoblacionObj"].ToString();
                                string PoblFija = rows["PoblFija"].ToString();
                                string OtrosPoblF = rows["OtrosPoblF"].ToString();
                                string PoblFlotante = rows["PoblFlotante"].ToString();
                                string PersonasAusen = rows["PersonasAusen"].ToString();
                                string NPersonasEva = rows["NPersonasEva"].ToString();
                                string AnchoSalida = rows["AnchoSalida"].ToString();
                                string ConstExp = rows["ConstExp"].ToString();
                                string DistanciaTot = rows["DistanciaTot"].ToString();
                                string VeloDespl = rows["VeloDespl"].ToString();
                                string TiempoEjer = rows["TiempoEjer"].ToString();
                                string PoblObjEditar = rows["PoblObjEditar"].ToString();
                                //VARIABLES DETALLE 4
                                string Complejidad = rows["Complejidad"].ToString();
                                string ComplejidadEditar = rows["ComplejidadEditar"].ToString();
                                //VARIABLES DETALLE 1
                                string TipoInte = rows["TipoInte"].ToString();
                                string IdentiDet3 = rows["IdentiDet3"].ToString();
                                string ReaccionAdec = rows["ReaccionAdec"].ToString();
                                string TipoInteEditar = rows["TipoInteEditar"].ToString();
                                //VARIABLES DETALLE 2
                                string ComplejidadD2 = rows["ComplejidadD2"].ToString();
                                string ComplejidadEditarD2 = rows["ComplejidadEditarD2"].ToString();
                                //VARIABLES DETALLE 3
                                string NombreAct = rows["NombreAct"].ToString();
                                string NombreActEditar = rows["NombreActEditar"].ToString();
                                //VARIABLES DETALLE 5
                                string IdentiDet5 = rows["IdentiDet5"].ToString();
                                string IdentiDet5Editar = rows["IdentiDet5Editar"].ToString();
                                //VALIDACIONES
                                string Tabla1 = rows["Tabla1"].ToString();
                                string Tabla2 = rows["Tabla2"].ToString();
                                string Tabla3 = rows["Tabla3"].ToString();
                                string Tabla4 = rows["Tabla4"].ToString();
                                string Tabla5 = rows["Tabla5"].ToString();
                                string Tabla6 = rows["Tabla6"].ToString();
                                string CampoDetalles = rows["CampoDetalles"].ToString();
                                string DatoDetalle = rows["DatoDetalle"].ToString();

                                List<string> crudVariables = new List<string>() { Area, CentroTrabajo, ClaseSimu, Fecha,
                                PoblacionObj, PoblFija, OtrosPoblF, PoblFlotante, PersonasAusen,
                                NPersonasEva, AnchoSalida, ConstExp, DistanciaTot, VeloDespl, TiempoEjer, PoblObjEditar };

                                List<string> crudVariablesdet4 = new List<string>() { Complejidad, ComplejidadEditar };

                                List<string> crudVariablesdet1 = new List<string>() { TipoInte, IdentiDet3, ReaccionAdec, TipoInteEditar };

                                List<string> crudVariablesdet2 = new List<string>() { ComplejidadD2, ComplejidadEditarD2 };

                                List<string> crudVariablesdet3 = new List<string>() { NombreAct, NombreActEditar };

                                List<string> crudVariablesdet5 = new List<string>() { IdentiDet5, IdentiDet5Editar };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);

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
                                        Screenshot("Abrir Programa", true, file, desktopSession);
                                        //VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 6
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla6, user, motor, CampoDetalles, DatoDetalle, DatoDetalle, CampoDetalles, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 5
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla5, user, motor, CampoDetalles, DatoDetalle, DatoDetalle, CampoDetalles, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 4
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla4, user, motor, CampoDetalles, DatoDetalle, DatoDetalle, CampoDetalles, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 3
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla3, user, motor, CampoDetalles, DatoDetalle, DatoDetalle, CampoDetalles, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, CampoDetalles, DatoDetalle, DatoDetalle, CampoDetalles, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, CampoDetalles, DatoDetalle, DatoDetalle, CampoDetalles, c, ErrorValidacion, "No se agregó correctamente", 0, "1");
                                        //Debugger.Launch();
                                        Thread.Sleep(1000);                                        
                                        //ELIMINAR REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Area, Area, "COD_AREA", c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        ////PASAR A DETALLE 4
                                        var TPageControl = desktopSession.FindElementsByClassName("TPageControl");
                                        desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 55, 15);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);

                                        //////////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDKSoMasim(desktopSession, 1, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);
                                        Thread.Sleep(2000);
                                        //Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Area, Area, "COD_AREA", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDKSoMasim(desktopSession, 2, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);
                                        Thread.Sleep(2000);
                                        //Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Area, PoblacionObj, "POB_OBSE", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDKSoMasim(desktopSession, 2, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsomasim.CRUDKSoMasim(desktopSession, 2, crudVariables, file);
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);
                                        Thread.Sleep(2000);
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Area, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Area, PoblObjEditar, "POB_OBSE", c, ErrorValidacion, "El registro no se editó correctamente", 1);
                                        Thread.Sleep(2000);
                                        //////////AGREGAR DETALLE 4
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 1);
                                        var ElementList4 = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle4KSoMasim(desktopSession, 1, crudVariablesdet4, file);
                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 4", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle4KSoMasim(desktopSession, 2, crudVariablesdet4, file);
                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 4", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle4KSoMasim(desktopSession, 2, crudVariablesdet4, file);
                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);

                                        Thread.Sleep(2000);
                                        valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsomasim.CRUDDetalle4KSoMasim(desktopSession, 2, crudVariablesdet4, file);
                                            desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 4", true, file);

                                        //PASAR A DETALLE 1
                                        desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 55, 15);
                                        desktopSession.Mouse.Click(null);

                                        ////////AGREGAR DETALLE 1
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 1);
                                        var ElementList1 = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle1KSoMasim(desktopSession, 1, crudVariablesdet1, file);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 1", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle1KSoMasim(desktopSession, 2, crudVariablesdet1, file);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 1", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle1KSoMasim(desktopSession, 2, crudVariablesdet1, file);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsomasim.CRUDDetalle1KSoMasim(desktopSession, 2, crudVariablesdet1, file);
                                            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 1", true, file);

                                        //PASAR A DETALLE 2 - 3
                                        desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 55, 15);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);

                                        ////////AGREGAR DETALLE 2
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 1);
                                        var ElementList2 = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle2KSoMasim(desktopSession, 1, crudVariablesdet2, file);
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 2", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle2KSoMasim(desktopSession, 2, crudVariablesdet2, file);
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 2", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle2KSoMasim(desktopSession, 2, crudVariablesdet2, file);
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 2", true, file);

                                        Thread.Sleep(2000);
                                        valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsomasim.CRUDDetalle2KSoMasim(desktopSession, 2, crudVariablesdet2, file);
                                            desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 2", true, file);

                                        ////////AGREGAR DETALLE 3
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 2);
                                        var ElementList3 = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle3KSoMasim(desktopSession, 1, crudVariablesdet3, file);
                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 3", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle3KSoMasim(desktopSession, 2, crudVariablesdet3, file);
                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 3", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle3KSoMasim(desktopSession, 2, crudVariablesdet3, file);
                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 3", true, file);

                                        Thread.Sleep(2000);
                                        valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsomasim.CRUDDetalle3KSoMasim(desktopSession, 2, crudVariablesdet3, file);
                                            desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 3", true, file);

                                        ////PASAR A DETALLE 5
                                        desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 55, 10);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);

                                        //////////AGREGAR DETALLE 5
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 1);
                                        var ElementList5 = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle5KSoMasim(desktopSession, 1, crudVariablesdet5, file);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 5", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle5KSoMasim(desktopSession, 2, crudVariablesdet5, file);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 5", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle5KSoMasim(desktopSession, 2, crudVariablesdet5, file);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 5", true, file);

                                        Thread.Sleep(2000);
                                        valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsomasim.CRUDDetalle5KSoMasim(desktopSession, 2, crudVariablesdet5, file);
                                            desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 5", true, file);

                                        //PASAR A DETALLE 6
                                        desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 40, 10);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Mouse.Click(null);
                                        for (int i = 1; i <= 6; i++)
                                        {
                                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        }
                                        TPageControl = desktopSession.FindElementsByClassName("TPageControl");
                                        desktopSession.Mouse.MouseMove(TPageControl[1].Coordinates, 25, 10);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);

                                        //////////AGREGAR DETALLE 6
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 1);
                                        var ElementList6 = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle5KSoMasim(desktopSession, 1, crudVariablesdet5, file);
                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 6", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle5KSoMasim(desktopSession, 2, crudVariablesdet5, file);
                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 6", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomasim.CRUDDetalle5KSoMasim(desktopSession, 2, crudVariablesdet5, file);
                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 6", true, file);

                                        Thread.Sleep(2000);
                                        valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsomasim.CRUDDetalle5KSoMasim(desktopSession, 2, crudVariablesdet5, file);
                                            desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 6", true, file);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

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
                                        CrudKsomasim.PreliKSoMasim(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //RECTIFICACION DE CAMPOS DE EXCEL
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, Area, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));
                                        newFunctions_1.lapiz(desktopSession);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        //ELIMINAR DETALLE 6
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList6[1], botonesNavegador, file);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Detalle 6", true, file);

                                        ////PASAR A DETALLE 5
                                        desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 40, 10);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        Thread.Sleep(1000);

                                        //ELIMINAR DETALLE 5
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList5[1], botonesNavegador, file);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Detalle 5", true, file);

                                        ////PASAR A DETALLE 4
                                        desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 55, 15);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);

                                        //ELIMINAR DETALLE 4
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList4[1], botonesNavegador, file);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Detalle 4", true, file);

                                        //PASAR A DETALLE 2-3
                                        desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 55, 15);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);

                                        //ELIMINAR DETALLE 3
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList3[2], botonesNavegador, file);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Detalle 3", true, file);

                                        //ELIMINAR DETALLE 2
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        var allFrame = desktopSession.FindElementsByClassName("IEFrame");
                                        new Actions(desktopSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) - 40, (allFrame[0].Size.Height / 2) + 40).DoubleClick().Click().Perform();
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Detalle 2", true, file);

                                        ////PASAR A DETALLE 1
                                        desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 55, 15);
                                        desktopSession.Mouse.Click(null);

                                        //ELIMINAR DETALLE 1
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[1], botonesNavegador, file);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Detalle 1", true, file);

                                        //ELIMINAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Registro", true, file);

                                        //Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Area, "", "COD_AREA", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Area, "D", c, ErrorValidacion);

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

        [TestMethod]
        public void KSoMatrlCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_5.KSoMatrlCheckList")
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
                                //Datos Login
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                
                                rows["FechaApro"].ToString().Length != 0 && rows["FechaApro"].ToString() != null &&
                                rows["FechaAproEditar"].ToString().Length != 0 && rows["FechaAproEditar"].ToString() != null &&

                                rows["NombreN"].ToString().Length != 0 && rows["NombreN"].ToString() != null &&
                                rows["NombreNEditar"].ToString().Length != 0 && rows["NombreNEditar"].ToString() != null &&

                                rows["DescriN"].ToString().Length != 0 && rows["DescriN"].ToString() != null &&
                                rows["DescriNEditar"].ToString().Length != 0 && rows["DescriNEditar"].ToString() != null &&

                                rows["DescriO"].ToString().Length != 0 && rows["DescriO"].ToString() != null &&
                                rows["DescriOEditar"].ToString().Length != 0 && rows["DescriOEditar"].ToString() != null &&

                                rows["EvaluaC"].ToString().Length != 0 && rows["EvaluaC"].ToString() != null &&
                                rows["EvaluaCEditar"].ToString().Length != 0 && rows["EvaluaCEditar"].ToString() != null &&

                                rows["VeriAct"].ToString().Length != 0 && rows["VeriAct"].ToString() != null &&
                                rows["VeriActEditar"].ToString().Length != 0 && rows["VeriActEditar"].ToString() != null
                                )
                            {
                                //VARIABLES GENERALES
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();
                                //VARIABLES CRUD
                                string FechaApro = rows["FechaApro"].ToString();
                                string FechaAproEditar = rows["FechaAproEditar"].ToString();
                                //VARIABLES DETALLE 1
                                string NombreN = rows["NombreN"].ToString();
                                string NombreNEditar = rows["NombreNEditar"].ToString();                                
                                //VARIABLES DETALLE 2
                                string DescriN = rows["DescriN"].ToString();
                                string DescriNEditar = rows["DescriNEditar"].ToString();                                
                                //VARIABLES DETALLE 3
                                string DescriO = rows["DescriO"].ToString();
                                string DescriOEditar = rows["DescriOEditar"].ToString();
                                //VARIABLES DETALLE 4
                                string EvaluaC = rows["EvaluaC"].ToString();
                                string EvaluaCEditar = rows["EvaluaCEditar"].ToString();
                                //VARIABLES DETALLE 5
                                string VeriAct = rows["VeriAct"].ToString();
                                string VeriActEditar = rows["VeriActEditar"].ToString();
                                string Dato = rows["Dato"].ToString();
                                string Tabla1 = rows["Tabla1"].ToString();
                                string Tabla2 = rows["Tabla2"].ToString();
                                string Tabla3 = rows["Tabla3"].ToString();
                                string Tabla4 = rows["Tabla4"].ToString();
                                string Tabla5 = rows["Tabla5"].ToString();

                                List<string> crudVariables = new List<string>() { FechaApro, FechaAproEditar };

                                List<string> crudVariablesdet1 = new List<string>() { NombreN, NombreNEditar };

                                List<string> crudVariablesdet2 = new List<string>() { DescriN, DescriNEditar };

                                List<string> crudVariablesdet3 = new List<string>() { DescriO, DescriOEditar };

                                List<string> crudVariablesdet4 = new List<string>() { EvaluaC, EvaluaCEditar };

                                List<string> crudVariablesdet5 = new List<string>() { VeriAct, VeriActEditar };


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);

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
                                        Screenshot("Abrir Programa", true, file, desktopSession);
                                        //VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        

                                        //ELIMINAR REGISTROS DETALLES 5 PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla5, user, motor, Campo, Dato, Dato, Campo, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS DETALLES 4 PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla4, user, motor, Campo, Dato, Dato, Campo, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS DETALLES 3 PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla3, user, motor, Campo, Dato, Dato, Campo, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS DETALLES 2 PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, Campo, Dato, Dato, Campo, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS DETALLES 1 PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo, Dato, Dato, Campo, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, "COD_EMPR", "9", "9", "COD_EMPR", c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //////////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDKSoMatrl(desktopSession, 1, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

                                        //Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, FechaApro, FechaApro, "FEC_APRO", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDKSoMatrl(desktopSession, 2, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        //Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, FechaApro, FechaApro, "FEC_APRO", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDKSoMatrl(desktopSession, 2, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, FechaAproEditar, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, FechaAproEditar, FechaAproEditar, "FEC_APRO", c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        ////////AGREGAR DETALLE 1
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 5);
                                        var ElementList1 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList1[5].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDDetalle1KSoMatrl(desktopSession, 1, crudVariablesdet1, file);
                                        desktopSession.Mouse.MouseMove(ElementList1[5].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 1", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList1[5].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDDetalle1KSoMatrl(desktopSession, 2, crudVariablesdet1, file);
                                        desktopSession.Mouse.MouseMove(ElementList1[5].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 1", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList1[5].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDDetalle1KSoMatrl(desktopSession, 2, crudVariablesdet1, file);
                                        desktopSession.Mouse.MouseMove(ElementList1[5].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 1", true, file);

                                        ////////AGREGAR DETALLE 2
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 4);
                                        var ElementList2 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList2[4].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDDetalle2KSoMatrl(desktopSession, 1, crudVariablesdet2, file);
                                        desktopSession.Mouse.MouseMove(ElementList2[4].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 2", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[4].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDDetalle2KSoMatrl(desktopSession, 2, crudVariablesdet2, file);
                                        desktopSession.Mouse.MouseMove(ElementList2[4].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 2", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[4].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDDetalle2KSoMatrl(desktopSession, 2, crudVariablesdet2, file);
                                        desktopSession.Mouse.MouseMove(ElementList2[4].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 2", true, file);

                                        ////////AGREGAR DETALLE 3
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 2);
                                        var ElementList3 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDDetalle3KSoMatrl(desktopSession, 1, crudVariablesdet3, file);
                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 3", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDDetalle3KSoMatrl(desktopSession, 2, crudVariablesdet3, file);
                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 3", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDDetalle3KSoMatrl(desktopSession, 2, crudVariablesdet3, file);
                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 3", true, file);

                                        //////////AGREGAR DETALLE 4
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 3);
                                        var ElementList4 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList4[3].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDDetalle4KSoMatrl(desktopSession, 1, crudVariablesdet4, file);
                                        desktopSession.Mouse.MouseMove(ElementList4[3].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 4", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList4[3].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDDetalle4KSoMatrl(desktopSession, 2, crudVariablesdet4, file);
                                        desktopSession.Mouse.MouseMove(ElementList4[3].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 4", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList4[3].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDDetalle4KSoMatrl(desktopSession, 2, crudVariablesdet4, file);
                                        desktopSession.Mouse.MouseMove(ElementList4[3].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 4", true, file);

                                        ////////AGREGAR DETALLE 5
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 1);
                                        var ElementList5 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDDetalle5KSoMatrl(desktopSession, 1, crudVariablesdet5, file);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 5", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDDetalle5KSoMatrl(desktopSession, 2, crudVariablesdet5, file);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 5", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsomatrl.CRUDDetalle5KSoMatrl(desktopSession, 2, crudVariablesdet5, file);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 5", true, file);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

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
                                        //RECTIFICACION DE CAMPOS DE EXCEL
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, FechaAproEditar, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));
                                        newFunctions_1.lapiz(desktopSession);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //ELIMINAR DETALLE 5
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList5[1], botonesNavegador, file);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Detalle 5", true, file);

                                        //ELIMINAR DETALLE 4
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList4[3], botonesNavegador, file);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Detalle 4", true, file);

                                        //ELIMINAR DETALLE 3
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList3[2], botonesNavegador, file);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Detalle 3", true, file);

                                        //ELIMINAR DETALLE 2
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList3[4], botonesNavegador, file);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Detalle 2", true, file);

                                        //ELIMINAR DETALLE 1
                                        desktopSession.Mouse.MouseMove(ElementList1[5].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[5], botonesNavegador, file);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Detalle 1", true, file);

                                        //ELIMINAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        var allFrame = desktopSession.FindElementsByClassName("IEFrame");
                                        new Actions(desktopSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) - 20, (allFrame[0].Size.Height / 2) + 80).DoubleClick().Click().Perform();
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Registro", true, file);

                                        //Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, FechaAproEditar, "", "FEC_APRO", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, FechaAproEditar, "D", c, ErrorValidacion);

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

        [TestMethod]
        public void KSoViepiCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_5.KSoViepiCheckList")
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
                                //Datos Login
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&

                                rows["Identificacion"].ToString().Length != 0 && rows["Identificacion"].ToString() != null &&
                                rows["IdentificacionEditar"].ToString().Length != 0 && rows["IdentificacionEditar"].ToString() != null &&

                                rows["FechaI"].ToString().Length != 0 && rows["FechaI"].ToString() != null &&
                                rows["Seguimiento"].ToString().Length != 0 && rows["Seguimiento"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["Estado"].ToString().Length != 0 && rows["Estado"].ToString() != null &&
                                rows["ObservacionEditar"].ToString().Length != 0 && rows["ObservacionEditar"].ToString() != null &&

                                rows["FechaPreli"].ToString().Length != 0 && rows["FechaPreli"].ToString() != null
                                )
                            {
                                //VARIABLES GENERALES
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();
                                //VARIABLES CRUD
                                string Identificacion = rows["Identificacion"].ToString();
                                string IdentificacionEditar = rows["IdentificacionEditar"].ToString();
                                //VARIABLES DETALLE 1
                                string FechaI = rows["FechaI"].ToString();
                                string Seguimiento = rows["Seguimiento"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string Estado = rows["Estado"].ToString();
                                string ObservacionEditar = rows["ObservacionEditar"].ToString();
                                //VARIABLES PRELIMINAR
                                string FechaPreli = rows["FechaPreli"].ToString();

                                List<string> crudVariables = new List<string>() { Identificacion, IdentificacionEditar };

                                List<string> crudVariablesdet1 = new List<string>() { FechaI, Seguimiento, Observacion, Estado, ObservacionEditar };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);

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
                                        Screenshot("Abrir Programa", true, file, desktopSession);
                                        //VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINA REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Identificacion, Identificacion, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        ////////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoviepi.CRUDKSoViepi(desktopSession, 1, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

                                        //Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Identificacion, Identificacion, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoviepi.CRUDKSoViepi(desktopSession, 2, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        //Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Identificacion, Identificacion, "COD_EMPL", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoviepi.CRUDKSoViepi(desktopSession, 2, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, QbeFilter, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, IdentificacionEditar, IdentificacionEditar, "COD_EMPL", c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        ////////AGREGAR DETALLE 1
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                        var ElementList1 = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoviepi.CRUDDetalle1KSoViepi(desktopSession, 1, crudVariablesdet1, file);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 1", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoviepi.CRUDDetalle1KSoViepi(desktopSession, 2, crudVariablesdet1, file);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 1", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoviepi.CRUDDetalle1KSoViepi(desktopSession, 2, crudVariablesdet1, file);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 1", true, file);
                                        
                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

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
                                        CrudKsoviepi.PreliKSoViepi(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli, FechaPreli);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //RECTIFICACION DE CAMPOS DE EXCEL
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, IdentificacionEditar, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        //Celda = Celda_Excel(ruta, nom_empr);
                                        //Celda.ForEach(u => ErrorValidacion.Add(u));
                                        newFunctions_1.lapiz(desktopSession);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR DETALLE 1
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[1], botonesNavegador, file);
                                        Thread.Sleep(1000);

                                        //ELIMINAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        Thread.Sleep(1000);

                                        //Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, IdentificacionEditar, "", "COD_EMPL", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, IdentificacionEditar, "D", c, ErrorValidacion);

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

    }
}

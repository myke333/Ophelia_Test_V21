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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosPn;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_9;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_11;
using OpenQA.Selenium.Appium;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosEd;

namespace Ophelia_Test_V21.TestMethods.ModulosNM
{
    /// <summary>
    /// Descripción resumida de ModulosNM_68
    /// </summary>
    [TestClass]
    public class ModulosNM_131 : FuncionesVitales
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
        public ModulosNM_131()
        {
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
        // Encolelo ahi en el test configuration q se desencolo
        #endregion
        
        [TestMethod]
        public void AbrirPrograma()
        {

            string codProgram = "kgepbsca";
            string user = "UNatalia";
            string motor = "SQL";

            AbrirPrograma a = new AbrirPrograma(codProgram, user);
            desktopSession = a.Start2(motor, "");


        }

        [TestMethod]
        public void KNmHmovcChecklist()
        {

            List<string> errorMessages = new List<string>();
            List<string> listaResultado;
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_131.KNmHmovcChecklist")
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
                                rows["BanderaSuma"].ToString().Length != 0 && rows["BanderaSuma"].ToString() != null &&


                                //Variables Validaciones Crud Principal
                                rows["FechaIniContr"].ToString().Length != 0 && rows["FechaIniContr"].ToString() != null &&
                                rows["FechaFinContr"].ToString().Length != 0 && rows["FechaFinContr"].ToString() != null &&
                                rows["ObjContr"].ToString().Length != 0 && rows["ObjContr"].ToString() != null &&
                                rows["ActObjContr"].ToString().Length != 0 && rows["ActObjContr"].ToString() != null &&

                                //Variables Validaciones detalle 1
                                rows["ValorTotal"].ToString().Length != 0 && rows["ValorTotal"].ToString() != null &&
                                rows["nPerio"].ToString().Length != 0 && rows["nPerio"].ToString() != null &&
                                rows["ValorMen"].ToString().Length != 0 && rows["ValorMen"].ToString() != null &&
                                rows["ActValorMen"].ToString().Length != 0 && rows["ActValorMen"].ToString() != null &&

                                //Variables Validaciones detalle 2
                                rows["fechaFinal"].ToString().Length != 0 && rows["fechaFinal"].ToString() != null &&
                                rows["ActFechaFinal"].ToString().Length != 0 && rows["ActFechaFinal"].ToString() != null &&
                                rows["cesionario"].ToString().Length != 0 && rows["cesionario"].ToString() != null &&

                                //Variables Validaciones detalle 3
                                rows["Concepto"].ToString().Length != 0 && rows["Concepto"].ToString() != null &&
                                rows["fechaInicial"].ToString().Length != 0 && rows["fechaInicial"].ToString() != null &&
                                rows["fechaFin"].ToString().Length != 0 && rows["fechaFin"].ToString() != null &&
                                rows["fechaReinicio"].ToString().Length != 0 && rows["fechaReinicio"].ToString() != null &&
                                rows["NroDias"].ToString().Length != 0 && rows["NroDias"].ToString() != null &&
                                rows["ActNroDias"].ToString().Length != 0 && rows["ActNroDias"].ToString() != null &&
                                rows["motivo"].ToString().Length != 0 && rows["motivo"].ToString() != null &&

                                //Variables Validaciones
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null)


                                ////Datos Cruds
                                //rows["Codigo"].ToString().Length != 0 && rows["Codigo"].ToString() != null &&
                                //rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
                                //rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
                                //rows["ActNombre"].ToString().Length != 0 && rows["ActNombre"].ToString() != null &&

                                //rows["Cargo"].ToString().Length != 0 && rows["Cargo"].ToString() != null &&
                                //rows["Dependencia"].ToString().Length != 0 && rows["Dependencia"].ToString() != null &&
                                //rows["Horas"].ToString().Length != 0 && rows["Horas"].ToString() != null &&
                                //rows["ActHoras"].ToString().Length != 0 && rows["ActHoras"].ToString() != null &&

                                //rows["CargoDet2"].ToString().Length != 0 && rows["CargoDet2"].ToString() != null &&
                                //rows["HorasMin"].ToString().Length != 0 && rows["HorasMin"].ToString() != null &&
                                //rows["HorasMax"].ToString().Length != 0 && rows["HorasMax"].ToString() != null &&
                                //rows["ActHorasMax"].ToString().Length != 0 && rows["ActHorasMax"].ToString() != null &&

                                //rows["CargoDet3"].ToString().Length != 0 && rows["CargoDet3"].ToString() != null &&
                                //rows["DependenciaDet3"].ToString().Length != 0 && rows["DependenciaDet3"].ToString() != null &&
                                //rows["ActDependenciaDet3"].ToString().Length != 0 && rows["ActDependenciaDet3"].ToString() != null)

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
                                string BanderaSuma = rows["BanderaSuma"].ToString();

                                //Variables Crud principal
                                string FechaIniContr = rows["FechaIniContr"].ToString();
                                string FechaFinContr = rows["FechaFinContr"].ToString();
                                string ObjContr = rows["ObjContr"].ToString();
                                string ActObjContr = rows["ActObjContr"].ToString();
                                
                                //Variables validacion
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string EditCampo = rows["EditCampo"].ToString();

                                //Variables validacion DETALLE 1
                                string ValorTotal = rows["ValorTotal"].ToString();
                                string nPerio = rows["nPerio"].ToString();
                                string ValorMen = rows["ValorMen"].ToString();
                                string ActValorMen = rows["ActValorMen"].ToString();

                                //Variables validacion DETALLE 2
                                string fechaFinal = rows["fechaFinal"].ToString();
                                string ActFechaFinal = rows["ActFechaFinal"].ToString();
                                string cesionario = rows["cesionario"].ToString();


                                //Variables validacion DETALLE 3
                                string Concepto = rows["Concepto"].ToString();
                                string fechaInicial = rows["fechaInicial"].ToString();
                                string fechaFin = rows["fechaFin"].ToString();
                                string fechaReinicio = rows["fechaReinicio"].ToString(); 
                                string NroDias = rows["NroDias"].ToString();
                                string ActNroDias = rows["ActNroDias"].ToString();
                                string motivo = rows["motivo"].ToString();
                                ////Variables crud principal
                                //string Codigo = rows["Codigo"].ToString();
                                //string Nombre = rows["Nombre"].ToString();
                                //string Fecha = rows["Fecha"].ToString();
                                //string ActNombre = rows["ActNombre"].ToString();

                                ////Variables Crud Det 1
                                //string Cargo = rows["Cargo"].ToString();
                                //string Dependencia = rows["Dependencia"].ToString();
                                //string Horas = rows["Horas"].ToString();
                                //string ActHoras = rows["ActHoras"].ToString();

                                ////Variables Crud Det 2
                                //string CargoDet2 = rows["CargoDet2"].ToString();
                                //string HorasMin = rows["HorasMin"].ToString();
                                //string HorasMax = rows["HorasMax"].ToString();
                                //string ActHorasMax = rows["ActHorasMax"].ToString();

                                ////Variables Crud Det 3
                                //string CargoDet3 = rows["CargoDet3"].ToString();
                                //string DependenciaDet3 = rows["DependenciaDet3"].ToString();
                                //string ActDependenciaDet3 = rows["ActDependenciaDet3"].ToString();

                                //string Tabla = rows["Tabla"].ToString();
                                //string Campo = rows["Campo"].ToString();


                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { FechaIniContr, FechaFinContr, ObjContr, ActObjContr };
                                //LISTA DETALLE 1
                                List<string> crudDet1 = new List<string>() { ValorTotal, nPerio, ValorMen, ActValorMen };
                                //LISTA DETALLE 2
                                List<string> crudDet2 = new List<string>() { fechaFinal, ActFechaFinal, cesionario };
                                //LISTA DETALLE 3
                                List<string> crudDet3 = new List<string>() { Concepto, fechaInicial, fechaFin, fechaReinicio,  NroDias, ActNroDias, motivo};

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
                                        //      VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        //      VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        //  VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);


                                        //  NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);




                                        //   AGREGAR REGISTRO Principal
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKnmhmovc.AgregarCrud(desktopSession, 0, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);
                                        Thread.Sleep(1500);

                                        //              LUPA 
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }



                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmhmovc.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);
                                        Thread.Sleep(1000);



                                        //              LUPA 
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmhmovc.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);
                                        Thread.Sleep(3000);



                                        // LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }



                                        ////AGREGAR REGISTRO DETALLE 1                                        
                                        Dictionary<string, Point> botonesNavegador2 = new Dictionary<string, Point>();
                                        botonesNavegador2 = PruebaCRUD.findname(desktopSession, 28, 1);
                                        var ElementList2 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Insertar"].X, botonesNavegador2["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKnmhmovc.AgregarCrud1(desktopSession, 0, crudDet1);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 1", true, file);
                                        Thread.Sleep(1500);


                                        //              LUPA 
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }



                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                        CrudKnmhmovc.AgregarCrud1(desktopSession, 1, crudDet1);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Cancelar"].X, botonesNavegador2["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 1", true, file);
                                        Thread.Sleep(1000);


                                        //              LUPA 
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmhmovc.AgregarCrud1(desktopSession, 1, crudDet1);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);
                                        Thread.Sleep(3000);

                                        CrudKnmhmovc.ClickButton(desktopSession, 0);
                                        Thread.Sleep(2000);



                                        //              LUPA 
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////AGREGAR REGISTRO DETALLE 2

                                        Dictionary<string, Point> botonesNavegador3 = new Dictionary<string, Point>();
                                        botonesNavegador3 = PruebaCRUD.findname(desktopSession, 25, 1);
                                        var ElementList3 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Insertar"].X, botonesNavegador3["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKnmhmovc.AgregarCrud2(desktopSession, 0, crudDet2);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos DETALLE 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Aplicar"].X, botonesNavegador3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados DETALLE 2", true, file);
                                        Thread.Sleep(1500);

                                        //              LUPA 
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Editar"].X, botonesNavegador3["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar DETALLE 2", true, file);

                                        CrudKnmhmovc.AgregarCrud2(desktopSession, 1, crudDet2);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados DETALLE 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Cancelar"].X, botonesNavegador3["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados DETALLE 2", true, file);
                                        Thread.Sleep(2000);


                                        //              LUPA 
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Editar"].X, botonesNavegador3["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar DETALLE 2", true, file);

                                        CrudKnmhmovc.AgregarCrud2(desktopSession, 1, crudDet2);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados DETALLE 2", true, file);

                                        Thread.Sleep(2000);
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Aplicar"].X, botonesNavegador3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);


                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada DETALLE 2", true, file);
                                        Thread.Sleep(2000);

                                        CrudKnmhmovc.ClickButton(desktopSession, 1);
                                        Thread.Sleep(2000);

                                        //              LUPA 
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////AGREGAR REGISTRO DETALLE 3

                                        Dictionary<string, Point> botonesNavegador4 = new Dictionary<string, Point>();
                                        botonesNavegador4 = PruebaCRUD.findname(desktopSession, 26, 1);
                                        var ElementList4 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Insertar"].X, botonesNavegador4["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKnmhmovc.AgregarCrud3(desktopSession, 0, crudDet3);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos DETALLE 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Aplicar"].X, botonesNavegador4["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados DETALLE 3", true, file);
                                        Thread.Sleep(1500);


                                        //////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);


                                        //              LUPA 
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Editar"].X, botonesNavegador4["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar DETALLE 3", true, file);

                                        CrudKnmhmovc.AgregarCrud3(desktopSession, 1, crudDet3);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados DETALLE 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Cancelar"].X, botonesNavegador4["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados DETALLE 3", true, file);
                                        Thread.Sleep(1000);


                                        //////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        //              LUPA 
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Editar"].X, botonesNavegador4["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar DETALLE 3", true, file);

                                        CrudKnmhmovc.AgregarCrud3(desktopSession, 1, crudDet3);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados DETALLE 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Aplicar"].X, botonesNavegador4["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada DETALLE 3", true, file);
                                        Thread.Sleep(3000);


                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSuma, file);
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
                                        Thread.Sleep(2000);


                                        CrudKnmhmovc.ClickButton(desktopSession, 1);


                                        //Eliminar Registro Detalle 3
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList4[1], botonesNavegador4, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 3", true, file);
                                        Thread.Sleep(2000);


                                        CrudKnmhmovc.ClickButton(desktopSession, 0);
                                        Thread.Sleep(2000);

                                        //Eliminar Registro Detalle 2
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList3[1], botonesNavegador3, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 2", true, file);


                                        CrudKnmhmovc.ClickButton(desktopSession, 2);
                                        Thread.Sleep(2000);


                                        //Eliminar Registro Detalle 1
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList2[1], botonesNavegador2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 1", true, file);


                                        //Eliminar Registro CRUD Principal
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Principal", true, file);


                                        ////VALIDACIÓN ELIMINAR REGISTRO
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "D", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);


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
    }
}
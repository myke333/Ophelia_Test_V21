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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_9;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_10;
using OpenQA.Selenium.Appium;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosEd;

namespace Ophelia_Test_V21.TestMethods.ModulosNM
{
    [TestClass]
    public class ModulosNM_97 : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        protected static WindowsDriver<WindowsElement> desktopSession;

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
        public ModulosNM_97()
        { }

        private TestContext testContextInstance;

        public TestContext TestContext
        {
            get
            { return testContextInstance; }
            set
            { testContextInstance = value; }
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
        public void KNmNeparChecKlist()
        {
            //Debugger.Launch();
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_97.KNmNeparChecKlist")
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

                                //Variables Validaciones Crud principal
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&

                                //Datos Crud Principal
                                rows["Razon"].ToString().Length != 0 && rows["Razon"].ToString() != null &&
                                rows["NIT"].ToString().Length != 0 && rows["NIT"].ToString() != null &&
                                rows["Digito"].ToString().Length != 0 && rows["Digito"].ToString() != null &&
                                rows["Pais"].ToString().Length != 0 && rows["Pais"].ToString() != null &&
                                rows["Dpto"].ToString().Length != 0 && rows["Dpto"].ToString() != null &&
                                rows["Muni"].ToString().Length != 0 && rows["Muni"].ToString() != null &&
                                rows["Locali"].ToString().Length != 0 && rows["Locali"].ToString() != null &&
                                rows["Direccion"].ToString().Length != 0 && rows["Direccion"].ToString() != null &&
                                rows["Apelli"].ToString().Length != 0 && rows["Apelli"].ToString() != null &&
                                rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
                                rows["SuelBa"].ToString().Length != 0 && rows["SuelBa"].ToString() != null &&
                                rows["AuTra"].ToString().Length != 0 && rows["AuTra"].ToString() != null &&
                                rows["HoExDiu"].ToString().Length != 0 && rows["HoExDiu"].ToString() != null &&
                                rows["HoExNoc"].ToString().Length != 0 && rows["HoExNoc"].ToString() != null &&
                                rows["HoExDDF"].ToString().Length != 0 && rows["HoExDDF"].ToString() != null &&
                                rows["HoReDDF"].ToString().Length != 0 && rows["HoReDDF"].ToString() != null &&
                                rows["HoExNDF"].ToString().Length != 0 && rows["HoExNDF"].ToString() != null &&
                                rows["HoReNDF"].ToString().Length != 0 && rows["HoReNDF"].ToString() != null &&
                                rows["VacaTiem"].ToString().Length != 0 && rows["VacaTiem"].ToString() != null &&
                                rows["Salud"].ToString().Length != 0 && rows["Salud"].ToString() != null &&
                                rows["Pension"].ToString().Length != 0 && rows["Pension"].ToString() != null &&
                                rows["FonSoli"].ToString().Length != 0 && rows["FonSoli"].ToString() != null &&
                                rows["FonSub"].ToString().Length != 0 && rows["FonSub"].ToString() != null &&
                                rows["SanPriv"].ToString().Length != 0 && rows["SanPriv"].ToString() != null &&
                                rows["Libranza"].ToString().Length != 0 && rows["Libranza"].ToString() != null &&
                                rows["ApoPenVo"].ToString().Length != 0 && rows["ApoPenVo"].ToString() != null &&
                                rows["ReteFuen"].ToString().Length != 0 && rows["ReteFuen"].ToString() != null &&
                                rows["CodVar"].ToString().Length != 0 && rows["CodVar"].ToString() != null &&
                                rows["LicNoRemu"].ToString().Length != 0 && rows["LicNoRemu"].ToString() != null &&
                                rows["ActNombre"].ToString().Length != 0 && rows["ActNombre"].ToString() != null)

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

                                //Variables validacion  crud Principal
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string EditCampo = rows["EditCampo"].ToString();

                                //Variables crud principal
                                string Razon = rows["Razon"].ToString();
                                string NIT = rows["NIT"].ToString();
                                string Digito = rows["Digito"].ToString();
                                string Pais = rows["Pais"].ToString();
                                string Dpto = rows["Dpto"].ToString();
                                string Muni = rows["Muni"].ToString();
                                string Locali = rows["Locali"].ToString();
                                string Direccion = rows["Direccion"].ToString();
                                string Apelli = rows["Apelli"].ToString();
                                string Nombre = rows["Nombre"].ToString();
                                string SuelBa = rows["SuelBa"].ToString();
                                string AuTra = rows["AuTra"].ToString();
                                string HoExDiu = rows["HoExDiu"].ToString();
                                string HoExNoc = rows["HoExNoc"].ToString();
                                string HoExDDF = rows["HoExDDF"].ToString();
                                string HoReDDF = rows["HoReDDF"].ToString();
                                string HoExNDF = rows["HoExNDF"].ToString();
                                string HoReNDF = rows["HoReNDF"].ToString();
                                string VacaTiem = rows["VacaTiem"].ToString();
                                string Salud = rows["Salud"].ToString();
                                string Pension = rows["Pension"].ToString();
                                string FonSoli = rows["FonSoli"].ToString();
                                string FonSub = rows["FonSub"].ToString();
                                string SanPriv = rows["SanPriv"].ToString();
                                string Libranza = rows["Libranza"].ToString();
                                string ApoPenVo = rows["ApoPenVo"].ToString();
                                string ReteFuen = rows["ReteFuen"].ToString();
                                string CodVar = rows["CodVar"].ToString();
                                string LicNoRemu = rows["LicNoRemu"].ToString();
                                string ActNombre = rows["ActNombre"].ToString();

                                String campo = "1";

                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { Razon, NIT, Digito, Pais, Dpto, Muni, Locali, Direccion,
                                                                                  Apelli, Nombre, SuelBa, AuTra, HoExDiu, HoExNoc, HoExDDF,
                                                                                  HoReDDF, HoExNDF, HoReNDF, VacaTiem, Salud, Pension, 
                                                                                  FonSoli, FonSub, SanPriv, Libranza, ApoPenVo, ReteFuen, ActNombre, 
                                                                                  CodVar, LicNoRemu};

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

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, Campo, c, ErrorValidacion, "No se agregó correctamente", 0, "1");
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDet, user, motor, CampoDet, Codigo, Codigo, CampoDet, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //MODIFICAMOS
                                        //CAMBIAR EMPRESA
                                        CrudKnmnepar.Click(desktopSession, 0);

                                        //HACEMOS QBE
                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_1.lapiz(desktopSession);                                                                           

                                        
                                        Thread.Sleep(2000);

                                        //ELIMINAMOS REGISTRO Principal
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        ////LUPA
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(2000);

                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Principal", true, file);
                                        //VALIDACIÓN ELIMINAR REGISTRO
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "D", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);


                                        //AGREGAMOS EL REGISTRO                                        
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKnmnepar.AgregarCrud(desktopSession, 0, crudPrincipal, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados pestana Generalidades", true, file);

                                        Thread.Sleep(1500);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnepar.AgregarCrud(desktopSession, 1, crudPrincipal, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Cargo, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnepar.AgregarCrud(desktopSession, 1, crudPrincipal, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        Thread.Sleep(3000);
                                        ////validacion error al editar
                                        bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        //Debugger.Launch();
                                        if (valEditOra == true)
                                        {
                                            ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            newFunctions_3.LupaAud(desktopSession, "1", file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                            CrudKnmnepar.AgregarCrud(desktopSession, 1, crudPrincipal, file);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        //////Validacion Editar
                                        Thread.Sleep(1000);
                                        //// //Debugger.Launch();
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, ActCargo, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        //Debugger.Launch();
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle1, motor, user, CampoDetalle1, Codigo, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, ActCantidad, EditCampoDetalle1, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
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
                                        Thread.Sleep(2000);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        
                                        Thread.Sleep(2000);

                                                                          

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
        public void KNmNoecoChecKlist()
        {
            //Debugger.Launch();
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_97.KNmNoecoChecKlist")
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

                                //Variables Validaciones Crud principal
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&

                                //Variables Validaciones Crud Detalle
                                rows["TablaDet"].ToString().Length != 0 && rows["TablaDet"].ToString() != null &&
                                rows["CampoDet"].ToString().Length != 0 && rows["CampoDet"].ToString() != null &&
                                rows["EditCampoDet"].ToString().Length != 0 && rows["EditCampoDet"].ToString() != null &&

                                //Variables Validaciones Crud Detalle 2
                                rows["TablaDet2"].ToString().Length != 0 && rows["TablaDet2"].ToString() != null &&
                                rows["CampoDet2"].ToString().Length != 0 && rows["CampoDet2"].ToString() != null &&
                                rows["EditCampoDet2"].ToString().Length != 0 && rows["EditCampoDet2"].ToString() != null &&

                                //Variables Validaciones Crud Detalle 3
                                rows["TablaDet3"].ToString().Length != 0 && rows["TablaDet3"].ToString() != null &&
                                rows["CampoDet3"].ToString().Length != 0 && rows["CampoDet3"].ToString() != null &&
                                rows["EditCampoDet3"].ToString().Length != 0 && rows["EditCampoDet3"].ToString() != null &&

                                //Variables Validaciones Crud Detalle 4
                                rows["TablaDet4"].ToString().Length != 0 && rows["TablaDet4"].ToString() != null &&
                                rows["CampoDet4"].ToString().Length != 0 && rows["CampoDet4"].ToString() != null &&
                                rows["EditCampoDet4"].ToString().Length != 0 && rows["EditCampoDet4"].ToString() != null &&

                                //Variables Validaciones Crud Detalle 5
                                rows["TablaDet5"].ToString().Length != 0 && rows["TablaDet5"].ToString() != null &&
                                rows["CampoDet5"].ToString().Length != 0 && rows["CampoDet5"].ToString() != null &&
                                rows["EditCampoDet5"].ToString().Length != 0 && rows["EditCampoDet5"].ToString() != null &&

                                //Variables Validaciones Crud Detalle 6
                                rows["TablaDet6"].ToString().Length != 0 && rows["TablaDet6"].ToString() != null &&
                                rows["CampoDet6"].ToString().Length != 0 && rows["CampoDet6"].ToString() != null &&
                                rows["EditCampoDet6"].ToString().Length != 0 && rows["EditCampoDet6"].ToString() != null &&

                                //Variables Validaciones Crud Detalle 7
                                rows["TablaDet7"].ToString().Length != 0 && rows["TablaDet7"].ToString() != null &&
                                rows["CampoDet7"].ToString().Length != 0 && rows["CampoDet7"].ToString() != null &&
                                rows["EditCampoDet7"].ToString().Length != 0 && rows["EditCampoDet7"].ToString() != null &&

                                //Variables Validaciones Crud Detalle 8
                                rows["TablaDet8"].ToString().Length != 0 && rows["TablaDet8"].ToString() != null &&
                                rows["CampoDet8"].ToString().Length != 0 && rows["CampoDet8"].ToString() != null &&
                                rows["EditCampoDet8"].ToString().Length != 0 && rows["EditCampoDet8"].ToString() != null &&

                                //Datos Crud Principal
                                rows["CodOpe"].ToString().Length != 0 && rows["CodOpe"].ToString() != null &&
                                rows["NomOpe"].ToString().Length != 0 && rows["NomOpe"].ToString() != null &&
                                rows["NIT"].ToString().Length != 0 && rows["NIT"].ToString() != null &&
                                rows["ActNombre"].ToString().Length != 0 && rows["ActNombre"].ToString() != null &&

                                //Datos Crud Detalle
                                rows["CodNom"].ToString().Length != 0 && rows["CodNom"].ToString() != null &&
                                rows["DescriNom"].ToString().Length != 0 && rows["DescriNom"].ToString() != null &&
                                rows["ActDescri"].ToString().Length != 0 && rows["ActDescri"].ToString() != null &&

                                //Datos Crud Detalle 2
                                rows["CodPeHom"].ToString().Length != 0 && rows["CodPeHom"].ToString() != null &&
                                rows["NomPeHom"].ToString().Length != 0 && rows["NomPeHom"].ToString() != null &&
                                rows["CodPeKac"].ToString().Length != 0 && rows["CodPeKac"].ToString() != null &&
                                rows["NomPeHom1"].ToString().Length != 0 && rows["NomPeHom1"].ToString() != null &&
                                rows["ActNomPeHom1"].ToString().Length != 0 && rows["ActNomPeHom1"].ToString() != null &&

                                //Datos Crud Detalle 3
                                rows["CodTip"].ToString().Length != 0 && rows["CodTip"].ToString() != null &&
                                rows["NomTip"].ToString().Length != 0 && rows["NomTip"].ToString() != null &&
                                rows["ActNomTip"].ToString().Length != 0 && rows["ActNomTip"].ToString() != null &&

                                //Datos Crud Detalle 4
                                rows["CodConceKac"].ToString().Length != 0 && rows["CodConceKac"].ToString() != null &&
                                rows["CodConceHom"].ToString().Length != 0 && rows["CodConceHom"].ToString() != null &&
                                rows["NomConceHom"].ToString().Length != 0 && rows["NomConceHom"].ToString() != null &&
                                rows["Porce"].ToString().Length != 0 && rows["Porce"].ToString() != null &&
                                rows["CodConceKac1"].ToString().Length != 0 && rows["CodConceKac1"].ToString() != null &&
                                rows["CodConceHom1"].ToString().Length != 0 && rows["CodConceHom1"].ToString() != null &&
                                rows["NomConceHom1"].ToString().Length != 0 && rows["NomConceHom1"].ToString() != null &&
                                rows["Porce1"].ToString().Length != 0 && rows["Porce1"].ToString() != null &&
                                rows["CodConceKac2"].ToString().Length != 0 && rows["CodConceKac2"].ToString() != null &&
                                rows["CodConceHom2"].ToString().Length != 0 && rows["CodConceHom2"].ToString() != null &&
                                rows["NomConceHom2"].ToString().Length != 0 && rows["NomConceHom2"].ToString() != null &&
                                rows["Porce2"].ToString().Length != 0 && rows["Porce2"].ToString() != null &&
                                rows["CodConceKac3"].ToString().Length != 0 && rows["CodConceKac3"].ToString() != null &&
                                rows["CodConceHom3"].ToString().Length != 0 && rows["CodConceHom3"].ToString() != null &&
                                rows["NomConceHom3"].ToString().Length != 0 && rows["NomConceHom3"].ToString() != null &&
                                rows["CodConceKac4"].ToString().Length != 0 && rows["CodConceKac4"].ToString() != null &&
                                rows["CodConceHom4"].ToString().Length != 0 && rows["CodConceHom4"].ToString() != null &&
                                rows["NomConceHom4"].ToString().Length != 0 && rows["NomConceHom4"].ToString() != null &&
                                rows["Porce4"].ToString().Length != 0 && rows["Porce4"].ToString() != null &&
                                rows["CodConceKac5"].ToString().Length != 0 && rows["CodConceKac5"].ToString() != null &&
                                rows["CodConceHom5"].ToString().Length != 0 && rows["CodConceHom5"].ToString() != null &&
                                rows["NomConceHom5"].ToString().Length != 0 && rows["NomConceHom5"].ToString() != null &&
                                rows["Porce5"].ToString().Length != 0 && rows["Porce5"].ToString() != null &&
                                rows["ActNomConceHom5"].ToString().Length != 0 && rows["ActNomConceHom5"].ToString() != null &&

                                //Datos Crud Detalle 5
                                rows["CodHom"].ToString().Length != 0 && rows["CodHom"].ToString() != null &&

                                //Datos Crud Detalle 7
                                rows["CodInca"].ToString().Length != 0 && rows["CodInca"].ToString() != null &&
                                rows["NomInca"].ToString().Length != 0 && rows["NomInca"].ToString() != null &&
                                rows["ActNomInca"].ToString().Length != 0 && rows["ActNomInca"].ToString() != null &&

                                //Datos Crud Detalle 8
                                rows["CodExclu"].ToString().Length != 0 && rows["CodExclu"].ToString() != null &&
                                rows["ActCodExclu"].ToString().Length != 0 && rows["ActCodExclu"].ToString() != null &&

                                //Datos Crud Detalle 6
                                rows["CodHom1"].ToString().Length != 0 && rows["CodHom1"].ToString() != null)

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

                                //Variables validacion  crud Principal
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string EditCampo = rows["EditCampo"].ToString();

                                //Variables validacion  crud Detalle
                                string TablaDet = rows["TablaDet"].ToString();
                                string CampoDet = rows["CampoDet"].ToString();
                                string EditCampoDet = rows["EditCampoDet"].ToString();

                                //Variables validacion  crud Detalle 2
                                string TablaDet2 = rows["TablaDet2"].ToString();
                                string CampoDet2 = rows["CampoDet2"].ToString();
                                string EditCampoDet2 = rows["EditCampoDet2"].ToString();

                                //Variables validacion  crud Detalle 3
                                string TablaDet3 = rows["TablaDet3"].ToString();
                                string CampoDet3 = rows["CampoDet3"].ToString();
                                string EditCampoDet3 = rows["EditCampoDet3"].ToString();

                                //Variables validacion  crud Detalle 4
                                string TablaDet4 = rows["TablaDet4"].ToString();
                                string CampoDet4 = rows["CampoDet4"].ToString();
                                string EditCampoDet4 = rows["EditCampoDet4"].ToString();

                                //Variables validacion  crud Detalle 5
                                string TablaDet5 = rows["TablaDet5"].ToString();
                                string CampoDet5 = rows["CampoDet5"].ToString();
                                string EditCampoDet5 = rows["EditCampoDet5"].ToString();

                                //Variables validacion  crud Detalle 6
                                string TablaDet6 = rows["TablaDet6"].ToString();
                                string CampoDet6 = rows["CampoDet6"].ToString();
                                string EditCampoDet6 = rows["EditCampoDet6"].ToString();

                                //Variables crud principal
                                string CodOpe = rows["CodOpe"].ToString();
                                string NomOpe = rows["NomOpe"].ToString();
                                string NIT = rows["NIT"].ToString();
                                string ActNombre = rows["ActNombre"].ToString();

                                //Variables crud Detalle
                                string CodNom = rows["CodNom"].ToString();
                                string DescriNom = rows["DescriNom"].ToString();
                                string ActDescri = rows["ActDescri"].ToString();

                                //Variables crud Detalle 2
                                string CodPeHom = rows["CodPeHom"].ToString();
                                string NomPeHom = rows["NomPeHom"].ToString();
                                string CodPeKac = rows["CodPeKac"].ToString();
                                string NomPeHom1 = rows["NomPeHom1"].ToString();
                                string ActNomPeHom1 = rows["ActNomPeHom1"].ToString();

                                //Variables crud Detalle 3
                                string CodTip = rows["CodTip"].ToString();
                                string NomTip = rows["NomTip"].ToString();
                                string ActNomTip = rows["ActNomTip"].ToString();

                                //Variables crud Detalle 4
                                string CodConceKac = rows["CodConceKac"].ToString();
                                string CodConceHom = rows["CodConceHom"].ToString();
                                string NomConceHom = rows["NomConceHom"].ToString();
                                string Porce = rows["Porce"].ToString();
                                string CodConceKac1 = rows["CodConceKac1"].ToString();
                                string CodConceHom1 = rows["CodConceHom1"].ToString();
                                string NomConceHom1 = rows["NomConceHom1"].ToString();
                                string Porce1 = rows["Porce1"].ToString();
                                string CodConceKac2 = rows["CodConceKac2"].ToString();
                                string CodConceHom2 = rows["CodConceHom2"].ToString();
                                string NomConceHom2 = rows["NomConceHom2"].ToString();
                                string Porce2 = rows["Porce2"].ToString();
                                string CodConceKac3 = rows["CodConceKac3"].ToString();
                                string CodConceHom3 = rows["CodConceHom3"].ToString();
                                string NomConceHom3 = rows["NomConceHom3"].ToString();
                                string CodConceKac4 = rows["CodConceKac4"].ToString();
                                string CodConceHom4 = rows["CodConceHom4"].ToString();
                                string NomConceHom4 = rows["NomConceHom4"].ToString();
                                string Porce4 = rows["Porce4"].ToString();
                                string CodConceKac5 = rows["CodConceKac5"].ToString();
                                string CodConceHom5 = rows["CodConceHom5"].ToString();
                                string NomConceHom5 = rows["NomConceHom5"].ToString();
                                string Porce5 = rows["Porce5"].ToString();
                                string ActNomConceHom5 = rows["ActNomConceHom5"].ToString();

                                //Variables crud Detalle 5
                                string CodHom = rows["CodHom"].ToString();

                                //Variables crud Detalle 6
                                string CodHom1 = rows["CodHom1"].ToString();

                                //Variables crud Detalle 7
                                string CodInca = rows["CodInca"].ToString();
                                string NomInca = rows["NomInca"].ToString();
                                string ActNomInca = rows["ActNomInca"].ToString();

                                //Variables crud Detalle 8
                                string CodExclu = rows["CodExclu"].ToString();
                                string ActCodExclu = rows["ActCodExclu"].ToString();

                                String campo = "2";

                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { CodOpe, NomOpe, NIT, ActNombre };

                                //LISTA CRUD DETALLE
                                List<string> crudDet = new List<string>() { CodNom, DescriNom, ActDescri };

                                //LISTA CRUD DETALLE 2
                                List<string> crudDet2 = new List<string>() { CodPeHom, NomPeHom, CodPeKac, NomPeHom1, ActNomPeHom1 };

                                //LISTA CRUD DETALLE 3
                                List<string> crudDet3 = new List<string>() { CodTip, NomTip, ActNomTip };

                                //LISTA CRUD DETALLE 4
                                List<string> crudDet4 = new List<string>() { CodConceKac, CodConceHom, NomConceHom, Porce,
                                                                             CodConceKac1, CodConceHom1, NomConceHom1, Porce1,
                                                                             CodConceKac2, CodConceHom2, NomConceHom2, Porce2,
                                                                             CodConceKac3, CodConceHom3, NomConceHom3,
                                                                             CodConceKac4, CodConceHom4, NomConceHom4, Porce4,
                                                                             CodConceKac5, CodConceHom5, NomConceHom5, Porce5, ActNomConceHom5};

                                //LISTA CRUD DETALLE 5
                                List<string> crudDet5 = new List<string>() { CodHom };

                                //LISTA CRUD DETALLE 6
                                List<string> crudDet6 = new List<string>() { CodHom1 };

                                //LISTA CRUD DETALLE 7
                                List<string> crudDet7 = new List<string>() { CodInca, NomInca, ActNomInca };

                                //LISTA CRUD DETALLE 8
                                List<string> crudDet8 = new List<string>() { CodExclu, ActCodExclu };

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

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, Campo, c, ErrorValidacion, "No se agregó correctamente", 0, "1");
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDet, user, motor, CampoDet, Codigo, Codigo, CampoDet, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //MODIFICAMOS
                                        //CAMBIAR EMPRESA
                                        CrudKnmnoeco.Click(desktopSession, 6);

                                        //HACEMOS QBE
                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_1.lapiz(desktopSession);

                                        //Thread.Sleep(2000);

                                        CrudKnmnoeco.Click(desktopSession, 8);
                                        newFunctions_4.ScreenshotNuevo("Cambiar a ventana crud excluir contratistas", true, file);
                                        Thread.Sleep(1500);

                                        //ELIMINAMOS REGISTRO SUBTIPO DE COTIZANTE
                                        Dictionary<string, Point> botonesNavegador9 = new Dictionary<string, Point>();
                                        botonesNavegador9 = PruebaCRUD.findname(desktopSession, 30, 1);
                                        var ElementList9 = PruebaCRUD.NavClass(desktopSession);

                                        ////LUPA
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(1500);

                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList9[1], botonesNavegador9, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 7 (excluir contratistas)", true, file);
                                        Thread.Sleep(2000);
                                        ///////////////////////////////////////////////////////////////////////////////////////

                                        CrudKnmnoeco.Click(desktopSession, 7);
                                        newFunctions_4.ScreenshotNuevo("Cambiar a ventana crud tipos de incapacidad", true, file);
                                        Thread.Sleep(1500);

                                        //ELIMINAMOS REGISTRO SUBTIPO DE COTIZANTE
                                        Dictionary<string, Point> botonesNavegador8 = new Dictionary<string, Point>();
                                        botonesNavegador8 = PruebaCRUD.findname(desktopSession, 30, 1);
                                        var ElementList8 = PruebaCRUD.NavClass(desktopSession);

                                        ////LUPA
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(1500);

                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList8[1], botonesNavegador8, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 7 (tipos de incapacidad)", true, file);
                                        Thread.Sleep(2000);
                                        ///////////////////////////////////////////////////////////////////////////////////////

                                        CrudKnmnoeco.Click(desktopSession, 4);
                                        newFunctions_4.ScreenshotNuevo("Cambiar a ventana crud Subtipo de Cotizante", true, file);
                                        Thread.Sleep(1500);

                                        //ELIMINAMOS REGISTRO SUBTIPO DE COTIZANTE
                                        Dictionary<string, Point> botonesNavegador7 = new Dictionary<string, Point>();
                                        botonesNavegador7 = PruebaCRUD.findname(desktopSession, 30, 1);
                                        var ElementList7 = PruebaCRUD.NavClass(desktopSession);

                                        ////LUPA
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(1500);

                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList7[1], botonesNavegador7, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 6 (Subtipo cotizante)", true, file);
                                        Thread.Sleep(2000);
                                        ///////////////////////////////////////////////////////////////////////////////////////
                                        CrudKnmnoeco.Click(desktopSession, 3);
                                        newFunctions_4.ScreenshotNuevo("Cambiar a ventana crud Tipo de Cotizante", true, file);
                                        Thread.Sleep(1500);

                                        //ELIMINAMOS REGISTRO TIPO DE COTIZANTE
                                        Dictionary<string, Point> botonesNavegador6 = new Dictionary<string, Point>();
                                        botonesNavegador6 = PruebaCRUD.findname(desktopSession, 30, 1);
                                        var ElementList6 = PruebaCRUD.NavClass(desktopSession);

                                        ////LUPA
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(1500);

                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList6[1], botonesNavegador6, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 5 (Tipo Cotizante)", true, file);
                                        Thread.Sleep(2000);
                                        ///////////////////////////////////////////////////////////////////////////////////////
                                        CrudKnmnoeco.Click(desktopSession, 2);
                                        newFunctions_4.ScreenshotNuevo("Cambiar a ventana crud Horas Extras", true, file);
                                        Thread.Sleep(1500);

                                        //ELIMINAMOS REGISTRO Horas Extra
                                        Dictionary<string, Point> botonesNavegador5 = new Dictionary<string, Point>();
                                        botonesNavegador5 = PruebaCRUD.findname(desktopSession, 30, 1);
                                        var ElementList5 = PruebaCRUD.NavClass(desktopSession);

                                        ////LUPA
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(1500);

                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList5[1], botonesNavegador5, file);
                                        Thread.Sleep(2000);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList5[1], botonesNavegador5, file);
                                        Thread.Sleep(2000);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList5[1], botonesNavegador5, file);
                                        Thread.Sleep(2000);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList5[1], botonesNavegador5, file);
                                        Thread.Sleep(2000);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList5[1], botonesNavegador5, file);
                                        Thread.Sleep(2000);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList5[1], botonesNavegador5, file);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 4 (Horas Extra)", true, file);
                                        Thread.Sleep(2000);
                                        ///////////////////////////////////////////////////////////////////////////////////////
                                        CrudKnmnoeco.Click(desktopSession, 1);
                                        newFunctions_4.ScreenshotNuevo("Cambiar a ventana crud Tipos de Contrato", true, file);
                                        Thread.Sleep(1500);

                                        //ELIMINAMOS REGISTRO Tipos de Contrato
                                        Dictionary<string, Point> botonesNavegador4 = new Dictionary<string, Point>();
                                        botonesNavegador4 = PruebaCRUD.findname(desktopSession, 30, 1);
                                        var ElementList4 = PruebaCRUD.NavClass(desktopSession);

                                        ////LUPA
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(1500);

                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList4[1], botonesNavegador4, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 3 (Tipos de Contrato)", true, file);
                                        Thread.Sleep(2000);
                                        ///////////////////////////////////////////////////////////////////////////////////////
                                        CrudKnmnoeco.Click(desktopSession, 0);
                                        newFunctions_4.ScreenshotNuevo("Cambiar a ventana crud Periodos nomina", true, file);
                                        Thread.Sleep(1500);

                                        //ELIMINAMOS REGISTRO Periodos de Nomina
                                        Dictionary<string, Point> botonesNavegador3 = new Dictionary<string, Point>();
                                        botonesNavegador3 = PruebaCRUD.findname(desktopSession, 30, 1);
                                        var ElementList3 = PruebaCRUD.NavClass(desktopSession);

                                        ////LUPA
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(1500);

                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList3[1], botonesNavegador3, file);
                                        Thread.Sleep(2000);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList3[1], botonesNavegador3, file);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 2 (Periodos de Nomina)", true, file);
                                        Thread.Sleep(2000);
                                        ///////////////////////////////////////////////////////////////////////////////////////
                                        CrudKnmnoeco.Click(desktopSession, 5);
                                        newFunctions_4.ScreenshotNuevo("Cambiar a ventana crud Homologacion medios de pago", true, file);
                                        Thread.Sleep(1500);

                                        //ELIMINAMOS REGISTRO Homologacion medios de pago
                                        Dictionary<string, Point> botonesNavegador2 = new Dictionary<string, Point>();
                                        botonesNavegador2 = PruebaCRUD.findname(desktopSession, 30, 1);
                                        var ElementList2 = PruebaCRUD.NavClass(desktopSession);

                                        ////LUPA
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(1500);

                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList2[1], botonesNavegador2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle (Homologacion medios de pago)", true, file);
                                        Thread.Sleep(2000);
                                        ///////////////////////////////////////////////////////////////////////////////

                                        //ELIMINAMOS REGISTRO PRINCIPAL
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        ////LUPA
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(1500);

                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados crud principal", true, file);
                                        Thread.Sleep(2000);


                                        //AGREGAR REGISTRO Principal
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKnmnoeco.AgregarCrud(desktopSession, 0, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Cargo, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        Thread.Sleep(3000);
                                                                               
                                        //// //Debugger.Launch();
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, ActCargo, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //CrudKnmnoeco.Click(desktopSession, 0);
                                        //newFunctions_4.ScreenshotNuevo("Cambiar a ventana crud detalle perfiles", true, file);
                                        //Thread.Sleep(1500);

                                        ////AGREGAR REGISTRO DETALLE 1
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Insertar"].X, botonesNavegador2["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        //CrudKnmcaban.Click(desktopSession);
                                        //Thread.Sleep(1500);
                                        CrudKnmnoeco.AgregarCrud1(desktopSession, 0, crudDet);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1500);

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, Codigo, CampoDetalle1, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud1(desktopSession, 1, crudDet);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Cancelar"].X, botonesNavegador2["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, Cantidad, EditCampoDetalle1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud1(desktopSession, 1, crudDet);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        Thread.Sleep(3000);
                                        //validacion error al editar
                                        //bool valEditOra1 = PruebaCRUD.ValEditOra(desktopSession);
                                        //if (valEditOra1 == true)
                                        //{
                                        //    ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
                                        //    //LUPA
                                        //    errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //    newFunctions_3.LupaAud(desktopSession, "1", file);
                                        //    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //    //ACEPTAR EDICION
                                        //    desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        //    CrudKnmnoeco.AgregarCrud1(desktopSession, 1, crudDet);
                                        //    newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        //    desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    Thread.Sleep(3000);
                                        //}
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        CrudKnmnoeco.Click(desktopSession, 0);
                                        newFunctions_4.ScreenshotNuevo("Cambiar a ventana crud Periodos de Nomina", true, file);
                                        Thread.Sleep(1500);

                                        ////AGREGAR REGISTRO DETALLE 2
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Insertar"].X, botonesNavegador3["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        //CrudKnmcaban.Click(desktopSession);
                                        //Thread.Sleep(1500);
                                        CrudKnmnoeco.AgregarCrud2(desktopSession, 0, crudDet2);

                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Aplicar"].X, botonesNavegador3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);

                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Insertar"].X, botonesNavegador3["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1500);
                                        CrudKnmnoeco.AgregarCrud2(desktopSession, 2, crudDet2);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);
                                        Thread.Sleep(1500);
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Aplicar"].X, botonesNavegador3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, Codigo, CampoDetalle1, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Editar"].X, botonesNavegador3["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud2(desktopSession, 1, crudDet2);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Cancelar"].X, botonesNavegador3["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, Cantidad, EditCampoDetalle1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Editar"].X, botonesNavegador3["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud2(desktopSession, 1, crudDet2);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Aplicar"].X, botonesNavegador3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        Thread.Sleep(3000);
                                        //validacion error al editar
                                        bool valEditOra2 = PruebaCRUD.ValEditOra(desktopSession);
                                        //if (valEditOra2 == true)
                                        //{
                                        //    ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
                                        //    //LUPA
                                        //    errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //    newFunctions_3.LupaAud(desktopSession, "1", file);
                                        //    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //    //ACEPTAR EDICION
                                        //    desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Editar"].X, botonesNavegador3["Editar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        //    CrudKnmnoeco.AgregarCrud2(desktopSession, 1, crudDet2);
                                        //    newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        //    desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Aplicar"].X, botonesNavegador3["Aplicar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    Thread.Sleep(3000);
                                        //}
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        //Debugger.Launch();
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle1, motor, user, CampoDetalle1, Codigo, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, ActCantidad, EditCampoDetalle1, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        CrudKnmnoeco.Click(desktopSession, 1);
                                        newFunctions_4.ScreenshotNuevo("Cambiar a ventana crud Tipos de contrato", true, file);
                                        Thread.Sleep(1500);

                                        ////AGREGAR REGISTRO DETALLE 3
                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Insertar"].X, botonesNavegador4["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        //CrudKnmcaban.Click(desktopSession);
                                        //Thread.Sleep(1500);
                                        CrudKnmnoeco.AgregarCrud3(desktopSession, 0, crudDet3);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Aplicar"].X, botonesNavegador4["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, Codigo, CampoDetalle1, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Editar"].X, botonesNavegador4["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud3(desktopSession, 1, crudDet3);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Cancelar"].X, botonesNavegador4["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, Cantidad, EditCampoDetalle1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Editar"].X, botonesNavegador4["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud3(desktopSession, 1, crudDet3);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Aplicar"].X, botonesNavegador4["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        Thread.Sleep(3000);
                                        //validacion error al editar
                                        //bool valEditOra3 = PruebaCRUD.ValEditOra(desktopSession);
                                        //if (valEditOra3 == true)
                                        //{
                                        //    ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
                                        //    //LUPA
                                        //    errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //    newFunctions_3.LupaAud(desktopSession, "1", file);
                                        //    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //    //ACEPTAR EDICION
                                        //    desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Editar"].X, botonesNavegador4["Editar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        //    CrudKnmnoeco.AgregarCrud3(desktopSession, 1, crudDet3);
                                        //    newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        //    desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegador4["Aplicar"].X, botonesNavegador4["Aplicar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    Thread.Sleep(3000);
                                        //}
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        //Debugger.Launch();
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle1, motor, user, CampoDetalle1, Codigo, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, ActCantidad, EditCampoDetalle1, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        CrudKnmnoeco.Click(desktopSession, 2);
                                        newFunctions_4.ScreenshotNuevo("Cambiar a ventana crud Experiencia", true, file);
                                        Thread.Sleep(1500);

                                        ////AGREGAR REGISTRO DETALLE 4
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Insertar"].X, botonesNavegador5["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        //CrudKnmcaban.Click(desktopSession);
                                        //Thread.Sleep(1500);
                                        CrudKnmnoeco.AgregarCrud4(desktopSession, 0, crudDet4);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Aplicar"].X, botonesNavegador5["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        Thread.Sleep(1500);

                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Insertar"].X, botonesNavegador5["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1500);
                                        //CrudKnmcaban.Click(desktopSession);
                                        //Thread.Sleep(1500);
                                        CrudKnmnoeco.AgregarCrud4(desktopSession, 2, crudDet4);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);
                                        Thread.Sleep(1500);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Aplicar"].X, botonesNavegador5["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        Thread.Sleep(1500);

                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Insertar"].X, botonesNavegador5["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1500);
                                        //CrudKnmcaban.Click(desktopSession);
                                        //Thread.Sleep(1500);
                                        CrudKnmnoeco.AgregarCrud4(desktopSession, 3, crudDet4);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);
                                        Thread.Sleep(1500);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Aplicar"].X, botonesNavegador5["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        Thread.Sleep(1500);

                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Insertar"].X, botonesNavegador5["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1500);
                                        //CrudKnmcaban.Click(desktopSession);
                                        //Thread.Sleep(1500);
                                        CrudKnmnoeco.AgregarCrud4(desktopSession, 4, crudDet4);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);
                                        Thread.Sleep(1500);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Aplicar"].X, botonesNavegador5["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        Thread.Sleep(1500);

                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Insertar"].X, botonesNavegador5["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1500);
                                        //CrudKnmcaban.Click(desktopSession);
                                        //Thread.Sleep(1500);
                                        CrudKnmnoeco.AgregarCrud4(desktopSession, 5, crudDet4);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);
                                        Thread.Sleep(1500);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Aplicar"].X, botonesNavegador5["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        Thread.Sleep(1500);

                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Insertar"].X, botonesNavegador5["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1500);
                                        //CrudKnmcaban.Click(desktopSession);
                                        //Thread.Sleep(1500);
                                        CrudKnmnoeco.AgregarCrud4(desktopSession, 6, crudDet4);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);
                                        Thread.Sleep(1500);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Aplicar"].X, botonesNavegador5["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos insertados", true, file);

                                        Thread.Sleep(1500);
                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, Codigo, CampoDetalle1, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Editar"].X, botonesNavegador5["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud4(desktopSession, 1, crudDet4);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Cancelar"].X, botonesNavegador5["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, Cantidad, EditCampoDetalle1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Editar"].X, botonesNavegador5["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud4(desktopSession, 1, crudDet4);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Aplicar"].X, botonesNavegador5["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        Thread.Sleep(3000);
                                        //validacion error al editar
                                        //bool valEditOra4 = PruebaCRUD.ValEditOra(desktopSession);
                                        //if (valEditOra4 == true)
                                        //{
                                        //    ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
                                        //    //LUPA
                                        //    errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //    newFunctions_3.LupaAud(desktopSession, "1", file);
                                        //    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //    //ACEPTAR EDICION
                                        //    desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Editar"].X, botonesNavegador5["Editar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        //    CrudKnmnoeco.AgregarCrud4(desktopSession, 1, crudDet4);
                                        //    newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        //    desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador5["Aplicar"].X, botonesNavegador5["Aplicar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    Thread.Sleep(3000);
                                        //}
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        //Debugger.Launch();
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle1, motor, user, CampoDetalle1, Codigo, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, ActCantidad, EditCampoDetalle1, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        CrudKnmnoeco.Click(desktopSession, 3);
                                        newFunctions_4.ScreenshotNuevo("Cambiar a ventana crud Tipo de Cotizante", true, file);
                                        Thread.Sleep(1500);

                                        ////AGREGAR REGISTRO DETALLE 5
                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador6["Insertar"].X, botonesNavegador6["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        //CrudKnmcaban.Click(desktopSession);
                                        //Thread.Sleep(1500);
                                        CrudKnmnoeco.AgregarCrud5(desktopSession, 0, crudDet5);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador6["Aplicar"].X, botonesNavegador6["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, Codigo, CampoDetalle1, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador6["Editar"].X, botonesNavegador6["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud5(desktopSession, 1, crudDet5);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador6["Cancelar"].X, botonesNavegador6["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, Cantidad, EditCampoDetalle1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador6["Editar"].X, botonesNavegador6["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud5(desktopSession, 1, crudDet5);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador6["Aplicar"].X, botonesNavegador6["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        Thread.Sleep(3000);
                                        //validacion error al editar
                                        //bool valEditOra5 = PruebaCRUD.ValEditOra(desktopSession);
                                        //if (valEditOra5 == true)
                                        //{
                                        //    ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
                                        //    //LUPA
                                        //    errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //    newFunctions_3.LupaAud(desktopSession, "1", file);
                                        //    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //    //ACEPTAR EDICION
                                        //    desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador6["Editar"].X, botonesNavegador6["Editar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        //    CrudKnmnoeco.AgregarCrud5(desktopSession, 1, crudDet5);
                                        //    newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        //    desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador6["Aplicar"].X, botonesNavegador6["Aplicar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    Thread.Sleep(3000);
                                        //}
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        //Debugger.Launch();
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle1, motor, user, CampoDetalle1, Codigo, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, ActCantidad, EditCampoDetalle1, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        CrudKnmnoeco.Click(desktopSession, 4);
                                        newFunctions_4.ScreenshotNuevo("Cambiar a ventana crud Subtipo de Cotizante", true, file);
                                        Thread.Sleep(1500);

                                        ////AGREGAR REGISTRO DETALLE 6                                      
                                        
                                        desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegador7["Insertar"].X, botonesNavegador7["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        //CrudKnmcaban.Click(desktopSession);
                                        //Thread.Sleep(1500);
                                        CrudKnmnoeco.AgregarCrud6(desktopSession, 0, crudDet6);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegador7["Aplicar"].X, botonesNavegador7["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, Codigo, CampoDetalle1, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegador7["Editar"].X, botonesNavegador7["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud6(desktopSession, 1, crudDet6);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegador7["Cancelar"].X, botonesNavegador7["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, Cantidad, EditCampoDetalle1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegador7["Editar"].X, botonesNavegador7["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud6(desktopSession, 1, crudDet6);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegador7["Aplicar"].X, botonesNavegador7["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        Thread.Sleep(3000);
                                        //validacion error al editar
                                        //bool valEditOra6 = PruebaCRUD.ValEditOra(desktopSession);
                                        //if (valEditOra6 == true)
                                        //{
                                        //    ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
                                        //    //LUPA
                                        //    errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //    newFunctions_3.LupaAud(desktopSession, "1", file);
                                        //    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //    //ACEPTAR EDICION
                                        //    desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegador7["Editar"].X, botonesNavegador7["Editar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        //    CrudKnmnoeco.AgregarCrud6(desktopSession, 1, crudDet6);
                                        //    newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        //    desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegador7["Aplicar"].X, botonesNavegador7["Aplicar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    Thread.Sleep(3000);
                                        //}
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        //Debugger.Launch();
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle1, motor, user, CampoDetalle1, Codigo, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, ActCantidad, EditCampoDetalle1, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        CrudKnmnoeco.Click(desktopSession, 7);
                                        newFunctions_4.ScreenshotNuevo("Cambiar a ventana crud tipos de incapacidad", true, file);
                                        Thread.Sleep(1500);

                                        ////AGREGAR REGISTRO DETALLE 7                                      

                                        desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador8["Insertar"].X, botonesNavegador8["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        //CrudKnmcaban.Click(desktopSession);
                                        //Thread.Sleep(1500);
                                        CrudKnmnoeco.AgregarCrud7(desktopSession, 0, crudDet7);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador8["Aplicar"].X, botonesNavegador8["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, Codigo, CampoDetalle1, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador8["Editar"].X, botonesNavegador8["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud7(desktopSession, 1, crudDet7);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador8["Cancelar"].X, botonesNavegador8["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, Cantidad, EditCampoDetalle1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador8["Editar"].X, botonesNavegador8["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud7(desktopSession, 1, crudDet7);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador8["Aplicar"].X, botonesNavegador8["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        Thread.Sleep(3000);
                                        //validacion error al editar
                                        //bool valEditOra7 = PruebaCRUD.ValEditOra(desktopSession);
                                        //if (valEditOra7 == true)
                                        //{
                                        //    ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
                                        //    //LUPA
                                        //    errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //    newFunctions_3.LupaAud(desktopSession, "1", file);
                                        //    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //    //ACEPTAR EDICION
                                        //    desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador8["Editar"].X, botonesNavegador8["Editar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        //    CrudKnmnoeco.AgregarCrud7(desktopSession, 1, crudDet7);
                                        //    newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        //    desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador8["Aplicar"].X, botonesNavegador8["Aplicar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    Thread.Sleep(3000);
                                        //}
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);

                                        CrudKnmnoeco.Click(desktopSession, 8);
                                        newFunctions_4.ScreenshotNuevo("Cambiar a ventana crud excluir contratistas", true, file);
                                        Thread.Sleep(1500);

                                        ////AGREGAR REGISTRO DETALLE 8                                     

                                        desktopSession.Mouse.MouseMove(ElementList9[1].Coordinates, botonesNavegador9["Insertar"].X, botonesNavegador9["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        //CrudKnmcaban.Click(desktopSession);
                                        //Thread.Sleep(1500);
                                        CrudKnmnoeco.AgregarCrud8(desktopSession, 0, crudDet8);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList9[1].Coordinates, botonesNavegador9["Aplicar"].X, botonesNavegador9["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, Codigo, CampoDetalle1, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList9[1].Coordinates, botonesNavegador9["Editar"].X, botonesNavegador9["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud8(desktopSession, 1, crudDet8);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList9[1].Coordinates, botonesNavegador9["Cancelar"].X, botonesNavegador9["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, Cantidad, EditCampoDetalle1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList9[1].Coordinates, botonesNavegador9["Editar"].X, botonesNavegador9["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKnmnoeco.AgregarCrud8(desktopSession, 1, crudDet8);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList9[1].Coordinates, botonesNavegador9["Aplicar"].X, botonesNavegador9["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        Thread.Sleep(3000);
                                        //validacion error al editar
                                        //bool valEditOra8 = PruebaCRUD.ValEditOra(desktopSession);
                                        //if (valEditOra8 == true)
                                        //{
                                        //    ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
                                        //    //LUPA
                                        //    errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //    newFunctions_3.LupaAud(desktopSession, "1", file);
                                        //    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //    //ACEPTAR EDICION
                                        //    desktopSession.Mouse.MouseMove(ElementList9[1].Coordinates, botonesNavegador9["Editar"].X, botonesNavegador9["Editar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        //    CrudKnmnoeco.AgregarCrud8(desktopSession, 1, crudDet8);
                                        //    newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        //    desktopSession.Mouse.MouseMove(ElementList9[1].Coordinates, botonesNavegador9["Aplicar"].X, botonesNavegador9["Aplicar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    Thread.Sleep(3000);
                                        //}
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);


                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
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

                                        Thread.Sleep(2000);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                                                                
                                        Thread.Sleep(2000);                                        



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
        public void KNmLinoeChecKlist()
        {
            //Debugger.Launch();
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_97.KNmLinoeChecKlist")
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
                                rows["Prefi"].ToString().Length != 0 && rows["Prefi"].ToString() != null &&
                                rows["TMR"].ToString().Length != 0 && rows["TMR"].ToString() != null &&
                                rows["FechIni"].ToString().Length != 0 && rows["FechIni"].ToString() != null &&
                                rows["FechFin"].ToString().Length != 0 && rows["FechFin"].ToString() != null &&
                                rows["Obser"].ToString().Length != 0 && rows["Obser"].ToString() != null &&
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null)

                            {
                                //Variables obligatorias
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();

                                //Variables generales
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string Prefi = rows["Prefi"].ToString();
                                string TMR = rows["TMR"].ToString();
                                string FechIni = rows["FechIni"].ToString();
                                string FechFin = rows["FechFin"].ToString();
                                string Obser = rows["Obser"].ToString();


                                //LISTA datos
                                List<string> crudPrincipal = new List<string>() { Prefi, TMR, FechIni, FechFin, Obser };


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

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        CrudKnmlinoe.Click(desktopSession, 1, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Pasamos a la empresa DIGITAL WARE", true, file);
                                        Thread.Sleep(2000);

                                        CrudKnmlinoe.Click(desktopSession, 0, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Se llenan datos necesarios", true, file);
                                        Thread.Sleep(2000);

                                        //CrudKnmcomps.ClickButtonAceptarEnter(desktopSession, 0);
                                        //Thread.Sleep(2000);

                                        //CrudKnmcomps.ClickButtonAceptarEnter(desktopSession, 1);
                                        //Thread.Sleep(2000);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(2000);

                                        CrudKnmlinoe.ClickButtonAceptarEnter(desktopSession, 0);
                                        Thread.Sleep(5000);

                                        newFunctions_4.ScreenshotNuevo("Mensaje de aceptar", true, file);
                                        Thread.Sleep(4000);
                                        CrudKnmlinoe.ClickButtonAceptarEnter(desktopSession, 1);

                                        newFunctions_4.ScreenshotNuevo("Mensaje de proceso satisfactorio", true, file);
                                        Thread.Sleep(2000);
                                        CrudKnmlinoe.ClickButtonAceptarEnter(desktopSession, 1);


                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Maestro Finalizado", true, file);

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

using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using Ophelia_Test_V21.AFuncionesGlobales;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBp;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosEd;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosFd;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGd;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGe;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosIG;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.ValidacionQueries;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Threading;

//AUTOR JUANGO
namespace Ophelia_Test_V21.TestMethods.ModulosBI
{
    /// <summary>
    /// Descripción resumida de ModulosBi_10
    /// </summary>
    [TestClass]
    public class ModulosBi_10 : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        protected static WindowsDriver<WindowsElement> desktopSession;
        protected static WindowsDriver<WindowsElement> rootSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();

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

        [TestMethod]
        public void AbrirPrograma()
        {
            string user = "UCalida1";
            string motor = "SQL";
            string codProgram = "KIgLindi";

            string id = "1213141516";
            string codSolicitud = "4321";
            string codEsta = "1234";
            string fecha = "04/01/2021";
            string tipoApli = "V";
            string observa = "NINGUNA";
            List<string> data = new List<string> { id, codSolicitud, codEsta, fecha, tipoApli, observa };
            //List<string> data1 = new List<string> { idDetalle, accion, servicio, metodo, editDetalle };
            //List<string> data2 = new List<string> { arbol, id, editDetalle2 };

            AbrirPrograma a = new AbrirPrograma(codProgram, user);
            desktopSession = a.Start2(motor, "");
            Thread.Sleep(2000);

            ////error clic derecho al iniciar programa
            //WindowsDriver<WindowsElement> rootSession = null;
            //rootSession = newFunctions_4.RootSessionNew();
            //WebDriverWait close = new WebDriverWait(rootSession, TimeSpan.FromSeconds(5));
            //Thread.Sleep(2000);
            //var parentElement = rootSession.FindElement(By.Name(rootSession.Title));
            //Thread.Sleep(2000);
            //new Actions(rootSession).MoveToElement(parentElement, -100, 30).MoveByOffset(parentElement.Size.Width / 2, parentElement.Size.Height / 2).Click().Perform();
            //desktopSession = newFunctions_4.RootSessionNew();
            //desktopSession = newFunctions_4.ReloadXSession(desktopSession, "IEFrame");
            //close = new WebDriverWait(desktopSession, TimeSpan.FromSeconds(5));
            //Thread.Sleep(2000);

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
                //string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);

                ////VALIDACION CODIGO DEL PROGRAMA
                //FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);

                ////VERSION
                //FuncionesGlobales.GetVersion(desktopSession, file);
                //Thread.Sleep(1000);
                ////NOTAS
                //newFunctions_4.openInnerNote(desktopSession, file);
                //Thread.Sleep(1000);

                //// Insertar
                //Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                //botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                //var ElementList = PruebaCRUD.NavClass(desktopSession);

                //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);

                CrudKedeseva.CrudKEdeseva(desktopSession, 1, data);
                Thread.Sleep(1000);

                //// Aplicar
                //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);

                //// Editar
                //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);

                //CrudKnmrcnar.CrudKNmrcnar(desktopSession, 2, data);

                //// Descartar Edicion
                //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);

                //// Editar
                //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);

                //CrudKnmrcnar.CrudKNmrcnar(desktopSession, 2, data);

                //// Aplicar Edicion
                //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);


                //Detalle 1
                ////desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Áreas de Experiencia").Coordinates);
                ////desktopSession.Mouse.Click(null);

                //Dictionary<string, Point> botonesNavegadorDetalle = new Dictionary<string, Point>();
                //botonesNavegadorDetalle = PruebaCRUD.findname(desktopSession, 30, 1);
                //ElementList = PruebaCRUD.NavClass(desktopSession);

                ////Insertar Detalle 1
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Insertar"].X, botonesNavegadorDetalle["Insertar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                //CrudKfdlogis.CrudKFdlogisDetalle1(desktopSession, 1, data1);

                //// Aplicar detalle 1
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);

                //// Editar detalle 1
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);

                //CrudKfdlogis.CrudKFdlogisDetalle1(desktopSession, 2, data1);

                //// Descartar Edicion detalle 1
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Cancelar"].X, botonesNavegadorDetalle["Cancelar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);

                //// Editar detalle 1
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);

                //CrudKfdlogis.CrudKFdlogisDetalle1(desktopSession, 2, data1);

                //// Aplicar Edicion detalle 1
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);


                ////Detalle 2
                //desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Clasificacion").Coordinates);
                //desktopSession.Mouse.Click(null);

                //Dictionary<string, Point> botonesNavegadorDetalle2 = new Dictionary<string, Point>();
                //botonesNavegadorDetalle2 = PruebaCRUD.findname(desktopSession, 28, 1);
                //ElementList = PruebaCRUD.NavClass(desktopSession);

                ////Insertar Detalle 2
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Insertar"].X, botonesNavegadorDetalle2["Insertar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                //CrudKfdlogis.CrudKFdlogisDetalle2(desktopSession, 1, data2);

                //// Aplicar detalle 2
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);

                //// Editar detalle 2
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);

                //CrudKfdlogis.CrudKFdlogisDetalle2(desktopSession, 2, data2);

                //// Descartar Edicion detalle 2
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Cancelar"].X, botonesNavegadorDetalle2["Cancelar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);

                //// Editar detalle 2
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);

                //CrudKfdlogis.CrudKFdlogisDetalle2(desktopSession, 2, data2);

                //// Aplicar Edicion detalle 2
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);

                ////Detalle 3
                //Dictionary<string, Point> botonesNavegadorDetalle3 = new Dictionary<string, Point>();
                //desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Perfiles").Coordinates);
                //desktopSession.Mouse.Click(null);

                ////Dictionary<string, Point> botonesNavegadorDetalle = new Dictionary<string, Point>();
                //botonesNavegadorDetalle = PruebaCRUD.findname(desktopSession, 25, 1);
                //ElementList = PruebaCRUD.NavClass(desktopSession);

                ////Insertar Detalle 3
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle3["Insertar"].X, botonesNavegadorDetalle3["Insertar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                //CrudKbihvext.CrudKBihvextDetalle(desktopSession, 1, tipo, req, espec, editDetalle);

                //// Aplicar detalle 3
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle3["Aplicar"].X, botonesNavegadorDetalle3["Aplicar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);

                //// Editar detalle 3
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle3["Editar"].X, botonesNavegadorDetalle3["Editar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);

                //CrudKbihvext.CrudKBihvextDetalle(desktopSession, 2, tipo, req, espec, editDetalle);

                //// Descartar Edicion detalle 3
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle3["Cancelar"].X, botonesNavegadorDetalle3["Cancelar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);

                //// Editar detalle 3
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle3["Editar"].X, botonesNavegadorDetalle3["Editar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);

                //CrudKbihvext.CrudKBihvextDetalle(desktopSession, 2, tipo, req, espec, editDetalle);

                //// Aplicar Edicion detalle 3
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle3["Aplicar"].X, botonesNavegadorDetalle3["Aplicar"].Y);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);
            }
        }

        [TestMethod]
        public void KBiEdnfoChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosBI.ModulosBi_10.KBiEdnfoChecklist")
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
                                //Datos Login
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["id"].ToString().Length != 0 && rows["id"].ToString() != null &&
                                rows["modalidad"].ToString().Length != 0 && rows["modalidad"].ToString() != null &&
                                rows["curso"].ToString().Length != 0 && rows["curso"].ToString() != null &&
                                rows["nomInsti"].ToString().Length != 0 && rows["nomInsti"].ToString() != null &&
                                rows["nomEspe"].ToString().Length != 0 && rows["nomEspe"].ToString() != null &&
                                rows["fechaIni"].ToString().Length != 0 && rows["fechaIni"].ToString() != null &&
                                rows["editCabeza"].ToString().Length != 0 && rows["editCabeza"].ToString() != null &&
                                rows["tipo"].ToString().Length != 0 && rows["tipo"].ToString() != null &&
                                rows["req"].ToString().Length != 0 && rows["req"].ToString() != null &&
                                rows["espec"].ToString().Length != 0 && rows["espec"].ToString() != null &&
                                rows["editDetalle"].ToString().Length != 0 && rows["editDetalle"].ToString() != null &&
                                rows["campoEdit"].ToString().Length != 0 && rows["campoEdit"].ToString() != null)
                            {
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
                                string BanderaPreli = rows["BanderaPreli"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();

                                string id = rows["id"].ToString();
                                string modalidad = rows["modalidad"].ToString();
                                string curso = rows["curso"].ToString();
                                string nomInsti = rows["nomInsti"].ToString();
                                string nomEspe = rows["nomEspe"].ToString();
                                string fechaIni = rows["fechaIni"].ToString();
                                string editCabeza = rows["editCabeza"].ToString();
                                string tipo = rows["tipo"].ToString();
                                string req = rows["req"].ToString();
                                string espec = rows["espec"].ToString();
                                string editDetalle = rows["editDetalle"].ToString();
                                string campoEdit = rows["campoEdit"].ToString();
                                List<string> Lupas = new List<string> { id, modalidad, curso };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "BI", motor);

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

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera("BI_DTPNF", user, motor, "COD_ESPE", "999", "999", "COD_ESPE", c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, id, id, Campo, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        // Insertar
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Insertar Registro", true, file);

                                        CrudKbiednfo.CrudKBiednfo(desktopSession, 1, Lupas, nomInsti, nomEspe, fechaIni, editCabeza);

                                        // Aplicar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Registro Insertado", true, file);

                                        //Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, id, id, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        // Editar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);

                                        CrudKbiednfo.CrudKBiednfo(desktopSession, 2, Lupas, nomInsti, nomEspe, fechaIni, editCabeza);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        // Descartar Edicion
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada", true, file);

                                        //Validacion Editar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, id, nomEspe, campoEdit, c, ErrorValidacion, "No se descartó correctamente", 1);

                                        // Editar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);

                                        CrudKbiednfo.CrudKBiednfo(desktopSession, 2, Lupas, nomInsti, nomEspe, fechaIni, editCabeza);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        // Aplicar Edicion
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        //Validacion Editar Logbo
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, id, "E", c, ErrorValidacion);

                                        //Validacion Editar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, id, editCabeza, campoEdit, c, ErrorValidacion, "No se editó correctamente", 1);

                                        rootSession = RootSession();
                                        rootSession.Mouse.MouseMove(rootSession.FindElementByName("Perfiles").Coordinates);
                                        rootSession.Mouse.Click(null);
                                        //Detalle
                                        Dictionary<string, Point> botonesNavegadorDetalle = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle = PruebaCRUD.findname(desktopSession, 30, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //Eliminar Detalle
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorDetalle, file);

                                        //Insertar Detalle
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Insertar"].X, botonesNavegadorDetalle["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Insertar Detalle", true, file);

                                        CrudKbiednfo.CrudKBiednfoDetalle(desktopSession, 1, tipo, req, espec, editDetalle);

                                        // Aplicar detalle
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Detalle Insertado", true, file);

                                        // Editar detalle
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle", true, file);

                                        CrudKbiednfo.CrudKBiednfoDetalle(desktopSession, 2, tipo, req, espec, editDetalle);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        // Descartar Edicion detalle
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Cancelar"].X, botonesNavegadorDetalle["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada", true, file);

                                        // Editar detalle
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle", true, file);

                                        CrudKbiednfo.CrudKBiednfoDetalle(desktopSession, 2, tipo, req, espec, editDetalle);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        // Aplicar Edicion detalle
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////SUMATORIA
                                        //errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //newFunctions_1.lapiz(desktopSession);


                                        //REPORTE PRELIMINAR
                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        ////Rectificar Pie Pagina PDF PRELIMINAR
                                        //listaResultado = Textopdf(pdf1, codProgram, user);
                                        //listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //REPORTE DINAMICO
                                        string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
                                        errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        ////Rectificar Pie Pagina PDF DINAMICO
                                        //listaResultado = Textopdf(pdf, codProgram, user);
                                        //listaResultado.ForEach(e => ErrorValidacion.Add(e));
                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //RECTIFICACION DE CAMPOS DE EXCEL       
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, id, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_1.lapiz(desktopSession);

                                        //Eliminar Detalle
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorDetalle, file);

                                        //Eliminar Cabecera
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);

                                        //Validacion Eliminar Logbo
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, id, "D", c, ErrorValidacion);

                                        //Validacion Eliminar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, id, "", Campo, c, ErrorValidacion, "No se eliminó correctamente", 2);
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
        public void KBiHvextChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosBI.ModulosBi_10.KBiHvextChecklist")
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
                                //Datos Login
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["id"].ToString().Length != 0 && rows["id"].ToString() != null &&
                                rows["actEmpr"].ToString().Length != 0 && rows["actEmpr"].ToString() != null &&
                                rows["nomEmpr"].ToString().Length != 0 && rows["nomEmpr"].ToString() != null &&
                                rows["fechaIng"].ToString().Length != 0 && rows["fechaIng"].ToString() != null &&
                                rows["editCabeza"].ToString().Length != 0 && rows["editCabeza"].ToString() != null &&
                                rows["pais"].ToString().Length != 0 && rows["pais"].ToString() != null &&
                                rows["mun"].ToString().Length != 0 && rows["mun"].ToString() != null &&
                                rows["tipo"].ToString().Length != 0 && rows["tipo"].ToString() != null &&
                                rows["req"].ToString().Length != 0 && rows["req"].ToString() != null &&
                                rows["espec"].ToString().Length != 0 && rows["espec"].ToString() != null &&
                                rows["editDetalle"].ToString().Length != 0 && rows["editDetalle"].ToString() != null &&
                                rows["codigo"].ToString().Length != 0 && rows["codigo"].ToString() != null &&
                                rows["editDetalle1"].ToString().Length != 0 && rows["editDetalle1"].ToString() != null &&
                                rows["campoEdit"].ToString().Length != 0 && rows["campoEdit"].ToString() != null &&
                                rows["herra"].ToString().Length != 0 && rows["herra"].ToString() != null &&
                                rows["observa"].ToString().Length != 0 && rows["observa"].ToString() != null &&
                                rows["nivel"].ToString().Length != 0 && rows["nivel"].ToString() != null &&
                                rows["cal"].ToString().Length != 0 && rows["cal"].ToString() != null &&
                                rows["editDetalle2"].ToString().Length != 0 && rows["editDetalle2"].ToString() != null)
                            {
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
                                string BanderaPreli = rows["BanderaPreli"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();

                                string id = rows["id"].ToString();
                                string actEmpr = rows["actEmpr"].ToString();
                                string nomEmpr = rows["nomEmpr"].ToString();
                                string fechaIng = rows["fechaIng"].ToString();
                                string editCabeza = rows["editCabeza"].ToString();
                                string pais = rows["pais"].ToString();
                                string dep = rows["dep"].ToString();
                                string mun = rows["mun"].ToString();
                                string tipo = rows["tipo"].ToString();
                                string req = rows["req"].ToString();
                                string espec = rows["espec"].ToString();
                                string editDetalle = rows["editDetalle"].ToString();
                                string codigo = rows["codigo"].ToString();
                                string editDetalle1 = rows["editDetalle1"].ToString();
                                string campoEdit = rows["campoEdit"].ToString();
                                string herra = rows["herra"].ToString();
                                string observa = rows["observa"].ToString();
                                string nivel = rows["nivel"].ToString();
                                string cal = rows["cal"].ToString();
                                string editDetalle2 = rows["editDetalle2"].ToString();

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "BI", motor);

                                try
                                {
                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();

                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
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


                                        //ELIMINAR REGISTROS PEGADOS DETALLE
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera("BI_ARHEX", user, motor, "COD_AREX", "1", "1", "COD_AREX", c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera("BI_PROYL", user, motor, "RMT_PROY", "1", "1", "RMT_PROY", c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera("BI_DTPHV", user, motor, "COD_EMPR", "9", "9", "COD_REQU", c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, id, id, Campo, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        // Insertar
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Insertar Registro", true, file);

                                        List<string> Lupas = new List<string> { id, actEmpr };
                                        CrudKbihvext.CrudKBihvext(desktopSession, 1, Lupas, nomEmpr, fechaIng, pais, dep, mun, editCabeza);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Datos Ingresados", true, file);

                                        // Aplicar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Registro Insertado", true, file);

                                        //Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, id, id, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        // Editar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);

                                        CrudKbihvext.CrudKBihvext(desktopSession, 2, Lupas, nomEmpr, fechaIng, pais, dep, mun, editCabeza);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        // Descartar Edicion
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada", true, file);

                                        //Validacion Editar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, id, nomEmpr, campoEdit, c, ErrorValidacion, "No se descartó correctamente", 1);

                                        // Editar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);

                                        CrudKbihvext.CrudKBihvext(desktopSession, 2, Lupas, nomEmpr, fechaIng, pais, dep, mun, editCabeza);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        // Aplicar Edicion
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        //Validacion Editar Logbo
                                       // ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, id, "E", c, ErrorValidacion);

                                        //Validacion Editar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, id, editCabeza, campoEdit, c, ErrorValidacion, "No se editó correctamente", 1);

                                        //Detalle 1
                                        desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Áreas de Experiencia").Coordinates);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);

                                        Dictionary<string, Point> botonesNavegadorDetalle = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle = PruebaCRUD.findname(desktopSession, 30, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //Insertar Detalle 1
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Insertar"].X, botonesNavegadorDetalle["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Insertar Detalle 1", true, file);

                                        CrudKbihvext.CrudKBihvextDetalle1(desktopSession, 1, codigo, editDetalle1);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Datos Ingresados Detalle 1", true, file);

                                        // Aplicar detalle 1
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Detalle 1 Insertado", true, file);

                                        // Editar detalle 1
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                        CrudKbihvext.CrudKBihvextDetalle1(desktopSession, 2, codigo, editDetalle1);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        // Descartar Edicion detalle 1
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Cancelar"].X, botonesNavegadorDetalle["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada Detalle 1", true, file);

                                        // Editar detalle 1
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                        CrudKbihvext.CrudKBihvextDetalle1(desktopSession, 2, codigo, editDetalle1);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        // Aplicar Edicion detalle 1
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 1", true, file);

                                        //Detalle 2
                                        desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Proyectos y Logros").Coordinates);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);

                                        Dictionary<string, Point> botonesNavegadorDetalle2 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle2 = PruebaCRUD.findname(desktopSession, 28, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //Insertar Detalle 2
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Insertar"].X, botonesNavegadorDetalle2["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Insertar Detalle 2", true, file);

                                        CrudKbihvext.CrudKBihvextDetalle2(desktopSession, 1, herra, nivel, cal, observa, editDetalle2);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 2", true, file);

                                        // Aplicar detalle 2
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Detalle 2 Insertado", true, file);

                                        // Editar detalle 2
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);

                                        CrudKbihvext.CrudKBihvextDetalle2(desktopSession, 2, herra, nivel, cal, observa, editDetalle2);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Campos Editados Detalle 2", true, file);

                                        // Descartar Edicion detalle 2
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Cancelar"].X, botonesNavegadorDetalle2["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada Detalle 2", true, file);

                                        // Editar detalle 2
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);

                                        CrudKbihvext.CrudKBihvextDetalle2(desktopSession, 2, herra, nivel, cal, observa, editDetalle2);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                        // Aplicar Edicion detalle 2
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 2", true, file);

                                        //Detalle 3
                                        desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Perfiles").Coordinates);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);

                                        Dictionary<string, Point> botonesNavegadorDetalle3 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle3 = PruebaCRUD.findname(desktopSession, 28, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //Eliminar Detalle 3
                                        //desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        if(motor == "SQL")
                                        {
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorDetalle3, file);
                                        }
                                        else
                                        {
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorDetalle3, file);
                                            for (int i = 0; i < 3; i++)
                                            {
                                                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle3["Borrar"].X, botonesNavegadorDetalle3["Borrar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                Thread.Sleep(1000);
                                                rootSession = RootSession();
                                                Thread.Sleep(1000);
                                                rootSession = ReloadSession(rootSession, "TMessageForm");
                                                Thread.Sleep(1000);
                                                rootSession.Mouse.MouseMove(rootSession.FindElementByName("OK").Coordinates);
                                                rootSession.Mouse.Click(null);
                                            }
                                        }
                                        Thread.Sleep(1000);

                                        //Insertar Detalle 3
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle3["Insertar"].X, botonesNavegadorDetalle3["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Insertar Detalle 3", true, file);

                                        CrudKbihvext.CrudKBihvextDetalle(desktopSession, 1, tipo, req, espec, editDetalle);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 3", true, file);

                                        // Aplicar detalle 3
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle3["Aplicar"].X, botonesNavegadorDetalle3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Detalle 3 Insertado", true, file);                                        

                                        // Editar detalle 3
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle3["Editar"].X, botonesNavegadorDetalle3["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 3", true, file);

                                        CrudKbihvext.CrudKBihvextDetalle(desktopSession, 2, tipo, req, espec, editDetalle);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Dastos Editados Detalle 3", true, file);

                                        // Descartar Edicion detalle 3
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle3["Cancelar"].X, botonesNavegadorDetalle3["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada Detalle 3", true, file);

                                        // Editar detalle 3
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle3["Editar"].X, botonesNavegadorDetalle3["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 3", true, file);

                                        CrudKbihvext.CrudKBihvextDetalle(desktopSession, 2, tipo, req, espec, editDetalle);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 3", true, file);

                                        // Aplicar Edicion detalle 3
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle3["Aplicar"].X, botonesNavegadorDetalle3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 3", true, file);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(500);
                                        //CrudKbicarpe.ventanaSum(desktopSession);
                                        //Thread.Sleep(500);
                                        

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
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, id, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_1.lapiz(desktopSession);
                                        //Debugger.Launch();
                                        //Eliminar Detalle 1
                                        desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Áreas de Experiencia").Coordinates);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorDetalle, file);
                                        Thread.Sleep(1000);

                                        /*//Eliminar Detalle 2
                                        desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Proyectos y Logros").Coordinates);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        
                                        var herramientas = rootSession.FindElementsByName("Herramientas");
                                        herramientas[1].Click();
                                        Thread.Sleep(1000);

                                        rootSession = RootSession();
                                        botonesNavegadorDetalle2 = PruebaCRUD.findname(rootSession, 28, 1);
                                        ElementList = PruebaCRUD.NavClass(rootSession);*/
                                        /*desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Borrar"].X, botonesNavegadorDetalle2["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Detalle 2", true, file);
                                        rootSession = ReloadSession(rootSession, "TMessageForm");
                                        Thread.Sleep(1000);
                                        rootSession.Mouse.MouseMove(rootSession.FindElementByName("OK").Coordinates);
                                        rootSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Detalle 2 Eliminado", true, file);*/

                                        //Eliminar Detalle 3
                                        desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Perfiles").Coordinates);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);

                                        rootSession = RootSession();
                                        botonesNavegadorDetalle3 = PruebaCRUD.findname(rootSession, 28, 1);
                                        ElementList = PruebaCRUD.NavClass(rootSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle3["Borrar"].X, botonesNavegadorDetalle3["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Detalle 3", true, file);
                                        Thread.Sleep(1000);
                                        rootSession = ReloadSession(rootSession, "TMessageForm");
                                        Thread.Sleep(1000);
                                        rootSession.Mouse.MouseMove(rootSession.FindElementByName("OK").Coordinates);
                                        rootSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Detalle 3 Eliminado", true, file);

                                        //Eliminar Cabecera
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        Thread.Sleep(1000);

                                        //Validacion Eliminar Logbo
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, id, "D", c, ErrorValidacion);

                                        //Validacion Eliminar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, id, "", Campo, c, ErrorValidacion, "No se eliminó correctamente", 2);
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

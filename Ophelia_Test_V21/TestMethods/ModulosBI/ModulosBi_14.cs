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
using OpenQA.Selenium.Appium;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas;

namespace Ophelia_Test_V21.TestMethods.ModulosBI
{

    [TestClass]
    public class ModulosBi_14 : FuncionesVitales
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
        public ModulosBi_14()
        {

        }

        private TestContext testContextInstance;

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
        /*
                [TestMethod]
                public void AbrirPrograma()
                {
                    string programa = "KBiParbi";
                    AbrirPrograma a = new AbrirPrograma(programa, "UCalida1");
                    desktopSession = a.Start2("SQL", "");
                    //AbrirPrograma a = new AbrirPrograma(programa, "TPRUEBAS");
                    //desktopSession = a.Start2("ORA", "");
                }

        */

        //[TestMethod]
        //public void KBilncinChecklist()
        //{
        //    //Debugger.Launch();
        //    List<string> errorMessages = new List<string>();
        //    List<string> listaResultado;
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

        //        if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosBI.ModulosBI_14.KBilncinChecKlist")
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
        //                        rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null)
        //                    // rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null )
        //                    /*
        //                                                    //Datos Generales
        //                                                    rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
        //                                                    rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
        //                                                    rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
        //                                                    rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
        //                                                    rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
        //                                                    rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null )
        //                                                    /*
        //                                                    // Validacion Tabla
        //                                                    rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
        //                                                    rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
        //                                                    rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&
        //                                                    //Nombre de los otros campos                        
        //                                                    rows["CampoFecResol"].ToString().Length != 0 && rows["CampoFecResol"].ToString() != null &&
        //                                                    rows["CampoTitulo"].ToString().Length != 0 && rows["CampoTitulo"].ToString() != null &&
        //                                                    // Editar los campos
        //                                                    rows["EditCampoFecResol"].ToString().Length != 0 && rows["EditCampoFecResol"].ToString() != null &&
        //                                                    rows["EditCampoTitulo"].ToString().Length != 0 && rows["EditCampoTitulo"].ToString() != null &&
        //                                                    //Datos Crud Principal
        //                                                    rows["NumResol"].ToString().Length != 0 && rows["NumResol"].ToString() != null &&
        //                                                    rows["FecResol"].ToString().Length != 0 && rows["FecResol"].ToString() != null &&
        //                                                    rows["TituloResol"].ToString().Length != 0 && rows["TituloResol"].ToString() != null &&
        //                                                    //Edicion de campos
        //                                                    rows["ActNumResol"].ToString().Length != 0 && rows["ActNumResol"].ToString() != null &&
        //                                                    rows["ActFecResol"].ToString().Length != 0 && rows["ActFecResol"].ToString() != null &&
        //                                                    rows["ActTituloResol"].ToString().Length != 0 && rows["ActTituloResol"].ToString() != null)
        //                    */
        //                    {

        //                        //Variables obligatorias
        //                        string user = rows["User"].ToString();
        //                        string motor = rows["Motor"].ToString();
        //                        string codProgram = rows["CodProgram"].ToString();
        //                        // string nomProgram = rows["NomProgram"].ToString();
        //                        /*
        //                                                        //Variables generales
        //                                                        string TipoQbe = rows["TipoQbe"].ToString();
        //                                                        string QbeFilter = rows["QbeFilter"].ToString();
        //                                                        string banExcel = rows["banExcel"].ToString();
        //                                                        string BanderaLupa = rows["BanderaLupa"].ToString();
        //                                                        string BanderaPreli = rows["BanderaPreli"].ToString();
        //                                                        string BanderaSum = rows["BanderaSum"].ToString();
        //                        /*
        //                                                        //Variables Validacion Tabla
        //                                                        string Tabla = rows["Tabla"].ToString();
        //                                                        string Campo = rows["Campo"].ToString();
        //                                                        string EditCampo = rows["EditCampo"].ToString();

        //                                                        //Variables Campos
        //                                                        string CampoFecResol = rows["CampoFecResol"].ToString();
        //                                                        string CampoTitulo = rows["CampoTitulo"].ToString();
        //                                                        //Variable Editar Campos
        //                                                        string EditCampoFecResol = rows["EditCampoFecResol"].ToString();
        //                                                        string EditCampoTitulo = rows["EditCampoTitulo"].ToString();

        //                                                        //Variables Ingresar Crud Principal
        //                                                        string NumResol = rows["NumResol"].ToString();
        //                                                        string FecResol = rows["FecResol"].ToString();
        //                                                        string TituloResol = rows["TituloResol"].ToString();

        //                                                        //Variables Editar Crud
        //                                                        string ActNumResol = rows["ActNumResol"].ToString();
        //                                                        string ActFecResol = rows["ActFecResol"].ToString();
        //                                                        string ActTituloResol = rows["ActTituloResol"].ToString();
        //                        */
        //                        //  string campo = "0";


        //                        //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
        //                        string file = CrearDocumentoWordDinamico(codProgram, "BI", motor);

        //                        try
        //                        {
        //                            List<string> ErrorValidacion = new List<string>();
        //                            List<string> errors = new List<string>();

        //                            //Abrir programa

        //                            AbrirPrograma a = new AbrirPrograma(codProgram, user);

        //                            //CRUD
        //                            List<string> crudPrincipal = new List<string>() { };

        //                            desktopSession = a.Start2(motor, "");

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

        //                                //VALIDACION CODIGO DEL PROGRAMA
        //                                errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                /*
        //                                                                        //VALIDACION NOMBRE DEL PROGRAMA
        //                                                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
        //                                                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                                                        /*     
        //                                                                              //VERSION
        //                                                                              FuncionesGlobales.GetVersion(desktopSession, file);
        //                                                                              newFunctions_4.ScreenshotNuevo("Consultar Version", true, file);
        //                                                                              Thread.Sleep(1000);

        //                                                                              //NOTAS
        //                                                                              newFunctions_4.openInnerNote(desktopSession, file);
        //                                                                              newFunctions_4.ScreenshotNuevo("Abrir notas", true, file);
        //                                                                              Thread.Sleep(1000);
        //                                */
        //                                /*
        //                                                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
        //                                                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
        //                                                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

        //                                                                        //Agregar
        //                                                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
        //                                                                        desktopSession.Mouse.Click(null);

        //                                                                        CrudKSlMreec.CrudMreec(desktopSession, "0", crudPrincipal, file);
        //                                                                        CrudKSlMreec.BuscarRegistro(desktopSession, 0);


        //                                                                        //Aplicar
        //                                                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
        //                                                                        desktopSession.Mouse.Click(null);
        //                                                                        newFunctions_4.ScreenshotNuevo("Datos Guardados", true, file);

        //                                                                        //LUPA
        //                                                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                                                        Thread.Sleep(1000);


        //                                                                        //Modificar     
        //                                                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
        //                                                                        desktopSession.Mouse.Click(null);

        //                                                                        CrudKSlMreec.CrudMreec(desktopSession, "1", crudPrincipal, file);
        //                                                                        CrudKSlMreec.BuscarRegistro(desktopSession, 1);
        //                                                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                                                        Thread.Sleep(1000);

        //                                                                        //Descartar edicion - Cancelar
        //                                                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
        //                                                                        desktopSession.Mouse.Click(null);
        //                                                                        newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);

        //                                                                        //Confirmar Modificar
        //                                                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
        //                                                                        desktopSession.Mouse.Click(null);

        //                                                                        CrudKSlMreec.CrudMreec(desktopSession, "1", crudPrincipal, file);
        //                                                                        CrudKSlMreec.BuscarRegistro(desktopSession, 1);
        //                                                                        newFunctions_4.ScreenshotNuevo("Editar Datos", true, file);

        //                                                                        //Volver a aplicar
        //                                                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
        //                                                                        desktopSession.Mouse.Click(null);
        //                                                                        newFunctions_4.ScreenshotNuevo("Confirmar Edicion", true, file);

        //                                                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
        //                                                                        desktopSession.Mouse.Click(null);
        //                                                                        newFunctions_4.ScreenshotNuevo("Eliminar Datos", true, file);

        //                                                                        //Aceptar Eliminar
        //                                                                        WindowsDriver<WindowsElement> btnAcep = null;
        //                                                                        btnAcep = PruebaCRUD.RootSession();
        //                                                                        btnAcep.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //                                                                        newFunctions_4.ScreenshotNuevo("Confirmar Eliminar Datos", true, file);



        //                                                                        //QBE
        //                                                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
        //                                                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                                                        /*                                
        //                                                                          //SUMATORIA
        //                                                                           errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
        //                                                                           if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                                                           newFunctions_1.lapiz(desktopSession);
        //                                                                        */
        //                                /*
        //                                                                        //REPORTE DINAMICO
        //                                                                        string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
        //                                                                        errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
        //                                                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                                                        //Rectificar Pie Pagina PDF DINAMICO
        //                                                                        listaResultado = Textopdf(pdf, codProgram, user);
        //                                                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));



        //                                                                        //REPORTE PRELIMINAR
        //                                                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
        //                                                                        errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
        //                                                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                                                        //Rectificar Pie Pagina PDF PRELIMINAR
        //                                                                        listaResultado = Textopdf(pdf1, codProgram, user);
        //                                                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));


        //                                                                        //EXPORTAR EXCEL
        //                                                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
        //                                                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
        //                                                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                                                        Thread.Sleep(2000);

        //                                */



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
        //    }
        //}



        [TestMethod]
        public void KBiIncinChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosBI.ModulosBi_14.KBiIncinChecklist")
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

                                //Variables Validaciones
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&

                                //Datos Cruds
                                rows["Nincidente"].ToString().Length != 0 && rows["Nincidente"].ToString() != null &&
                                rows["Responsable"].ToString().Length != 0 && rows["Responsable"].ToString() != null &&
                                rows["QuienReport"].ToString().Length != 0 && rows["QuienReport"].ToString() != null &&
                                rows["NumEsta"].ToString().Length != 0 && rows["NumEsta"].ToString() != null &&
                                rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&                                
                                rows["ActNumEsta"].ToString().Length != 0 && rows["ActNumEsta"].ToString() != null)

                                

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

                                //Variables validacion
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string EditCampo = rows["EditCampo"].ToString();


                                //Variables crud principal
                                string QuienReport = rows["QuienReport"].ToString();
                                string Nincidente = rows["Nincidente"].ToString();
                                string Responsable = rows["Responsable"].ToString();

                                //Variables Crud Det 1
                                //string ActNomDet1 = rows["ActNomDet1"].ToString();

                                //Variables Crud Det 2
                                string NumEsta = rows["NumEsta"].ToString();
                                string Fecha = rows["Fecha"].ToString();
                                string ActNumEsta = rows["ActNumEsta"].ToString();


                                String campo = "1";

                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { QuienReport, Nincidente, Responsable };

                                ////LISTA DETALLE 1
                                //List<string> crudDet1 = new List<string>() { ActNomDet1 };

                                //LISTA DETALLE 2
                                List<string> crudDet2 = new List<string>() { NumEsta, Fecha, ActNumEsta };


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "BI", motor);

                                 
                                try
                                {
                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();

                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
                                    //Debugger.Launch();
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

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        newFunctions_1.lapiz(desktopSession);
                                        Thread.Sleep(1000);


                                        // Botones Crud PRINCIPAL
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);


                                        //// Botones Crud  DETALLE 1
                                        Dictionary<string, Point> botonesNavegador3 = new Dictionary<string, Point>();
                                        botonesNavegador3 = PruebaCRUD.findname(desktopSession, 27, 2);
                                        var ElementList3 = PruebaCRUD.NavClass(desktopSession);



                                        //     CRUD PRINCIPAL SOLO PERMITE EDITAR.

                                        ////    DESCARTAR EDICION PRINCIPAL
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);
                                        
                                        // Ingresa Datos en el CRUD
                                        CrudKbiincin.CrudPrincipal(desktopSession, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);
                                        Thread.Sleep(1000);
                                        ////////////////////////////////////


                                        ////    ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        // Ingresa Datos en el CRUD
                                        CrudKbiincin.CrudPrincipal(desktopSession, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);
                                        Thread.Sleep(3000);
                                        ///////////////////////



                                        //////////                AGREGAR REGISTRO DETALLE 2      //////////                                  
                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador3["Insertar"].X, botonesNavegador3["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKbiincin.CrudDet2(desktopSession, 1, crudDet2);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos DETALLE 2", true, file);
                                        Thread.Sleep(1000);

                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador3["Aplicar"].X, botonesNavegador3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados DETALLE 2", true, file);
                                        Thread.Sleep(1500);
                                        //////////////////////////////


                                        //              LUPA 
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador3["Editar"].X, botonesNavegador3["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar DETALLE 2", true, file);

                                        CrudKbiincin.CrudDet2(desktopSession, 2, crudDet2);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados DETALLE 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador3["Cancelar"].X, botonesNavegador3["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados DETALLE 2", true, file);
                                        Thread.Sleep(2000);
                                        /////////////////////


                                        //LUPA 
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador3["Editar"].X, botonesNavegador3["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar DETALLE 2", true, file);

                                        CrudKbiincin.CrudDet2(desktopSession, 2, crudDet2);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados DETALLE 2", true, file);

                                        Thread.Sleep(2000);
                                        desktopSession.Mouse.MouseMove(ElementList3[2].Coordinates, botonesNavegador3["Aplicar"].X, botonesNavegador3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada DETALLE 2", true, file);
                                        Thread.Sleep(2000);
                                        ////////////////////


                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //REPORTE PRELIMINAR 1
                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = CrudKbiincin.ReportePreliminar1(desktopSession, BanderaPreli, file, pdf1);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));
                                        Thread.Sleep(1000);


                                        //REPORTE PRELIMINAR 2
                                        string pdf2 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = CrudKbiincin.ReportePreliminar2(desktopSession, BanderaPreli, file, pdf2);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf2, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));
                                        Thread.Sleep(1500);



                                        //REPORTE PRELIMINAR 3
                                        string pdf3 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = CrudKbiincin.ReportePreliminar3(desktopSession, BanderaPreli, file, pdf3);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf3, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));
                                        Thread.Sleep(1000);


                                        //REPORTE PRELIMINAR 4
                                        string pdf4 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = CrudKbiincin.ReportePreliminar4(desktopSession, BanderaPreli, file, pdf4);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf4, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));
                                        Thread.Sleep(1500);


                                        //REPORTE PRELIMINAR 5
                                        string pdf5 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = CrudKbiincin.ReportePreliminar5(desktopSession, BanderaPreli, file, pdf5);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf5, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));
                                        Thread.Sleep(1500);

                                        //REPORTE PRELIMINAR 6
                                        string pdf6 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = CrudKbiincin.ReportePreliminar6(desktopSession, BanderaPreli, file, pdf6);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf6, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));
                                        Thread.Sleep(1500);


                                        //REPORTE DINAMICO
                                        string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
                                        errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        newFunctions_1.lapiz(desktopSession);


                                        //Eliminar Registro Detalle 1
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList3[2], botonesNavegador3, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 1", true, file);
                                        Thread.Sleep(1000);

                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Principal", true, file);
                                        Thread.Sleep(1500);

                                        newFunctions_4.ScreenshotNuevo("Programa finalizado", true, file);


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
        public void KBilncarChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosBI.ModulosBi_14.KBilncarChecKlist")
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
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null)
                            /*
                                                            //Datos Generales
                                                            rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                                            rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                                            rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                                            rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                                            rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                                            rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&

                                                            // Validacion Tabla
                                                            rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                                            rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                                            rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&
                                                            //Nombre de los otros campos                        
                                                            rows["CampoFecResol"].ToString().Length != 0 && rows["CampoFecResol"].ToString() != null &&
                                                            rows["CampoTitulo"].ToString().Length != 0 && rows["CampoTitulo"].ToString() != null &&
                                                            // Editar los campos
                                                            rows["EditCampoFecResol"].ToString().Length != 0 && rows["EditCampoFecResol"].ToString() != null &&
                                                            rows["EditCampoTitulo"].ToString().Length != 0 && rows["EditCampoTitulo"].ToString() != null &&
                                                            //Datos Crud Principal
                                                            rows["NumResol"].ToString().Length != 0 && rows["NumResol"].ToString() != null &&
                                                            rows["FecResol"].ToString().Length != 0 && rows["FecResol"].ToString() != null &&
                                                            rows["TituloResol"].ToString().Length != 0 && rows["TituloResol"].ToString() != null &&
                                                            //Edicion de campos
                                                            rows["ActNumResol"].ToString().Length != 0 && rows["ActNumResol"].ToString() != null &&
                                                            rows["ActFecResol"].ToString().Length != 0 && rows["ActFecResol"].ToString() != null &&
                                                            rows["ActTituloResol"].ToString().Length != 0 && rows["ActTituloResol"].ToString() != null)
                            */
                            {

                                //Variables obligatorias
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                /*
                                                                //Variables generales
                                                                string TipoQbe = rows["TipoQbe"].ToString();
                                                                string QbeFilter = rows["QbeFilter"].ToString();
                                                                string banExcel = rows["banExcel"].ToString();
                                                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                                                string BanderaPreli = rows["BanderaPreli"].ToString();
                                                                string BanderaSum = rows["BanderaSum"].ToString();

                                                                //Variables Validacion Tabla
                                                                string Tabla = rows["Tabla"].ToString();
                                                                string Campo = rows["Campo"].ToString();
                                                                string EditCampo = rows["EditCampo"].ToString();

                                                                //Variables Campos
                                                                string CampoFecResol = rows["CampoFecResol"].ToString();
                                                                string CampoTitulo = rows["CampoTitulo"].ToString();
                                                                //Variable Editar Campos
                                                                string EditCampoFecResol = rows["EditCampoFecResol"].ToString();
                                                                string EditCampoTitulo = rows["EditCampoTitulo"].ToString();

                                                                //Variables Ingresar Crud Principal
                                                                string NumResol = rows["NumResol"].ToString();
                                                                string FecResol = rows["FecResol"].ToString();
                                                                string TituloResol = rows["TituloResol"].ToString();

                                                                //Variables Editar Crud
                                                                string ActNumResol = rows["ActNumResol"].ToString();
                                                                string ActFecResol = rows["ActFecResol"].ToString();
                                                                string ActTituloResol = rows["ActTituloResol"].ToString();
                                */
                                //  string campo = "0";


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "BI", motor);

                                try
                                {
                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();

                                    //Abrir programa

                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);

                                    //CRUD
                                    List<string> crudPrincipal = new List<string>() { };

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

                                        //VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        newFunctions_4.ScreenshotNuevo("Consultar Version", true, file);
                                        Thread.Sleep(1000);

                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        newFunctions_4.ScreenshotNuevo("Abrir notas", true, file);
                                        Thread.Sleep(1000);

                                        /*                                 Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                                                         botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                                                         var ElementList = PruebaCRUD.NavClass(desktopSession);
                                       */
                                        CrudKBiIncar.CrudIncar(desktopSession, "0", crudPrincipal, file);
                                        newFunctions_4.ScreenshotNuevo("Ingresar Datos", true, file);

                                        Thread.Sleep(1000);

                                        CrudKBiIncar.CrudIncar(desktopSession, "1", crudPrincipal, file);
                                        newFunctions_4.ScreenshotNuevo("Boton Aceptar", true, file);
                                        /*  

                                                                                CrudKSlMreec.CrudMreec(desktopSession, "1", crudPrincipal, file);
                                                                                CrudKSlMreec.BuscarRegistro(desktopSession, 1);
                                                                                newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                                                                Thread.Sleep(1000);



                                                                                //Aceptar Eliminar
                                                                                WindowsDriver<WindowsElement> btnAcep = null;
                                                                                btnAcep = PruebaCRUD.RootSession();
                                                                                btnAcep.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                                                                newFunctions_4.ScreenshotNuevo("Confirmar Eliminar Datos", true, file);


                                        */
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
            }

        }

        [TestMethod]
        public void KBiHincrChecklist()
        {
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

                    if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosBI.ModulosBi_14.KBiHincrChecklist")
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

                                    //Variables Validaciones
                                    rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                    rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                    rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&

                                    //Datos Crud
                                    rows["Cargo"].ToString().Length != 0 && rows["Cargo"].ToString() != null &&
                                    rows["Plazas"].ToString().Length != 0 && rows["Plazas"].ToString() != null &&
                                    rows["Resolu"].ToString().Length != 0 && rows["Resolu"].ToString() != null &&
                                    rows["Fech"].ToString().Length != 0 && rows["Fech"].ToString() != null &&
                                    rows["ActPlaza"].ToString().Length != 0 && rows["ActPlaza"].ToString() != null)

                                    

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

                                    //Variables validacion
                                    string Tabla = rows["Tabla"].ToString();
                                    string Campo = rows["Campo"].ToString();
                                    string EditCampo = rows["EditCampo"].ToString();


                                    //Variables crud principal
                                    string Cargo = rows["Cargo"].ToString();
                                    string Plazas = rows["Plazas"].ToString();
                                    string Resolu = rows["Resolu"].ToString();
                                    string Fech = rows["Fech"].ToString();
                                    string ActPlaza = rows["ActPlaza"].ToString();

                                    
                                    //String campo = "1";

                                    //LISTA CRUD PRINCIPAL
                                    List<string> crudPrincipal = new List<string>() { Cargo, Plazas, Resolu, Fech, ActPlaza, };


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


                                            // Botones Crud PRINCIPAL
                                            Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                            botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 0);
                                            var ElementList = PruebaCRUD.NavClass(desktopSession);

                                                            
                                            //////////           CRUD PRICIPAL    ////////// 
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                            desktopSession.Mouse.Click(null);

                                            CrudKBiHincr.CrudPrincipal(desktopSession, 1, crudPrincipal);
                                            newFunctions_4.ScreenshotNuevo("Insertar Datos ", true, file);
                                            Thread.Sleep(1000);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Datos Insertados ", true, file);
                                            Thread.Sleep(1500);
                                            //////////////////////////////


                                            ////              LUPA 
                                            //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                            ////DESCARTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar ", true, file);

                                            CrudKBiHincr.CrudPrincipal(desktopSession, 2, crudPrincipal);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados ", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados ", true, file);
                                            Thread.Sleep(2000);

                                            /////////////////////


                                            //LUPA 
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                            ////ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar ", true, file);

                                            CrudKBiHincr.CrudPrincipal(desktopSession, 2, crudPrincipal);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados ", true, file);

                                            Thread.Sleep(2000);
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);

                                            newFunctions_4.ScreenshotNuevo("Edición Aplicada ", true, file);
                                            Thread.Sleep(2000);
                                            ////////////////////


                                            //QBE
                                            errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                            //REPORTE DINAMICO
                                            string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
                                            errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                            //REPORTE PRELIMINAR
                                            string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                            errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                            //EXPORTAR EXCEL
                                            string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                            errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            Thread.Sleep(2000);


                                            //Eliminar Registro Detalle 1
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                            newFunctions_4.ScreenshotNuevo("Datos Eliminados ", true, file);
                                            Thread.Sleep(1000);

                                            newFunctions_4.ScreenshotNuevo("Programa finalizado", true, file);


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



        [TestMethod]
        public void KBiDatadChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosBI.ModulosBi_14.KBiDatadChecklist")
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

                                // Validacion Tabla
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&
                                //Nombre de los otros campos                        
                                rows["CampoEstatu"].ToString().Length != 0 && rows["CampoEstatu"].ToString() != null &&
                                rows["CampoPeso"].ToString().Length != 0 && rows["CampoPeso"].ToString() != null &&
                                rows["CampoMedPer"].ToString().Length != 0 && rows["CampoMedPer"].ToString() != null &&
                                rows["CampoTelMed"].ToString().Length != 0 && rows["CampoTelMed"].ToString() != null &&
                                rows["CampoEmpSeg"].ToString().Length != 0 && rows["CampoEmpSeg"].ToString() != null &&
                                rows["CampoNumSeg"].ToString().Length != 0 && rows["CampoNumSeg"].ToString() != null &&
                                rows["CampoEnfer"].ToString().Length != 0 && rows["CampoEnfer"].ToString() != null &&
                                rows["CampoEmerg"].ToString().Length != 0 && rows["CampoEmerg"].ToString() != null &&
                                // Editar los campos
                                rows["EditEstatu"].ToString().Length != 0 && rows["EditEstatu"].ToString() != null &&
                                rows["EditPeso"].ToString().Length != 0 && rows["EditPeso"].ToString() != null &&
                                rows["EditMedPer"].ToString().Length != 0 && rows["EditMedPer"].ToString() != null &&
                                rows["EditTelMed"].ToString().Length != 0 && rows["EditTelMed"].ToString() != null &&
                                rows["EditEmpSeg"].ToString().Length != 0 && rows["EditEmpSeg"].ToString() != null &&
                                rows["EditNumSeg"].ToString().Length != 0 && rows["EditNumSeg"].ToString() != null &&
                                rows["EditEnfer"].ToString().Length != 0 && rows["EditEnfer"].ToString() != null &&
                                rows["EditEmerg"].ToString().Length != 0 && rows["EditEmerg"].ToString() != null &&
                                //Datos Crud Principal
                                rows["CodEmp"].ToString().Length != 0 && rows["CodEmp"].ToString() != null &&
                                rows["Estatu"].ToString().Length != 0 && rows["Estatu"].ToString() != null &&
                                rows["Peso"].ToString().Length != 0 && rows["Peso"].ToString() != null &&
                                rows["MedPer"].ToString().Length != 0 && rows["MedPer"].ToString() != null &&
                                rows["TelMed"].ToString().Length != 0 && rows["TelMed"].ToString() != null &&
                                rows["EmpSeg"].ToString().Length != 0 && rows["EmpSeg"].ToString() != null &&
                                rows["NumSeg"].ToString().Length != 0 && rows["NumSeg"].ToString() != null &&
                                rows["Enfer"].ToString().Length != 0 && rows["Enfer"].ToString() != null &&
                                rows["Emerg"].ToString().Length != 0 && rows["Emerg"].ToString() != null &&
                                //Edicion de campos
                                rows["ActEstatu"].ToString().Length != 0 && rows["ActEstatu"].ToString() != null &&
                                rows["ActPeso"].ToString().Length != 0 && rows["ActPeso"].ToString() != null &&
                                rows["ActMedPer"].ToString().Length != 0 && rows["ActMedPer"].ToString() != null &&
                                rows["ActTelMed"].ToString().Length != 0 && rows["ActTelMed"].ToString() != null &&
                                rows["ActEmpSeg"].ToString().Length != 0 && rows["ActEmpSeg"].ToString() != null &&
                                rows["ActNumSeg"].ToString().Length != 0 && rows["ActNumSeg"].ToString() != null &&
                                rows["ActEnfer"].ToString().Length != 0 && rows["ActEnfer"].ToString() != null &&
                                rows["ActEmerg"].ToString().Length != 0 && rows["ActEmerg"].ToString() != null)

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

                                //Variables Validacion Tabla
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string EditCampo = rows["EditCampo"].ToString();

                                //Variables Campos
                                string CampoEstatu = rows["CampoEstatu"].ToString();
                                string CampoPeso = rows["CampoPeso"].ToString();
                                string CampoMedPer = rows["CampoMedPer"].ToString();
                                string CampoTelMed = rows["CampoTelMed"].ToString();
                                string CampoEmpSeg = rows["CampoEmpseg"].ToString();
                                string CampoNumSeg = rows["CampoNumSeg"].ToString();
                                string CampoEnfer = rows["CampoEnfer"].ToString();
                                string CampoEmerg = rows["CampoEmerg"].ToString();
                                //Variable Editar Campos
                                string EditEstatu = rows["EditEstatu"].ToString();
                                string EditPeso = rows["EditPeso"].ToString();
                                string EditMedPer = rows["EditMedPer"].ToString();
                                string EditTelmed = rows["EditTelMed"].ToString();
                                string EditEmpSeg = rows["EditEmpSeg"].ToString();
                                string EditNumSeg = rows["EditNumSeg"].ToString();
                                string EditEnfer = rows["EditEnfer"].ToString();
                                string EditEmerg = rows["EditEmerg"].ToString();

                                //Variables Ingresar Crud Principal
                                string CodEmp = rows["CodEmp"].ToString();
                                string Estatu = rows["Estatu"].ToString();
                                string Peso = rows["Peso"].ToString();
                                string MedPer = rows["MedPer"].ToString();
                                string TelMed = rows["TelMed"].ToString();
                                string EmpSeg = rows["EmpSeg"].ToString();
                                string NumSeg = rows["NumSeg"].ToString();
                                string Enfer = rows["Enfer"].ToString();
                                string Emerg = rows["Emerg"].ToString();

                                //Variables Editar Crud
                                string ActEstatu = rows["ActEstatu"].ToString();
                                string ActPeso = rows["ActPeso"].ToString();
                                string ActMedPer = rows["ActMedPer"].ToString();
                                string ActTelMed = rows["TelMed"].ToString();
                                string ActEmpSeg = rows["ActEmpSeg"].ToString();
                                string ActNumSeg = rows["ActNumSeg"].ToString();
                                string ActEnfer = rows["Enfer"].ToString();
                                string ActEmerg = rows["ActEmerg"].ToString();

                                string campo = "0";


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "BI", motor);

                                try
                                {
                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();

                                    //Abrir programa

                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);

                                    //CRUD
                                    List<string> crudPrincipal = new List<string>() {CodEmp, Estatu, Peso, MedPer, TelMed, EmpSeg, NumSeg,
                                                                                     Enfer, Emerg, ActEstatu, ActPeso, ActMedPer, ActTelMed, ActEmpSeg, ActNumSeg, ActEnfer, ActEmerg };

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


                                        //VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        newFunctions_4.ScreenshotNuevo("Consultar Version", true, file);
                                        Thread.Sleep(1000);

                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        newFunctions_4.ScreenshotNuevo("Abrir notas", true, file);
                                        Thread.Sleep(1000);



                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //Agregar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        CrudKBiDatad.CrudData(desktopSession, "0", crudPrincipal, file);
                                        newFunctions_4.ScreenshotNuevo("Ingresar Datos", true, file);

                                        //Aplicar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Guardados", true, file);

                                        //LUPA
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        Thread.Sleep(1000);


                                        //Modificar     
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        CrudKBiDatad.CrudData(desktopSession, "1", crudPrincipal, file);
                                        CrudKBiDatad.CrudData(desktopSession, "2", crudPrincipal, file);

                                        Thread.Sleep(1000);

                                        //Descartar edicion - Cancelar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);

                                        //Confirmar Modificar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKBiDatad.CrudData(desktopSession, "1", crudPrincipal, file);
                                        CrudKBiDatad.CrudData(desktopSession, "2", crudPrincipal, file);
                                        newFunctions_4.ScreenshotNuevo("Editar Datos", true, file);

                                        //Volver a aplicar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Confirmar Edicion", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Datos", true, file);

                                        //Aceptar Eliminar
                                        WindowsDriver<WindowsElement> btnAcep = null;
                                        btnAcep = PruebaCRUD.RootSession();
                                        btnAcep.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                        newFunctions_4.ScreenshotNuevo("Confirmar Eliminar Datos", true, file);



                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        /*                                
                                          //SUMATORIA
                                           errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                           if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                           newFunctions_1.lapiz(desktopSession);
                                        */

                                        //REPORTE DINAMICO
                                        string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
                                        errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF DINAMICO
                                        listaResultado = Textopdf(pdf, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        CrudKBiDatad.reportePreliminar(desktopSession, 0, file);

                                        ////REPORTE PRELIMINAR

                                        //CrudKBiDatad.reportePreliminar(desktopSession, 0, file);

                                        //REPORTE PRELIMINAR
                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }

                                        Thread.Sleep(1000);

                                        CrudKBiDatad.reportePreliminar(desktopSession, 1, file);

                                        //REPORTE PRELIMINAR
                                        string pdf2 = @"C:\Reportes\ReportePDF2_" + "Preliminar" + "_" + Hora();
                                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf2, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }

                                        Thread.Sleep(1000);


                                        CrudKBiDatad.reportePreliminar(desktopSession, 2, file);

                                        //REPORTE PRELIMINAR
                                        string pdf3 = @"C:\Reportes\ReportePDF3_" + "Preliminar" + "_" + Hora();
                                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf3, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }

                                        Thread.Sleep(1000);

                                        //Debugger.Launch();
                                        CrudKBiDatad.reportePreliminar(desktopSession, 3, file);
                                        Thread.Sleep(2000);

                                        //REPORTE PRELIMINAR
                                        string pdf4 = @"C:\Reportes\ReportePDF4_" + "Preliminar" + "_" + Hora();
                                        Thread.Sleep(2000);
                                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf4, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                                        //Debugger.Launch();
                                        Thread.Sleep(2000);


                                        CrudKBiDatad.reportePreliminar(desktopSession, 4, file);

                                        Thread.Sleep(2000);

                                        //REPORTE PRELIMINAR
                                        string pdf5 = @"C:\Reportes\ReportePDF5_" + "Preliminar" + "_" + Hora();
                                        Thread.Sleep(2000);
                                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf5, nomProgram);
                                        Thread.Sleep(2000);
                                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }

                                        Thread.Sleep(1000);

                                        CrudKBiDatad.reportePreliminar(desktopSession, 5, file);
                                        Thread.Sleep(2000);


                                        //REPORTE PRELIMINAR
                                        string pdf6 = @"C:\Reportes\ReportePDF6_" + "Preliminar" + "_" + Hora();
                                        Thread.Sleep(2000);
                                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf6, nomProgram);
                                        Thread.Sleep(2000);
                                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }




                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(2000);

                                        ////Eliminar
                                        //PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        //newFunctions_4.ScreenshotNuevo("Datos Eliminados Principal", true, file);



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
            }
        }

        [TestMethod]
        public void KBiVliscCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosBI.ModulosBi_14.KBiVliscChecKlist")
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

                                //Datos Obligatorios
                                rows["Campo1"].ToString().Length != 0 && rows["Campo1"].ToString() != null &&
                                rows["Campo2"].ToString().Length != 0 && rows["Campo2"].ToString() != null &&


                                //Datos Generales
                                //
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null)
                            
                            {

                                //Variables obligatorias
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();

                                //Variables generales                                
                                string BanderaPreli = rows["BanderaPreli"].ToString();
                                
                                ////Variables Validacion Tabla
                                //string Tabla = rows["Tabla"].ToString();
                                //string Campo = rows["Campo"].ToString();
                                //string EditCampo = rows["EditCampo"].ToString();

                                                                
                                //Variables Ingresar Crud Principal
                                string Campo1 = rows["Campo1"].ToString();
                                string Campo2 = rows["Campo2"].ToString();


                                //  string campo = "0";

                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { Campo1, Campo2 };


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
                                        //VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        //VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        newFunctions_4.ScreenshotNuevo("Consultar Version", true, file);
                                        Thread.Sleep(1000);


                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        newFunctions_4.ScreenshotNuevo("Abrir notas", true, file);
                                        Thread.Sleep(1000);


                                        CrudKBivlisc.CrudVlisc(desktopSession, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Ingresar Datos", true, file);
                                        Thread.Sleep(1000);


                                        CrudKBivlisc.ClickAceptar(desktopSession);
                                        newFunctions_4.ScreenshotNuevo("Boton Aceptar", true, file);
                                        Thread.Sleep(1000);
;

                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                                        newFunctions_4.ScreenshotNuevo("Programa Finalizado", true, file);                                       

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
            }
        }



        [TestCleanup]
        public void Limpiar()
        {
            if (desktopSession != null)
            {
                desktopSession.Close();
                desktopSession.Dispose();
            }


        }





    }
}


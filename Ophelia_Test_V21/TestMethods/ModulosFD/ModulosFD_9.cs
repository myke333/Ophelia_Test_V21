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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosFd;
using OpenQA.Selenium.Appium;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas;

namespace Ophelia_Test_V21.TestMethods.ModulosFD
{

    [TestClass]
    public class ModulosFD_9 : FuncionesVitales
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
        public ModulosFD_9()
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

        [TestMethod]
        public void KFdPmaleChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosFD.ModulosFD_9.KFdPmaleChecklist")
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
                                string file = CrearDocumentoWordDinamico(codProgram, "FD", motor);

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

                                        CrudKfdpmale.CrudPmaleChecklist(desktopSession, 0);


                                        /*
                                                                                //VALIDACION NOMBRE DEL PROGRAMA
                                                                                errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                                                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                                /*                                    
                                                                                //VERSION
                                                                                FuncionesGlobales.GetVersion(desktopSession, file);
                                                                                newFunctions_4.ScreenshotNuevo("Consultar Version", true, file);
                                                                                Thread.Sleep(1000);

                                                                                //NOTAS
                                                                                newFunctions_4.openInnerNote(desktopSession, file);
                                                                                newFunctions_4.ScreenshotNuevo("Abrir notas", true, file);
                                                                                Thread.Sleep(1000);
                                                */


                                        /*
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
        */

                                        /*



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
        public void KFdHforoChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosFD.ModulosFD_9.KFdHforoChecklist")
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
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null)
                            /*
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

                                //Variables generales
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();
                                /*
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
                                string campo = "0";


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "FD", motor);

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



                                        /*
                                                                                //VALIDACION NOMBRE DEL PROGRAMA
                                                                                errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                                                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                                /*                                    
                                                                                //VERSION
                                                                                FuncionesGlobales.GetVersion(desktopSession, file);
                                                                                newFunctions_4.ScreenshotNuevo("Consultar Version", true, file);
                                                                                Thread.Sleep(1000);

                                                                                //NOTAS
                                                                                newFunctions_4.openInnerNote(desktopSession, file);
                                                                                newFunctions_4.ScreenshotNuevo("Abrir notas", true, file);
                                                                                Thread.Sleep(1000);
                                                */


                                        /*
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
        */

                                        /*



                                                                                Thread.Sleep(1000);



                                                                                //Aceptar Eliminar
                                                                                WindowsDriver<WindowsElement> btnAcep = null;
                                                                                btnAcep = PruebaCRUD.RootSession();
                                                                                btnAcep.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                                                                newFunctions_4.ScreenshotNuevo("Confirmar Eliminar Datos", true, file);


                                        */

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

    }
}

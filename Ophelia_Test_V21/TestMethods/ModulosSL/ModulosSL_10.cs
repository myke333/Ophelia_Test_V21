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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSl;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_10;
using OpenQA.Selenium.Appium;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosEd;

namespace Ophelia_Test_V21.TestMethods.ModulosSL
{

    [TestClass]
    public class ModulosSL_10 : FuncionesVitales
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
        public ModulosSL_10()
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
        public void KSlAppruChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSL.ModulosSL_10.KSlAppruChecklist")
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
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null)
                                //rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null)

                                //// Validacion Tabla
                                //rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                //rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                //rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&
                                ////Nombre de los otros campos                        
                                //rows["CampoFecResol"].ToString().Length != 0 && rows["CampoFecResol"].ToString() != null &&
                                //rows["CampoTitulo"].ToString().Length != 0 && rows["CampoTitulo"].ToString() != null &&
                                //// Editar los campos
                                //rows["EditCampoFecResol"].ToString().Length != 0 && rows["EditCampoFecResol"].ToString() != null &&
                                //rows["EditCampoTitulo"].ToString().Length != 0 && rows["EditCampoTitulo"].ToString() != null &&
                                ////Datos Crud Principal
                                //rows["NumResol"].ToString().Length != 0 && rows["NumResol"].ToString() != null &&
                                //rows["FecResol"].ToString().Length != 0 && rows["FecResol"].ToString() != null &&
                                //rows["TituloResol"].ToString().Length != 0 && rows["TituloResol"].ToString() != null &&
                                ////Edicion de campos
                                //rows["ActNumResol"].ToString().Length != 0 && rows["ActNumResol"].ToString() != null &&
                                //rows["ActFecResol"].ToString().Length != 0 && rows["ActFecResol"].ToString() != null &&
                                //rows["ActTituloResol"].ToString().Length != 0 && rows["ActTituloResol"].ToString() != null)

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
                                //string BanderaSum = rows["BanderaSum"].ToString();

                                ////Variables Validacion Tabla
                                //string Tabla = rows["Tabla"].ToString();
                                //string Campo = rows["Campo"].ToString();
                                //string EditCampo = rows["EditCampo"].ToString();

                                ////Variables Campos
                                //string CampoFecResol = rows["CampoFecResol"].ToString();
                                //string CampoTitulo = rows["CampoTitulo"].ToString();
                                ////Variable Editar Campos
                                //string EditCampoFecResol = rows["EditCampoFecResol"].ToString();
                                //string EditCampoTitulo = rows["EditCampoTitulo"].ToString();

                                ////Variables Ingresar Crud Principal
                                //string NumResol = rows["NumResol"].ToString();
                                //string FecResol = rows["FecResol"].ToString();
                                //string TituloResol = rows["TituloResol"].ToString();

                                ////Variables Editar Crud
                                //string ActNumResol = rows["ActNumResol"].ToString();
                                //string ActFecResol = rows["ActFecResol"].ToString();
                                //string ActTituloResol = rows["ActTituloResol"].ToString();

                                string campo = "0";


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "SL", motor);

                                try
                                {
                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();

                                    //Abrir programa

                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);

                                    //CRUD
                                    List<string> crudPrincipal = new List<string>() {};

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

                                        ////VALIDACION CODIGO DEL PROGRAMA
                                        //errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

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


                                        //Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        //botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        //var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        ////Agregar
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        //desktopSession.Mouse.Click(null);

                                        //CrudKSlMreec.CrudMreec(desktopSession, "0", crudPrincipal, file);
                                        //CrudKSlMreec.BuscarRegistro(desktopSession, 0);


                                        ////Aplicar
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //newFunctions_4.ScreenshotNuevo("Datos Guardados", true, file);

                                        ////              LUPA 
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //Thread.Sleep(1000);


                                        ////Modificar     
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        //desktopSession.Mouse.Click(null);

                                        //CrudKSlMreec.CrudMreec(desktopSession, "1", crudPrincipal, file);
                                        //CrudKSlMreec.BuscarRegistro(desktopSession, 1);
                                        //newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        //Thread.Sleep(1000);

                                        ////Descartar edicion - Cancelar
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);

                                        ////Confirmar Modificar
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        //desktopSession.Mouse.Click(null);

                                        //CrudKSlMreec.CrudMreec(desktopSession, "1", crudPrincipal, file);
                                        //CrudKSlMreec.BuscarRegistro(desktopSession, 1);
                                        //newFunctions_4.ScreenshotNuevo("Editar Datos", true, file);

                                        ////Volver a aplicar
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //newFunctions_4.ScreenshotNuevo("Confirmar Edicion", true, file);

                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //newFunctions_4.ScreenshotNuevo("Eliminar Datos", true, file);

                                        ////Aceptar Eliminar
                                        //WindowsDriver<WindowsElement> btnAcep = null;
                                        //btnAcep = PruebaCRUD.RootSession();
                                        //btnAcep.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                        //newFunctions_4.ScreenshotNuevo("Confirmar Eliminar Datos", true, file);



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

        [TestMethod]
        public void KSlPuofeChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSL.ModulosSL_10.KSlPuofeChecklist")
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
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null)
                            //rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null)

                            //// Validacion Tabla
                            //rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                            //rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                            //rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&
                            ////Nombre de los otros campos                        
                            //rows["CampoFecResol"].ToString().Length != 0 && rows["CampoFecResol"].ToString() != null &&
                            //rows["CampoTitulo"].ToString().Length != 0 && rows["CampoTitulo"].ToString() != null &&
                            //// Editar los campos
                            //rows["EditCampoFecResol"].ToString().Length != 0 && rows["EditCampoFecResol"].ToString() != null &&
                            //rows["EditCampoTitulo"].ToString().Length != 0 && rows["EditCampoTitulo"].ToString() != null &&
                            ////Datos Crud Principal
                            //rows["NumResol"].ToString().Length != 0 && rows["NumResol"].ToString() != null &&
                            //rows["FecResol"].ToString().Length != 0 && rows["FecResol"].ToString() != null &&
                            //rows["TituloResol"].ToString().Length != 0 && rows["TituloResol"].ToString() != null &&
                            ////Edicion de campos
                            //rows["ActNumResol"].ToString().Length != 0 && rows["ActNumResol"].ToString() != null &&
                            //rows["ActFecResol"].ToString().Length != 0 && rows["ActFecResol"].ToString() != null &&
                            //rows["ActTituloResol"].ToString().Length != 0 && rows["ActTituloResol"].ToString() != null)

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
                                //string BanderaSum = rows["BanderaSum"].ToString();

                                ////Variables Validacion Tabla
                                //string Tabla = rows["Tabla"].ToString();
                                //string Campo = rows["Campo"].ToString();
                                //string EditCampo = rows["EditCampo"].ToString();

                                ////Variables Campos
                                //string CampoFecResol = rows["CampoFecResol"].ToString();
                                //string CampoTitulo = rows["CampoTitulo"].ToString();
                                ////Variable Editar Campos
                                //string EditCampoFecResol = rows["EditCampoFecResol"].ToString();
                                //string EditCampoTitulo = rows["EditCampoTitulo"].ToString();

                                ////Variables Ingresar Crud Principal
                                //string NumResol = rows["NumResol"].ToString();
                                //string FecResol = rows["FecResol"].ToString();
                                //string TituloResol = rows["TituloResol"].ToString();

                                ////Variables Editar Crud
                                //string ActNumResol = rows["ActNumResol"].ToString();
                                //string ActFecResol = rows["ActFecResol"].ToString();
                                //string ActTituloResol = rows["ActTituloResol"].ToString();

                                string campo = "0";


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "SL", motor);

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


                                        //Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        //botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        //var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        ////Agregar
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        //desktopSession.Mouse.Click(null);

                                        //CrudKSlMreec.CrudMreec(desktopSession, "0", crudPrincipal, file);
                                        //CrudKSlMreec.BuscarRegistro(desktopSession, 0);


                                        ////Aplicar
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //newFunctions_4.ScreenshotNuevo("Datos Guardados", true, file);

                                        //LUPA
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        Thread.Sleep(1000);


                                        ////Modificar     
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        //desktopSession.Mouse.Click(null);

                                        //CrudKSlMreec.CrudMreec(desktopSession, "1", crudPrincipal, file);
                                        //CrudKSlMreec.BuscarRegistro(desktopSession, 1);
                                        //newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        //Thread.Sleep(1000);

                                        ////Descartar edicion - Cancelar
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);

                                        ////Confirmar Modificar
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        //desktopSession.Mouse.Click(null);

                                        //CrudKSlMreec.CrudMreec(desktopSession, "1", crudPrincipal, file);
                                        //CrudKSlMreec.BuscarRegistro(desktopSession, 1);
                                        //newFunctions_4.ScreenshotNuevo("Editar Datos", true, file);

                                        ////Volver a aplicar
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //newFunctions_4.ScreenshotNuevo("Confirmar Edicion", true, file);

                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //newFunctions_4.ScreenshotNuevo("Eliminar Datos", true, file);

                                        ////Aceptar Eliminar
                                        //WindowsDriver<WindowsElement> btnAcep = null;
                                        //btnAcep = PruebaCRUD.RootSession();
                                        //btnAcep.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                        //newFunctions_4.ScreenshotNuevo("Confirmar Eliminar Datos", true, file);



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

        [TestMethod]
        public void KSlOfesaChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSL.ModulosSL_10.KSlOfesaChecklist")
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


                                //Datos Crud Principal
                                rows["CodArea"].ToString().Length != 0 && rows["CodArea"].ToString() != null &&
                                rows["NomArea"].ToString().Length != 0 && rows["NomArea"].ToString() != null &&
                                rows["CodAreaPa"].ToString().Length != 0 && rows["CodAreaPa"].ToString() != null &&
                                rows["ActNomArea"].ToString().Length != 0 && rows["ActNomArea"].ToString() != null &&

                                //Datos Crud Detalle 1
                                rows["EmpLid"].ToString().Length != 0 && rows["EmpLid"].ToString() != null &&

                                rows["Lider"].ToString().Length != 0 && rows["Lider"].ToString() != null &&
                                rows["ActLider"].ToString().Length != 0 && rows["ActLider"].ToString() != null &&

                                //Variables Validaciones Crud principal
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&

                                //Variables Validaciones detalle 1
                                rows["TablaDet"].ToString().Length != 0 && rows["TablaDet"].ToString() != null &&
                                rows["CampoDet"].ToString().Length != 0 && rows["CampoDet"].ToString() != null &&
                                rows["EditCampoDet"].ToString().Length != 0 && rows["EditCampoDet"].ToString() != null)

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
                                string CodArea = rows["CodArea"].ToString();
                                string NomArea = rows["NomArea"].ToString();
                                string CodAreaPa = rows["CodAreaPa"].ToString();
                                string ActNomArea = rows["ActNomArea"].ToString();

                                //Variables Crud Det 1
                                string EmpLid = rows["EmpLid"].ToString();
                                string Lider = rows["Lider"].ToString();
                                string ActLider = rows["ActLider"].ToString();

                                //Variables validacion  crud Principal
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string EditCampo = rows["EditCampo"].ToString();

                                //Variables validacion DETALLE 1
                                string TablaDet = rows["TablaDet"].ToString();
                                string CampoDet = rows["CampoDet"].ToString();
                                string EditCampoDet = rows["EditCampoDet"].ToString();


                                //String campo = "6";

                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { CodArea, NomArea, CodAreaPa, ActNomArea };
                                //LISTA DETALLE 1
                                List<string> crudDet1 = new List<string>() { EmpLid, Lider, ActLider };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "GE", motor);

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

                                        Thread.Sleep(1300);

                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 27, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        Thread.Sleep(2000);

                                        Dictionary<string, Point> botonesNavegador2 = new Dictionary<string, Point>();
                                        botonesNavegador2 = PruebaCRUD.findname(desktopSession, 27, 1);
                                        var ElementList2 = PruebaCRUD.NavClass(desktopSession);

                                        Thread.Sleep(5000);


                                        //AGREGAR REGISTRO Principal
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1300);


                                        //CrudKgeareas.AgregarCrud(desktopSession, 0, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);


                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);
                                        Thread.Sleep(1500);


                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);
                                        Thread.Sleep(500);


                                        //CrudKgeareas.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);
                                        Thread.Sleep(500);


                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        Thread.Sleep(500);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);
                                        Thread.Sleep(500);


                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        Thread.Sleep(500);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        //CrudKgeareas.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);
                                        Thread.Sleep(2000);


                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        ////AGREGAR REGISTRO DETALLE 1                        
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Insertar"].X, botonesNavegador2["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);


                                        //CrudKgeareas.AgregarCrudDet(desktopSession, 0, crudDet1);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);
                                        Thread.Sleep(1500);


                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        //CrudKgeareas.AgregarCrudDet(desktopSession, 1, crudDet1);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Cancelar"].X, botonesNavegador2["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);
                                        Thread.Sleep(1000);


                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        //CrudKgeareas.AgregarCrudDet(desktopSession, 1, crudDet1);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);
                                        Thread.Sleep(3000);


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
                                        Thread.Sleep(1500);

                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList2[1], botonesNavegador2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 1", true, file);
                                        Thread.Sleep(1500);


                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        //newFunctions_4.ScreenshotNuevo("Datos Eliminados Principal", true, file);
                                        Thread.Sleep(1500);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Principal", true, file);
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
        //public void KSlpacofChecKlist()
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

        //        if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosGE.ModulosGE_2.KGeAreasChecKlist")
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
        //                        rows["CodArea"].ToString().Length != 0 && rows["CodArea"].ToString() != null &&
        //                        rows["NomArea"].ToString().Length != 0 && rows["NomArea"].ToString() != null &&
        //                        rows["CodAreaPa"].ToString().Length != 0 && rows["CodAreaPa"].ToString() != null &&
        //                        rows["ActNomArea"].ToString().Length != 0 && rows["ActNomArea"].ToString() != null &&

        //                        //Datos Crud Detalle 1
        //                        rows["EmpLid"].ToString().Length != 0 && rows["EmpLid"].ToString() != null &&

        //                        rows["Lider"].ToString().Length != 0 && rows["Lider"].ToString() != null &&
        //                        rows["ActLider"].ToString().Length != 0 && rows["ActLider"].ToString() != null &&

        //                        //Variables Validaciones Crud principal
        //                        rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
        //                        rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
        //                        rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&

        //                        //Variables Validaciones detalle 1
        //                        rows["TablaDet"].ToString().Length != 0 && rows["TablaDet"].ToString() != null &&
        //                        rows["CampoDet"].ToString().Length != 0 && rows["CampoDet"].ToString() != null &&
        //                        rows["EditCampoDet"].ToString().Length != 0 && rows["EditCampoDet"].ToString() != null)

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


        //                        //Variables crud principal
        //                        string CodArea = rows["CodArea"].ToString();
        //                        string NomArea = rows["NomArea"].ToString();
        //                        string CodAreaPa = rows["CodAreaPa"].ToString();
        //                        string ActNomArea = rows["ActNomArea"].ToString();

        //                        //Variables Crud Det 1
        //                        string EmpLid = rows["EmpLid"].ToString();
        //                        string Lider = rows["Lider"].ToString();
        //                        string ActLider = rows["ActLider"].ToString();

        //                        //Variables validacion  crud Principal
        //                        string Tabla = rows["Tabla"].ToString();
        //                        string Campo = rows["Campo"].ToString();
        //                        string EditCampo = rows["EditCampo"].ToString();

        //                        //Variables validacion DETALLE 1
        //                        string TablaDet = rows["TablaDet"].ToString();
        //                        string CampoDet = rows["CampoDet"].ToString();
        //                        string EditCampoDet = rows["EditCampoDet"].ToString();


        //                        //String campo = "6";

        //                        //LISTA CRUD PRINCIPAL
        //                        List<string> crudPrincipal = new List<string>() { CodArea, NomArea, CodAreaPa, ActNomArea };
        //                        //LISTA DETALLE 1
        //                        List<string> crudDet1 = new List<string>() { EmpLid, Lider, ActLider };

        //                        //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
        //                        string file = CrearDocumentoWordDinamico(codProgram, "GE", motor);

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

        //                                Thread.Sleep(1300);

        //                                Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
        //                                botonesNavegador = PruebaCRUD.findname(desktopSession, 27, 0);
        //                                var ElementList = PruebaCRUD.NavClass(desktopSession);

        //                                Thread.Sleep(2000);

        //                                Dictionary<string, Point> botonesNavegador2 = new Dictionary<string, Point>();
        //                                botonesNavegador2 = PruebaCRUD.findname(desktopSession, 27, 1);
        //                                var ElementList2 = PruebaCRUD.NavClass(desktopSession);

        //                                Thread.Sleep(5000);


        //                                //AGREGAR REGISTRO Principal
        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                Thread.Sleep(1300);


        //                                //CrudKgeareas.AgregarCrud(desktopSession, 0, crudPrincipal);
        //                                //newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);


        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);
        //                                Thread.Sleep(1500);


        //                                //LUPA
        //                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


        //                                ////DESCARTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Editar", true, file);
        //                                Thread.Sleep(500);


        //                                //CrudKgeareas.AgregarCrud(desktopSession, 1, crudPrincipal);
        //                                //newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);
        //                                //Thread.Sleep(500);


        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
        //                                Thread.Sleep(500);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);
        //                                Thread.Sleep(500);


        //                                //LUPA
        //                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


        //                                ////ACEPTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
        //                                Thread.Sleep(500);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                //CrudKgeareas.AgregarCrud(desktopSession, 1, crudPrincipal);
        //                                //newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);
        //                                Thread.Sleep(2000);


        //                                //LUPA
        //                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


        //                                ////AGREGAR REGISTRO DETALLE 1                        
        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Insertar"].X, botonesNavegador2["Insertar"].Y);
        //                                desktopSession.Mouse.Click(null);


        //                                //CrudKgeareas.AgregarCrudDet(desktopSession, 0, crudDet1);
        //                                //newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);
        //                                Thread.Sleep(1500);


        //                                //LUPA
        //                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


        //                                ////DESCARTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                //CrudKgeareas.AgregarCrudDet(desktopSession, 1, crudDet1);
        //                                //newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Cancelar"].X, botonesNavegador2["Cancelar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);
        //                                Thread.Sleep(1000);


        //                                //LUPA
        //                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                ////ACEPTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                //CrudKgeareas.AgregarCrudDet(desktopSession, 1, crudDet1);
        //                                //newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);
        //                                Thread.Sleep(3000);


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
        //                                Thread.Sleep(1500);

        //                                PruebaCRUD.EliminarRegistro(desktopSession, ElementList2[1], botonesNavegador2, file);
        //                                newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 1", true, file);
        //                                Thread.Sleep(1500);


        //                                PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
        //                                //newFunctions_4.ScreenshotNuevo("Datos Eliminados Principal", true, file);
        //                                Thread.Sleep(1500);
        //                                newFunctions_4.ScreenshotNuevo("Datos Eliminados Principal", true, file);
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
        public void KSlNinhaChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSL.ModulosSL_10.KSlNinhaChecklist")
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
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null&&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null&&

                                //Variables Validaciones Crud Principal
                                rows["Grup"].ToString().Length != 0 && rows["Grup"].ToString() != null &&
                                rows["Req"].ToString().Length != 0 && rows["Req"].ToString() != null &&
                                rows["Id"].ToString().Length != 0 && rows["Id"].ToString() != null &&
                                

                                //Variables Validaciones Detalle 1
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null )
                                //rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null)


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
                                string Grup = rows["Grup"].ToString();
                                string Req = rows["Req"].ToString();
                                string Id = rows["Id"].ToString();
                                
                                //Variables Validaciones Detalle 1
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                //string EditCampo = rows["EditCampo"].ToString();

                                string campo = "1";

                                //CRUD
                                List<string> crudPrincipal = new List<string>() {Grup, Req, Id };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "SL", motor);

                                try
                                {
                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();

                                    //Abrir programa

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

                                        CrudKslninha.AgregarCrud(desktopSession, 0, crudPrincipal);
                                        //CrudKslninha.BuscarRegistro(desktopSession, 0);


                                        //Aplicar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Guardados", true, file);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        Thread.Sleep(1000);


                                        //Modificar     
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKslninha.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        //Descartar edicion - Cancelar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);

                                        //Confirmar Modificar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKslninha.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Editar Datos", true, file);

                                        //Volver a aplicar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Confirmar Edicion", true, file);

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

                                        //ELIMINAR REGISTRO
                                        //CrudKnmhdeco.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Principal", true, file);


                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //newFunctions_4.ScreenshotNuevo("Eliminar Datos", true, file);
                                        ////Aceptar Eliminar
                                        //WindowsDriver<WindowsElement> btnAcep = null;
                                        //btnAcep = PruebaCRUD.RootSession();
                                        //btnAcep.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                        //newFunctions_4.ScreenshotNuevo("Confirmar Eliminar Datos", true, file);

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
        public void KSlRccomChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSL.ModulosSL_10.KSlRccomChecklist")
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
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null )
                                //rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                //rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                //rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                //rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&

                                ////Variables Validaciones Crud Principal
                                //rows["Grup"].ToString().Length != 0 && rows["Grup"].ToString() != null &&
                                //rows["Req"].ToString().Length != 0 && rows["Req"].ToString() != null &&
                                //rows["Id"].ToString().Length != 0 && rows["Id"].ToString() != null &&


                                ////Variables Validaciones Detalle 1
                                //rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                //rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null)
                            //rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null)


                            {

                                //Variables obligatorias
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();

                                //Variables generales
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                //string banExcel = rows["banExcel"].ToString();
                                //string BanderaLupa = rows["BanderaLupa"].ToString();
                                //string BanderaPreli = rows["BanderaPreli"].ToString();
                                //string BanderaSum = rows["BanderaSum"].ToString();

                                ////Variables validacion  crud Principal
                                //string Grup = rows["Grup"].ToString();
                                //string Req = rows["Req"].ToString();
                                //string Id = rows["Id"].ToString();

                                ////Variables Validaciones Detalle 1
                                //string Tabla = rows["Tabla"].ToString();
                                //string Campo = rows["Campo"].ToString();
                                //string EditCampo = rows["EditCampo"].ToString();

                                string campo = "1";

                                //CRUD
                                List<string> crudPrincipal = new List<string>() { };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "SL", motor);

                                try
                                {
                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();

                                    //Abrir programa

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

                                        //VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////VALIDACION NOMBRE DEL PROGRAMA
                                        //errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        newFunctions_4.ScreenshotNuevo("Consultar Version", true, file);
                                        Thread.Sleep(1000);

                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        newFunctions_4.ScreenshotNuevo("Abrir notas", true, file);
                                        Thread.Sleep(1000);


                                        //Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        //botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        //var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        ////Agregar
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        //desktopSession.Mouse.Click(null);

                                        CrudKslrccom.AgregarCrud(desktopSession, 0, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Insertar datos", true, file);


                                        ////Aplicar
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //newFunctions_4.ScreenshotNuevo("Datos Guardados", true, file);

                                        ////LUPA
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //Thread.Sleep(1000);


                                        ////Modificar     
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        //desktopSession.Mouse.Click(null);

                                        //CrudKslninha.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        //newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        //Thread.Sleep(1000);

                                        ////Descartar edicion - Cancelar
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);

                                        ////Confirmar Modificar
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        //desktopSession.Mouse.Click(null);

                                        //CrudKslninha.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        //newFunctions_4.ScreenshotNuevo("Editar Datos", true, file);

                                        ////Volver a aplicar
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //newFunctions_4.ScreenshotNuevo("Confirmar Edicion", true, file);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_4.ScreenshotNuevo("QBE", true, file);

                                        CrudKslrccom.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Confirmar", true, file);

                                        //  //SUMATORIA
                                        //   errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        //   if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //   newFunctions_1.lapiz(desktopSession);


                                        ////REPORTE DINAMICO
                                        //string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
                                        //errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        ////Rectificar Pie Pagina PDF DINAMICO
                                        //listaResultado = Textopdf(pdf, codProgram, user);
                                        //listaResultado.ForEach(e => ErrorValidacion.Add(e));



                                        ////REPORTE PRELIMINAR
                                        //string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        //errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        ////Rectificar Pie Pagina PDF PRELIMINAR
                                        //listaResultado = Textopdf(pdf1, codProgram, user);
                                        //listaResultado.ForEach(e => ErrorValidacion.Add(e));


                                        ////EXPORTAR EXCEL
                                        //string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        //errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(2000);

                                        ////ELIMINAR REGISTRO
                                        ////CrudKnmhdeco.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        //PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        //newFunctions_4.ScreenshotNuevo("Datos Eliminados Principal", true, file);


                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //newFunctions_4.ScreenshotNuevo("Eliminar Datos", true, file);
                                        ////Aceptar Eliminar
                                        //WindowsDriver<WindowsElement> btnAcep = null;
                                        //btnAcep = PruebaCRUD.RootSession();
                                        //btnAcep.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                        //newFunctions_4.ScreenshotNuevo("Confirmar Eliminar Datos", true, file);

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

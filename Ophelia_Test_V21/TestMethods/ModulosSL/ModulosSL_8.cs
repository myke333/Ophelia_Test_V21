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
    /// <summary>
    /// Descripción resumida de ModulosSL_8
    /// </summary>
    [TestClass]
    public class ModulosSL_8 : FuncionesVitales
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

        public ModulosSL_8()
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
        public void KSlCalifChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSL.ModulosSL_8.KSlCalifChecklist")
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
                                                            //Variables Validaciones Crud principal
                                                            rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                                            rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                                            rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&
                                                            //Variables de los otros campos                        
                                                            rows["CampoTipo"].ToString().Length != 0 && rows["EditCampoTipo"].ToString() != null &&
                                                            rows["CampoNomPrue"].ToString().Length != 0 && rows["CampoNomPrue"].ToString() != null &&
                                                            rows["CampoTope"].ToString().Length != 0 && rows["CampoTope"].ToString() != null &&
                                                            rows["CampoDescrip"].ToString().Length != 0 && rows["CampoDescrip"].ToString() != null &&
                                                            rows["EditCampoTipo"].ToString().Length != 0 && rows["EditCampoTipo"].ToString() != null &&
                                                            rows["EditCampoNomPrue"].ToString().Length != 0 && rows["EditCampoNomPrue"].ToString() != null &&
                                                            rows["EditCampoTope"].ToString().Length != 0 && rows["EditCampoTope"].ToString() != null &&
                                                            rows["EditCampoDescrip"].ToString().Length != 0 && rows["EditCampoDescrip"].ToString() != null &&
                                                            //Datos Crud Principal
                                                            rows["CodPrue"].ToString().Length != 0 && rows["CodPrue"].ToString() != null &&
                                                            rows["Tipo"].ToString().Length != 0 && rows["Tipo"].ToString() != null &&
                                                            rows["NomPrue"].ToString().Length != 0 && rows["NomPrue"].ToString() != null &&
                                                            rows["Tope"].ToString().Length != 0 && rows["Tope"].ToString() != null &&
                                                            rows["Descrip"].ToString().Length != 0 && rows["Tope"].ToString() != null &&
                                                            //Edicion de campos
                                                            rows["ActTipo"].ToString().Length != 0 && rows["ActTipo"].ToString() != null &&
                                                            rows["ActNomPrue"].ToString().Length != 0 && rows["ActNomPrue"].ToString() != null &&
                                                            rows["ActTope"].ToString().Length != 0 && rows["ActTope"].ToString() != null &&
                                                            rows["ActDescrip"].ToString().Length != 0 && rows["ActDescrip"].ToString() != null)
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
                                                                //Variables validacion  crud Principal
                                                                string Tabla = rows["Tabla"].ToString();
                                                                string Campo = rows["Campo"].ToString();
                                                                string EditCampo = rows["EditCampo"].ToString();
                                                                string CampoTipo = rows["CampoTipo"].ToString();
                                                                string EditCampoTipo = rows["EditCampoTipo"].ToString();
                                                                string CampoNomPrue = rows["CampoNomPrue"].ToString();
                                                                string EditCampoNomPrue = rows["EditCampoNomPrue"].ToString();
                                                                string CampoTope = rows["CampoTope"].ToString();
                                                                string EditCampoTope = rows["EditCampoTope"].ToString();
                                                                string CampoDescrip = rows["CampoDescrip"].ToString();
                                                                string EditCampoDesrip = rows["EditCampoDescrip"].ToString();

                                                                //Variables Datos Crud Principal
                                                                string CodPrue = rows["CodPrue"].ToString();
                                                                string Tipo = rows["Tipo"].ToString();
                                                                string NomPrue = rows["NomPrue"].ToString();
                                                                string Tope = rows["Tope"].ToString();
                                                                string Descrip = rows["Descrip"].ToString();
                                                                //Edicion de datos Crud Principal
                                                                string ActTipo = rows["ActTipo"].ToString();
                                                                string ActNomPrue = rows["ActNomPrue"].ToString();
                                                                string ActTope = rows["ActTope"].ToString();
                                                                string ActDescrip = rows["ActDescrip"].ToString();
                                */
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
                                        /*
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
                                        */
                                        /*
                                                                                Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                                                                botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                                                                var ElementList = PruebaCRUD.NavClass(desktopSession);

                                                                                //Agregar
                                                                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                                                                desktopSession.Mouse.Click(null);

                                                                                CrudKSlPrueb.CrudPrueb(desktopSession, "0", crudPrincipal, file);


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

                                                                                CrudKSlPrueb.CrudPrueb(desktopSession, "1", crudPrincipal, file);
                                                                                newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                                                                Thread.Sleep(1000);

                                                                                //Descartar edicion - Cancelar
                                                                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                                                                desktopSession.Mouse.Click(null);
                                                                                newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);

                                                                                //Confirmar Modificar
                                                                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                                                                desktopSession.Mouse.Click(null);

                                                                                CrudKSlPrueb.CrudPrueb(desktopSession, "1", crudPrincipal, file);
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

        [TestMethod]
        public void KSlInhabChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSL.ModulosSL_8.KSlInhabChecklist")
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
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null)
                            //rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null )
                            /*
                                                            //Datos Generales
                                                            rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                                            rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                                            rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                                            rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                                            rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                                            rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null)
                                      */
                            /*
                                                        //Variables Validaciones Crud principal
                                                        rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                                        rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                                        rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&
                                                        //Variables de los otros campos                        
                                                        rows["CampoTipo"].ToString().Length != 0 && rows["EditCampoTipo"].ToString() != null &&
                                                        rows["CampoNomPrue"].ToString().Length != 0 && rows["CampoNomPrue"].ToString() != null &&
                                                        rows["CampoTope"].ToString().Length != 0 && rows["CampoTope"].ToString() != null &&
                                                        rows["CampoDescrip"].ToString().Length != 0 && rows["CampoDescrip"].ToString() != null &&
                                                        rows["EditCampoTipo"].ToString().Length != 0 && rows["EditCampoTipo"].ToString() != null &&
                                                        rows["EditCampoNomPrue"].ToString().Length != 0 && rows["EditCampoNomPrue"].ToString() != null &&
                                                        rows["EditCampoTope"].ToString().Length != 0 && rows["EditCampoTope"].ToString() != null &&
                                                        rows["EditCampoDescrip"].ToString().Length != 0 && rows["EditCampoDescrip"].ToString() != null &&
                                                        //Datos Crud Principal
                                                        rows["CodPrue"].ToString().Length != 0 && rows["CodPrue"].ToString() != null &&
                                                        rows["Tipo"].ToString().Length != 0 && rows["Tipo"].ToString() != null &&
                                                        rows["NomPrue"].ToString().Length != 0 && rows["NomPrue"].ToString() != null &&
                                                        rows["Tope"].ToString().Length != 0 && rows["Tope"].ToString() != null &&
                                                        rows["Descrip"].ToString().Length != 0 && rows["Tope"].ToString() != null &&
                                                        //Edicion de campos
                                                        rows["ActTipo"].ToString().Length != 0 && rows["ActTipo"].ToString() != null &&
                                                        rows["ActNomPrue"].ToString().Length != 0 && rows["ActNomPrue"].ToString() != null &&
                                                        rows["ActTope"].ToString().Length != 0 && rows["ActTope"].ToString() != null &&
                                                        rows["ActDescrip"].ToString().Length != 0 && rows["ActDescrip"].ToString() != null)
                        */
                            {

                                //Variables obligatorias
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                //string nomProgram = rows["NomProgram"].ToString();
                                /*
                                                                //Variables generales
                                                                string TipoQbe = rows["TipoQbe"].ToString();
                                                                string QbeFilter = rows["QbeFilter"].ToString();
                                                                string banExcel = rows["banExcel"].ToString();
                                                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                                                string BanderaPreli = rows["BanderaPreli"].ToString();
                                                                string BanderaSum = rows["BanderaSum"].ToString();
                                                                /*
                                                                                                //Variables validacion  crud Principal
                                                                                                string Tabla = rows["Tabla"].ToString();
                                                                                                string Campo = rows["Campo"].ToString();
                                                                                                string EditCampo = rows["EditCampo"].ToString();
                                                                                                string CampoTipo = rows["CampoTipo"].ToString();
                                                                                                string EditCampoTipo = rows["EditCampoTipo"].ToString();
                                                                                                string CampoNomPrue = rows["CampoNomPrue"].ToString();
                                                                                                string EditCampoNomPrue = rows["EditCampoNomPrue"].ToString();
                                                                                                string CampoTope = rows["CampoTope"].ToString();
                                                                                                string EditCampoTope = rows["EditCampoTope"].ToString();
                                                                                                string CampoDescrip = rows["CampoDescrip"].ToString();
                                                                                                string EditCampoDesrip = rows["EditCampoDescrip"].ToString();

                                                                                                //Variables Datos Crud Principal
                                                                                                string CodPrue = rows["CodPrue"].ToString();
                                                                                                string Tipo = rows["Tipo"].ToString();
                                                                                                string NomPrue = rows["NomPrue"].ToString();
                                                                                                string Tope = rows["Tope"].ToString();
                                                                                                string Descrip = rows["Descrip"].ToString();
                                                                                                //Edicion de datos Crud Principal
                                                                                                string ActTipo = rows["ActTipo"].ToString();
                                                                                                string ActNomPrue = rows["ActNomPrue"].ToString();
                                                                                                string ActTope = rows["ActTope"].ToString();
                                                                                                string ActDescrip = rows["ActDescrip"].ToString();
                                                                */
                                // string campo = "0";


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

                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //Agregar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKSlPrueb.CrudPrueb(desktopSession, "0", crudPrincipal, file);

                                        /*
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

                                                                                                                        CrudKSlPrueb.CrudPrueb(desktopSession, "1", crudPrincipal, file);
                                                                                                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                                                                                                        Thread.Sleep(1000);

                                                                                                                        //Descartar edicion - Cancelar
                                                                                                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                                                                                                        desktopSession.Mouse.Click(null);
                                                                                                                        newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);

                                                                                                                        //Confirmar Modificar
                                                                                                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                                                                                                        desktopSession.Mouse.Click(null);

                                                                                                                        CrudKSlPrueb.CrudPrueb(desktopSession, "1", crudPrincipal, file);
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
                                                                                */
                                        /*
                                                                                //QBE
                                                                                errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
                                                                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                                                                /*                                
                                                                                  //SUMATORIA
                                                                                   errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                                                                   if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                                                                   newFunctions_1.lapiz(desktopSession);
                                                                                */
                                        /*
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
        public void AbrirPrograma()
        {
            string programa = "KSlRecla";
            //AbrirPrograma a = new AbrirPrograma(programa, "UNatalia");
            //desktopSession = a.Start2("SQL", "");
             AbrirPrograma a = new AbrirPrograma(programa, "ONATALIA");
             desktopSession = a.Start2("ORA", "");
        }

        [TestMethod]
        public void KSlReclaChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSL.ModulosSL_8.KSlReclaChecklist")
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
                                                            //Variables Validaciones Crud principal
                                                            rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                                            rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                                            rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&
                                                            //Variables de los otros campos                        
                                                            rows["CampoTipo"].ToString().Length != 0 && rows["EditCampoTipo"].ToString() != null &&
                                                            rows["CampoNomPrue"].ToString().Length != 0 && rows["CampoNomPrue"].ToString() != null &&
                                                            rows["CampoTope"].ToString().Length != 0 && rows["CampoTope"].ToString() != null &&
                                                            rows["CampoDescrip"].ToString().Length != 0 && rows["CampoDescrip"].ToString() != null &&
                                                            rows["EditCampoTipo"].ToString().Length != 0 && rows["EditCampoTipo"].ToString() != null &&
                                                            rows["EditCampoNomPrue"].ToString().Length != 0 && rows["EditCampoNomPrue"].ToString() != null &&
                                                            rows["EditCampoTope"].ToString().Length != 0 && rows["EditCampoTope"].ToString() != null &&
                                                            rows["EditCampoDescrip"].ToString().Length != 0 && rows["EditCampoDescrip"].ToString() != null &&
                                                            //Datos Crud Principal
                                                            rows["CodPrue"].ToString().Length != 0 && rows["CodPrue"].ToString() != null &&
                                                            rows["Tipo"].ToString().Length != 0 && rows["Tipo"].ToString() != null &&
                                                            rows["NomPrue"].ToString().Length != 0 && rows["NomPrue"].ToString() != null &&
                                                            rows["Tope"].ToString().Length != 0 && rows["Tope"].ToString() != null &&
                                                            rows["Descrip"].ToString().Length != 0 && rows["Tope"].ToString() != null &&
                                                            //Edicion de campos
                                                            rows["ActTipo"].ToString().Length != 0 && rows["ActTipo"].ToString() != null &&
                                                            rows["ActNomPrue"].ToString().Length != 0 && rows["ActNomPrue"].ToString() != null &&
                                                            rows["ActTope"].ToString().Length != 0 && rows["ActTope"].ToString() != null &&
                                                            rows["ActDescrip"].ToString().Length != 0 && rows["ActDescrip"].ToString() != null)
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
                                                                //Variables validacion  crud Principal
                                                                string Tabla = rows["Tabla"].ToString();
                                                                string Campo = rows["Campo"].ToString();
                                                                string EditCampo = rows["EditCampo"].ToString();
                                                                string CampoTipo = rows["CampoTipo"].ToString();
                                                                string EditCampoTipo = rows["EditCampoTipo"].ToString();
                                                                string CampoNomPrue = rows["CampoNomPrue"].ToString();
                                                                string EditCampoNomPrue = rows["EditCampoNomPrue"].ToString();
                                                                string CampoTope = rows["CampoTope"].ToString();
                                                                string EditCampoTope = rows["EditCampoTope"].ToString();
                                                                string CampoDescrip = rows["CampoDescrip"].ToString();
                                                                string EditCampoDesrip = rows["EditCampoDescrip"].ToString();

                                                                //Variables Datos Crud Principal
                                                                string CodPrue = rows["CodPrue"].ToString();
                                                                string Tipo = rows["Tipo"].ToString();
                                                                string NomPrue = rows["NomPrue"].ToString();
                                                                string Tope = rows["Tope"].ToString();
                                                                string Descrip = rows["Descrip"].ToString();
                                                                //Edicion de datos Crud Principal
                                                                string ActTipo = rows["ActTipo"].ToString();
                                                                string ActNomPrue = rows["ActNomPrue"].ToString();
                                                                string ActTope = rows["ActTope"].ToString();
                                                                string ActDescrip = rows["ActDescrip"].ToString();
                                                             */  
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
                                        
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //Agregar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKSlRecla.AgregarCrud(desktopSession, 0);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

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

                                        CrudKSlRecla.AgregarCrud(desktopSession, 1);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        //Descartar edicion - Cancelar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);

                                        //Confirmar Modificar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKSlRecla.AgregarCrud(desktopSession, 1);
                                        newFunctions_4.ScreenshotNuevo("Editar Datos", true, file);

                                        //Volver a aplicar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Confirmar Edicion", true, file);

                                        
                                        
                                        
                                       //QBE
                                       errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
                                       if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_4.ScreenshotNuevo("Filtro QBE", true, file);
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
                                        newFunctions_4.ScreenshotNuevo("Reporte Dinamico", true, file);


                                        //REPORTE PRELIMINAR
                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                       errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                                       if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                       //Rectificar Pie Pagina PDF PRELIMINAR
                                       listaResultado = Textopdf(pdf1, codProgram, user);
                                       listaResultado.ForEach(e => ErrorValidacion.Add(e));
                                        newFunctions_4.ScreenshotNuevo("Reporte Preliminar", true, file);

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                       errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                       if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                       Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Exportar Excel", true, file);


                                        //Eliminar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Datos", true, file);

                                        //Aceptar Eliminar
                                        WindowsDriver<WindowsElement> btnAcep = null;
                                        btnAcep = PruebaCRUD.RootSession();
                                        btnAcep.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                        newFunctions_4.ScreenshotNuevo("Confirmar Eliminar Datos", true, file);


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
        public void KSlElegiChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSL.ModulosSL_8.KSlElegiChecklist")
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
                                //Variables Validaciones Crud principal
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&
                                //Variables de los otros campos                        
                                rows["CampoTipo"].ToString().Length != 0 && rows["EditCampoTipo"].ToString() != null &&
                                rows["CampoNomPrue"].ToString().Length != 0 && rows["CampoNomPrue"].ToString() != null &&
                                rows["CampoTope"].ToString().Length != 0 && rows["CampoTope"].ToString() != null &&
                                rows["CampoDescrip"].ToString().Length != 0 && rows["CampoDescrip"].ToString() != null &&
                                rows["EditCampoTipo"].ToString().Length != 0 && rows["EditCampoTipo"].ToString() != null &&
                                rows["EditCampoNomPrue"].ToString().Length != 0 && rows["EditCampoNomPrue"].ToString() != null &&
                                rows["EditCampoTope"].ToString().Length != 0 && rows["EditCampoTope"].ToString() != null &&
                                rows["EditCampoDescrip"].ToString().Length != 0 && rows["EditCampoDescrip"].ToString() != null &&
                                //Datos Crud Principal
                                rows["CodPrue"].ToString().Length != 0 && rows["CodPrue"].ToString() != null &&
                                rows["Tipo"].ToString().Length != 0 && rows["Tipo"].ToString() != null &&
                                rows["NomPrue"].ToString().Length != 0 && rows["NomPrue"].ToString() != null &&
                                rows["Tope"].ToString().Length != 0 && rows["Tope"].ToString() != null &&
                                rows["Descrip"].ToString().Length != 0 && rows["Tope"].ToString() != null &&
                                //Edicion de campos
                                rows["ActTipo"].ToString().Length != 0 && rows["ActTipo"].ToString() != null &&
                                rows["ActNomPrue"].ToString().Length != 0 && rows["ActNomPrue"].ToString() != null &&
                                rows["ActTope"].ToString().Length != 0 && rows["ActTope"].ToString() != null &&
                                rows["ActDescrip"].ToString().Length != 0 && rows["ActDescrip"].ToString() != null)
                            */
                            {
                                //Seleccionar los datos con la lupa.
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
                                //Variables validacion  crud Principal
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string EditCampo = rows["EditCampo"].ToString();
                                string CampoTipo = rows["CampoTipo"].ToString();
                                string EditCampoTipo = rows["EditCampoTipo"].ToString();
                                string CampoNomPrue = rows["CampoNomPrue"].ToString();
                                string EditCampoNomPrue = rows["EditCampoNomPrue"].ToString();
                                string CampoTope = rows["CampoTope"].ToString();
                                string EditCampoTope = rows["EditCampoTope"].ToString();
                                string CampoDescrip = rows["CampoDescrip"].ToString();
                                string EditCampoDesrip = rows["EditCampoDescrip"].ToString();
                               
                                //Variables Datos Crud Principal
                                string CodPrue = rows["CodPrue"].ToString();
                                string Tipo = rows["Tipo"].ToString();
                                string NomPrue = rows["NomPrue"].ToString();
                                string Tope = rows["Tope"].ToString();
                                string Descrip = rows["Descrip"].ToString();
                                //Edicion de datos Crud Principal
                                string ActTipo = rows["ActTipo"].ToString();
                                string ActNomPrue = rows["ActNomPrue"].ToString();
                                string ActTope = rows["ActTope"].ToString();
                                string ActDescrip = rows["ActDescrip"].ToString();
                                                                */
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

                                        //Agregar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKSlPrueb.CrudPrueb(desktopSession, "0", crudPrincipal, file);


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

                                        CrudKSlPrueb.CrudPrueb(desktopSession, "1", crudPrincipal, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        //Descartar edicion - Cancelar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);

                                        //Confirmar Modificar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKSlPrueb.CrudPrueb(desktopSession, "1", crudPrincipal, file);
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

        [TestMethod]
        public void KSlCrideChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSL.ModulosSL_8.KSlCrideChecklist")
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
                                rows["CampoNom"].ToString().Length != 0 && rows["CampoNom"].ToString() != null &&
                                rows["CampoPrio"].ToString().Length != 0 && rows["CampoPrio"].ToString() != null &&
                                // Editar los campos
                                rows["EditCampoNom"].ToString().Length != 0 && rows["EditCampoNom"].ToString() != null &&
                                rows["EditCampoPrio"].ToString().Length != 0 && rows["EditCampoPrio"].ToString() != null &&
                                //Datos Crud Principal
                                rows["Cod"].ToString().Length != 0 && rows["Cod"].ToString() != null &&
                                rows["Nom"].ToString().Length != 0 && rows["Nom"].ToString() != null &&
                                rows["Prio"].ToString().Length != 0 && rows["Prio"].ToString() != null &&
                                //Edicion de campos
                                rows["ActNom"].ToString().Length != 0 && rows["ActNom"].ToString() != null &&
                                rows["ActPrio"].ToString().Length != 0 && rows["ActPrio"].ToString() != null)

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
                                string CampoNom = rows["CampoNom"].ToString();
                                string CampoPrio = rows["CampoPrio"].ToString();
                                //Variable Editar Campos
                                string EditCampoNom = rows["EditCampoNom"].ToString();
                                string EditCampoPrio = rows["EditCampoPrio"].ToString();

                                //Variables Ingresar Crud Principal
                                string Cod = rows["Cod"].ToString();
                                string Nom = rows["Nom"].ToString();
                                string Prio = rows["Prio"].ToString();

                                //Variables Editar Crud
                                string ActNom = rows["ActNom"].ToString();
                                string ActPrio = rows["ActPrio"].ToString();

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
                                    List<string> crudPrincipal = new List<string>() { Cod, Nom, Prio, ActNom, ActPrio };

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
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //Agregar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKSlCride.CrudCride(desktopSession, "0", crudPrincipal, file);


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

                                        CrudKSlCride.CrudCride(desktopSession, "1", crudPrincipal, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        //Descartar edicion - Cancelar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);

                                        //Confirmar Modificar
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKSlCride.CrudCride(desktopSession, "1", crudPrincipal, file);
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
        public void KSlCrempChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSL.ModulosSL_8.KSlCrempChecklist")
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
                                // rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&

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
                                                            rows["CampoNom"].ToString().Length != 0 && rows["CampoNom"].ToString() != null &&
                                                            rows["CampoPrio"].ToString().Length != 0 && rows["CampoPrio"].ToString() != null &&
                                                            // Editar los campos
                                                            rows["EditCampoNom"].ToString().Length != 0 && rows["EditCampoNom"].ToString() != null &&
                                                            rows["EditCampoPrio"].ToString().Length != 0 && rows["EditCampoPrio"].ToString() != null &&
                                                            //Datos Crud Principal
                                                            rows["Cod"].ToString().Length != 0 && rows["Cod"].ToString() != null &&
                                                            rows["Nom"].ToString().Length != 0 && rows["Nom"].ToString() != null &&
                                                            rows["Prio"].ToString().Length != 0 && rows["Prio"].ToString() != null &&
                                                            //Edicion de campos
                                                            rows["ActNom"].ToString().Length != 0 && rows["ActNom"].ToString() != null &&
                                                            rows["ActPrio"].ToString().Length != 0 && rows["ActPrio"].ToString() != null)
                            */
                            {

                                //Variables obligatorias
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                //   string nomProgram = rows["NomProgram"].ToString();

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
                                                                string CampoNom = rows["CampoNom"].ToString();
                                                                string CampoPrio = rows["CampoPrio"].ToString();
                                                                //Variable Editar Campos
                                                                string EditCampoNom = rows["EditCampoNom"].ToString();
                                                                string EditCampoPrio = rows["EditCampoPrio"].ToString();

                                                                //Variables Ingresar Crud Principal
                                                                string Cod = rows["Cod"].ToString();
                                                                string Nom = rows["Nom"].ToString();
                                                                string Prio = rows["Prio"].ToString();

                                                                //Variables Editar Crud
                                                                string ActNom = rows["ActNom"].ToString();
                                                                string ActPrio = rows["ActPrio"].ToString();
                                */
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

                                                                                //Agregar
                                                                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                                                                desktopSession.Mouse.Click(null);

                                                                                CrudKSlCride.CrudCride(desktopSession, "0", crudPrincipal, file);


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

                                                                                CrudKSlCride.CrudCride(desktopSession, "1", crudPrincipal, file);
                                                                                newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                                                                Thread.Sleep(1000);

                                                                                //Descartar edicion - Cancelar
                                                                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                                                                desktopSession.Mouse.Click(null);
                                                                                newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);

                                                                                //Confirmar Modificar
                                                                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                                                                desktopSession.Mouse.Click(null);

                                                                                CrudKSlCride.CrudCride(desktopSession, "1", crudPrincipal, file);
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

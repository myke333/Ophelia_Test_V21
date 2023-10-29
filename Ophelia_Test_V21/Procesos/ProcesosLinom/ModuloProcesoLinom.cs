using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using OpenQA.Selenium.Appium;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesosIncap;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesosKNmAusen;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesosKNmEmbar;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesosKNmNoved;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesosLinom;
using System.Diagnostics;

namespace Ophelia_Test_V21.Procesos.ProcesosLinom
{
    /// <summary>
    /// Descripción resumida de ModuloProcesoLinom
    /// </summary>
    [TestClass]
    public class ModuloProcesoLinom : FuncionesVitales
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
        public ModuloProcesoLinom()
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

        [TestMethod]
        public void RegistroKNmIncap()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesosLinom.ModuloProcesoLinom.RegistroKNmIncap")
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

                                //Variables Validaciones Crud principal
                                rows["Id"].ToString().Length != 0 && rows["Id"].ToString() != null &&
                                rows["Concepto"].ToString().Length != 0 && rows["Concepto"].ToString() != null &&
                                rows["Numero"].ToString().Length != 0 && rows["Numero"].ToString() != null &&
                                rows["Codigo"].ToString().Length != 0 && rows["Codigo"].ToString() != null &&
                                rows["NumeroDias"].ToString().Length != 0 && rows["NumeroDias"].ToString() != null
                                )

                            {
                                //Variables obligatorias
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();

                                //Variables generales
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();

                                //Variables Crud principal
                                string Id = rows["Id"].ToString();
                                string Concepto = rows["Concepto"].ToString();
                                string Numero = rows["Numero"].ToString();
                                string Codigo = rows["Codigo"].ToString();
                                string NumeroDias = rows["NumeroDias"].ToString();
                                string dia = rows["dia"].ToString();
                                string mes = rows["mes"].ToString();
                                //String Codigo = "9";

                                String campo = "2";

                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { Id, Concepto, Numero, Codigo, NumeroDias, dia, mes };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "RegistroIncapacidad", motor);

                                string delete1 = $"delete from nm_incap where cod_empl= '{Id}' and cod_conc ='{Concepto}'";
                                FuncionesVitales.UpdateDeleteInsert(delete1, motor, user);
                                string delete2 = $"delete from nm_noved where cod_empl= '{Id}' and cod_conc ='{Concepto}'";
                                FuncionesVitales.UpdateDeleteInsert(delete2, motor, user);
                                string delete3 = $"delete from nm_ausen where cod_empl= '{Id}' and cod_conc ='{Concepto}'";
                                FuncionesVitales.UpdateDeleteInsert(delete3, motor, user); 
                                string delete4 = $"delete from nm_diasn where cod_empl= '{Id}' and cod_conc ='{Concepto}'";
                                FuncionesVitales.UpdateDeleteInsert(delete4, motor, user);
                                string delete5 = $"delete from nm_noved where cod_empl= '{Id}' and cod_conc ='1100'";
                                FuncionesVitales.UpdateDeleteInsert(delete5, motor, user);
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
                                        Thread.Sleep(1500);
                                        Thread.Sleep(1500);
                                        //AGREGAR REGISTRO Principal
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        FuncionesProcesosIncap.AgregarRegistro(desktopSession, 0, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);
                                        Thread.Sleep(1500);
                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        
                                        DateTime fechaActual = DateTime.Today;
                                        string año = fechaActual.ToString("yyyy");
                                        string fechainicio = año + mes + dia;
                                        //Validar correcto registro
                                        string registro = string.Empty;
                                        string sentencia = "select rmt_inca from nm_incap where cod_empl=" + "'" + Id + "'" + "and cod_conc=" + "'" + Concepto + "'" + " and fec_desd=" + "'" + fechainicio + "'";
                                        DataTable resultado = FuncionesVitales.SelectData(sentencia, user, motor);
                                        if (resultado.Rows.Count > 0)
                                        {
                                            foreach (DataRow rw in resultado.Rows)
                                            {
                                                registro = rw["rmt_inca"].ToString();

                                            }
                                        }
                                        else
                                        {
                                            Assert.Fail("Rgistro incapacidad para " + Id + " no fue correcta");
                                        }


                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }

                                    if (ErrorValidacion.Count > 0)
                                    {
                                        FuncionesVitales.ReporteErrores(ErrorValidacion, codProgram, motor, Modulo);
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, ErrorValidacion);
                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}", Environment.NewLine, errorMessageString));
                                    }

                                    Thread.Sleep(4000);
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
                                    //break;

                                }
                                catch (Exception e)
                                {
                                    Thread.Sleep(4000);

                                    //Add Juan
                                    newFunctions_4.ScreenshotNuevo("Error en la prueba", true, file);

                                    //Add Daniel
                                    TestContext.AddResultFile(newFunctions_4.ScreenshotInicio(codProgram));
                                    //
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
                                    Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    break;
                                }
                            }
                        }
                        //break;
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
        public void RegistroKNmAusen()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesosLinom.ModuloProcesoLinom.RegistroKNmAusen")
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

                                //Variables Validaciones Crud principal
                                rows["Id"].ToString().Length != 0 && rows["Id"].ToString() != null &&
                                rows["Concepto"].ToString().Length != 0 && rows["Concepto"].ToString() != null &&
                                rows["Ausencia"].ToString().Length != 0 && rows["Ausencia"].ToString() != null

                                )


                            {
                                //Variables obligatorias
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();

                                //Variables generales
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();

                                //Variables Crud principal
                                string Id = rows["Id"].ToString();
                                string Concepto = rows["Concepto"].ToString();
                                string Ausencia = rows["Ausencia"].ToString();
                                string diaInicio = rows["diaInicio"].ToString();
                                string diaFinal = rows["diaFinal"].ToString();
                                string mes = rows["mes"].ToString();
                                //String Codigo = "9";

                                string codigo = "1";

                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { Id, Concepto, Ausencia, diaInicio, diaFinal, mes };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "RegistroAusentismo", motor);
                                
                                
                                string delete2 = $"delete from nm_noved where cod_empl= '{Id}' and cod_conc ='{Concepto}'";
                                FuncionesVitales.UpdateDeleteInsert(delete2, motor, user);
                                string delete3 = $"delete from nm_ausen where cod_empl= '{Id}' and cod_conc ='{Concepto}'";
                                FuncionesVitales.UpdateDeleteInsert(delete3, motor, user);
                                string delete4 = $"delete from nm_diasn where cod_empl= '{Id}' and cod_conc ='{Concepto}'";
                                FuncionesVitales.UpdateDeleteInsert(delete4, motor, user);
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

                                        Thread.Sleep(3000);
                                        //AGREGAR REGISTRO Principal
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        FuncionesProcesosKNmAusen.AgregarRegistro(desktopSession, 0, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(3000);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, codigo);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        DateTime fechaActual = DateTime.Today;
                                        string año = fechaActual.ToString("yyyy");
                                        string fechainicio = año + mes + diaInicio;
                                        //Validar correcto registro
                                        string registroausen = string.Empty;
                                        string sentencia = "select rmt_ause from nm_ausen where cod_empl="+"'"+Id+"'"+"and cod_conc=" + "'" + Concepto + "'"+" and fec_desd=" + "'"+fechainicio+"'";
                                        DataTable resultado = FuncionesVitales.SelectData(sentencia, user, motor);
                                        if (resultado.Rows.Count > 0)
                                        {
                                            foreach (DataRow rw in resultado.Rows)
                                            {
                                                registroausen = rw["rmt_ause"].ToString();

                                            }
                                        }
                                        else
                                        {
                                            Assert.Fail("Rgistro ausentismos para " + Id + " no fue correcta");
                                        }


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
                                    //break;

                                }
                                catch (Exception e)
                                {
                                    Thread.Sleep(4000);

                                    //Add Juan
                                    newFunctions_4.ScreenshotNuevo("Error en la prueba", true, file);

                                    //Add Daniel
                                    TestContext.AddResultFile(newFunctions_4.ScreenshotInicio(codProgram));
                                    //
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
                                    Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    break;
                                }
                            }
                        }
                        //break;
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
        public void RegistroKNmEmbar()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesosLinom.ModuloProcesoLinom.RegistroKNmEmbar")
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

                                //Variables Validaciones Crud principal
                                rows["Id"].ToString().Length != 0 && rows["Id"].ToString() != null &&
                                rows["FechaIncio"].ToString().Length != 0 && rows["FechaIncio"].ToString() != null &&
                                rows["Concepto"].ToString().Length != 0 && rows["Concepto"].ToString() != null &&
                                rows["Porcentaje"].ToString().Length != 0 && rows["Porcentaje"].ToString() != null &&
                                rows["ValorFijo"].ToString().Length != 0 && rows["ValorFijo"].ToString() != null &&
                                rows["Cedula"].ToString().Length != 0 && rows["Cedula"].ToString() != null &&
                                rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
                                rows["Apellido"].ToString().Length != 0 && rows["Apellido"].ToString() != null &&
                                rows["Codigo"].ToString().Length != 0 && rows["Codigo"].ToString() != null &&
                                rows["Entidad"].ToString().Length != 0 && rows["Entidad"].ToString() != null &&
                                rows["TipoCuenta"].ToString().Length != 0 && rows["TipoCuenta"].ToString() != null &&
                                rows["Municipio"].ToString().Length != 0 && rows["Municipio"].ToString() != null &&
                                rows["Oficio"].ToString().Length != 0 && rows["Oficio"].ToString() != null &&
                                rows["FechaOficio"].ToString().Length != 0 && rows["FechaOficio"].ToString() != null &&
                                rows["Beneficiario"].ToString().Length != 0 && rows["Beneficiario"].ToString() != null &&
                                rows["Cuentas"].ToString().Length != 0 && rows["Cuentas"].ToString() != null &&
                                rows["NumeroExpediente"].ToString().Length != 0 && rows["NumeroExpediente"].ToString() != null &&
                                rows["NumeroProceso"].ToString().Length != 0 && rows["NumeroProceso"].ToString() != null)

                            {
                                //Variables obligatorias
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();

                                //Variables generales
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();

                                //Variables Crud principal
                                string Id = rows["Id"].ToString();
                                string FechaIncio = rows["FechaIncio"].ToString();
                                string Concepto = rows["Concepto"].ToString();
                                string Porcentaje = rows["Porcentaje"].ToString();
                                string ValorFijo = rows["ValorFijo"].ToString();
                                string Cedula = rows["Cedula"].ToString();
                                string Nombre = rows["Nombre"].ToString();
                                string Apellido = rows["Apellido"].ToString();
                                string Codigo = rows["Codigo"].ToString();
                                string Entidad = rows["Entidad"].ToString();
                                string Municipio = rows["Municipio"].ToString();
                                string Oficio = rows["Oficio"].ToString();
                                string FechaOficio = rows["FechaOficio"].ToString();
                                string Beneficiario = rows["Beneficiario"].ToString();
                                string Cuentas = rows["Cuentas"].ToString();
                                string NumeroExpediente = rows["NumeroExpediente"].ToString();
                                string NumeroProceso = rows["NumeroProceso"].ToString();
                                string banderacrud = rows["bandera"].ToString();



                                //String Codigo = "9";

                                String campo = "1";

                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { Id, FechaIncio, Concepto, Porcentaje, ValorFijo, Cedula, Nombre, Apellido, Codigo, Entidad, Municipio, Oficio, FechaOficio, Beneficiario, Cuentas, NumeroExpediente, NumeroProceso };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = FuncionesVitales.CrearDocumentoWordDinamico(codProgram, "RegistroEmbargos", motor);

                                string delete1 = $"delete from nm_embar where cod_empl= '{Id}' and cod_conc ='{Concepto}'";
                                FuncionesVitales.UpdateDeleteInsert(delete1, motor, user);
                                string delete2 = $"delete from nm_noved where cod_empl= '{Id}' and cod_conc ='{Concepto}'";
                                FuncionesVitales.UpdateDeleteInsert(delete2, motor, user);
                               
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

                                        Thread.Sleep(4000);
                                        //AGREGAR REGISTRO Principal
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        FuncionesProcesosNmEmbar.AgregarRegistro(desktopSession, banderacrud, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);
                                        Thread.Sleep(1500);

                                        WindowsDriver<WindowsElement> RootSession2 = null;
                                        RootSession2 = PruebaCRUD.RootSession();
                                        RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                        Thread.Sleep(2000);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        var dateTime = DateTime.Parse(FechaIncio);
                                        string fecha = dateTime.ToString("yyyyMMdd");
                                     
                                        //Validar correcto registro
                                        string registro = string.Empty;
                                        string sentencia = "select rmt_emba from nm_embar where cod_empl=" + "'" + Id + "'" + "and cod_conc=" + "'" + Concepto + "'" + " and fec_desd=" + "'" + fecha + "'";
                                        DataTable resultado = FuncionesVitales.SelectData(sentencia, user, motor);
                                        if (resultado.Rows.Count > 0)
                                        {
                                            foreach (DataRow rw in resultado.Rows)
                                            {
                                                registro = rw["rmt_emba"].ToString();

                                            }
                                        }
                                        else
                                        {
                                            Assert.Fail("Registro embargos para " + Id + " no fue correcta");
                                        }


                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }

                                    if (ErrorValidacion.Count > 0)
                                    {
                                        FuncionesVitales.ReporteErrores(ErrorValidacion, codProgram, motor, "NM");
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, ErrorValidacion);
                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}", Environment.NewLine, errorMessageString));
                                    }

                                    Thread.Sleep(4000);
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
                                    //break;

                                }
                                catch (Exception e)
                                {
                                    Thread.Sleep(4000);

                                    //Add Juan
                                    newFunctions_4.ScreenshotNuevo("Error en la prueba", true, file);

                                    //Add Daniel
                                    TestContext.AddResultFile(newFunctions_4.ScreenshotInicio(codProgram));
                                    //
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
                                    Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    break;
                                }
                            }
                        }
                        //break;
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
        public void RegistroKNmNoved()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesosLinom.ModuloProcesoLinom.RegistroKNmNoved")
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

                                //Variables Validaciones Crud principal
                                rows["Id"].ToString().Length != 0 && rows["Id"].ToString() != null &&
                                rows["Concepto"].ToString().Length != 0 && rows["Concepto"].ToString() != null &&
                                rows["Cantidad"].ToString().Length != 0 && rows["Cantidad"].ToString() != null &&
                                rows["ValorCuota"].ToString().Length != 0 && rows["ValorCuota"].ToString() != null &&
                                rows["NumeroCuotas"].ToString().Length != 0 && rows["NumeroCuotas"].ToString() != null &&
                                rows["TipoNovedad"].ToString().Length != 0 && rows["TipoNovedad"].ToString() != null
                               )

                            {
                                //Variables obligatorias
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();

                                //Variables generales
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();

                                //Variables Crud principal
                                string Id = rows["Id"].ToString();
                                string Concepto = rows["Concepto"].ToString();
                                string Cantidad = rows["Cantidad"].ToString();
                                string ValorCuota = rows["ValorCuota"].ToString();
                                string NumeroCuotas = rows["NumeroCuotas"].ToString();
                                string dia = rows["dia"].ToString();
                                string mes = rows["mes"].ToString();
                                string TipoNovedad = rows["TipoNovedad"].ToString();


                                //String Codigo = "9";

                                //String campo = "1";

                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { Id, Concepto, TipoNovedad, Cantidad, ValorCuota, NumeroCuotas, dia, mes };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = FuncionesVitales.CrearDocumentoWordDinamico(codProgram, "RegistroNovedadesNomina", motor);

                                string delete2 = $"delete from nm_noved where cod_empl= '{Id}' and cod_conc ='{Concepto}'";
                                FuncionesVitales.UpdateDeleteInsert(delete2, motor, user);
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
                                        Thread.Sleep(1500);
                                        Thread.Sleep(1500);

                                        //AGREGAR REGISTRO Principal
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        FuncionesProcesosKNmNoved.AgregarRegistro(desktopSession, TipoNovedad, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);
                                        Thread.Sleep(1500);

                                        WindowsDriver<WindowsElement> RootSession2 = null;
                                        RootSession2 = PruebaCRUD.RootSession();
                                        RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                        Thread.Sleep(2000);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        
                                        DateTime fechaActual = DateTime.Today;
                                        string año = fechaActual.ToString("yyyy");
                                        string fechainicio = año + mes + dia;
                                        //Validar correcto registro
                                        string registro = string.Empty;
                                        string sentencia = "select rmt_nove from nm_noved where cod_empl=" + "'" + Id + "'" + "and cod_conc=" + "'" + Concepto + "'" + " and fec_pago=" + "'" + fechainicio + "'";
                                        DataTable resultado = FuncionesVitales.SelectData(sentencia, user, motor);
                                        if (resultado.Rows.Count > 0)
                                        {
                                            foreach (DataRow rw in resultado.Rows)
                                            {
                                                registro = rw["rmt_nove"].ToString();

                                            }
                                        }
                                        else
                                        {
                                            Assert.Fail("Rgistro novedades para " + Id + " no fue correcta");
                                        }


                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }

                                    if (ErrorValidacion.Count > 0)
                                    {
                                        FuncionesVitales.ReporteErrores(ErrorValidacion, codProgram, motor, "NM");
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, ErrorValidacion);
                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba son:{0}{1}", Environment.NewLine, errorMessageString));
                                    }

                                    Thread.Sleep(4000);
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
                                    //break;

                                }
                                catch (Exception e)
                                {
                                    Thread.Sleep(4000);

                                    //Add Juan
                                    newFunctions_4.ScreenshotNuevo("Error en la prueba", true, file);

                                    //Add Daniel
                                    TestContext.AddResultFile(newFunctions_4.ScreenshotInicio(codProgram));
                                    //
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
                                    Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    break;
                                }
                            }
                        }
                        //break;
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
        public void CorrerProceso()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesosLinom.ModuloProcesoLinom.CorrerProceso")
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

                                //Variables Validaciones Crud principal
                                rows["Id"].ToString().Length != 0 && rows["Id"].ToString() != null)

                            {
                                //Variables obligatorias
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();

                                //Variables generales
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();

                                //Variables Crud principal
                                string Id = rows["Id"].ToString();
                                string dia = rows["dia"].ToString();
                                string mes = rows["mes"].ToString();

                                //String Codigo = "9";



                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { Id, mes , dia };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "LiquidacionNomina", motor);

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
                                        Thread.Sleep(3000);

                                        //Proyectar Fechas de liquidacion
                                        FuncionesProcesosLinom.CorrerNomina(desktopSession, 0, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Datos Proceso", true, file);
                                        Thread.Sleep(1500);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //Iniciar Proceso
                                        newFunctions_4.ScreenshotNuevo("Empleado a Liquidar", true, file);
                                        FuncionesProcesosLinom.CorrerNomina(desktopSession, 1, crudPrincipal);
                                        Thread.Sleep(10000);
                                        newFunctions_4.ScreenshotNuevo("Proceso Iniciado", true, file);
                                        WindowsDriver<WindowsElement> RootSession2 = null;
                                        RootSession2 = PruebaCRUD.RootSession();
                                        RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                        Thread.Sleep(10000);
                                        newFunctions_4.ScreenshotNuevo("Proceso Finalizado", true, file);

                                        
                                        WindowsDriver<WindowsElement> RootSession3 = null;
                                        RootSession3 = PruebaCRUD.RootSession();
                                        RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
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
                                    //break;

                                }
                                catch (Exception e)
                                {
                                    Thread.Sleep(4000);

                                    //Add Juan
                                    newFunctions_4.ScreenshotNuevo("Error en la prueba", true, file);

                                    //Add Daniel
                                    TestContext.AddResultFile(newFunctions_4.ScreenshotInicio(codProgram));
                                    //
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
                                    Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    break;
                                }
                            }
                        }
                        //break;
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
        public void VerificarRegistrosPrenomina()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesosLinom.ModuloProcesoLinom.VerificarRegistrosPrenomina")
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
                                string QbeFilter2 = rows["QbeFilter2"].ToString();
                                string concepto1 = rows["concepto1"].ToString();
                                string concepto2 = rows["concepto2"].ToString();
                                string concepto3 = rows["concepto3"].ToString();
                                string fechadesdeincap = rows["fechadesdeincap"].ToString();
                                string empresa = rows["empresa"].ToString();
                                String Codigo = "13";
                                String campo = "1";


                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "RegistrosPrenomina", motor);
                                
                               
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
                                        Thread.Sleep(5000);
                                        //QBE
                                        errors = FuncionesGlobales.QBEQry2Campos(desktopSession, TipoQbe, QbeFilter, QbeFilter2, file, campo, Codigo);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //Iniciar Proceso
                                        newFunctions_4.ScreenshotNuevo("Registros a Validar", true, file);
                                        FuncionesProcesosLinom.VerificarPrenomina(desktopSession, 0, crudPrincipal);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Registros", true, file);

                               

                                        //Validacion Salario
                                         string resultado= FuncionesVitales.ValidacionesSalario(0, motor, user, QbeFilter,empresa,concepto1);
                                        if (resultado.Contains("erronea"))
                                        {

                                            Assert.Fail("Error validacion salario");
                                        }
                                        //Validacion 2 dias incap
                                        string resultado1 = FuncionesVitales.ValidacionesIncapEmpresa(0, motor, user,QbeFilter,empresa,concepto2,QbeFilter2);
                                        if (resultado1.Contains("erronea"))
                                        {

                                            Assert.Fail("Error validacion incapacidad primeros dos dias");
                                        }

                                        //Validacion incap
                                        string incap = FuncionesVitales.ValidacionesIncapacidad(0, motor, user,QbeFilter,concepto3,fechadesdeincap,QbeFilter2);
                                        if (incap.Contains("erronea"))
                                        {

                                            Assert.Fail("Error validacion incapacidad");
                                        }

                                        //Validacion salud
                                        string salud = FuncionesVitales.ValidacionesSalud(0, motor, user, QbeFilter, concepto1, concepto2, concepto3, QbeFilter2);
                                        if (salud.Contains("erronea"))
                                        {

                                            Assert.Fail("Error validacion salud");
                                        }

                                        //Validacion pension
                                        string pension = FuncionesVitales.ValidacionesPension(0, motor, user, QbeFilter, concepto1, concepto2, concepto3, QbeFilter2);
                                        if (pension.Contains("erronea"))
                                        {

                                            Assert.Fail("Error validacion pension");
                                        }

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
                                    a.Stop();
                                    bandera = true;
                                    string delete1 = $"delete from nm_preno where cod_empl= '{QbeFilter}' and fec_acum='20230531'";
                                    FuncionesVitales.UpdateDeleteInsert(delete1, motor, user);

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
                                    //break;

                                }
                                catch (Exception e)
                                {
                                    Thread.Sleep(4000);

                                    //Add Juan
                                    newFunctions_4.ScreenshotNuevo("Error en la prueba", true, file);

                                    //Add Daniel
                                    TestContext.AddResultFile(newFunctions_4.ScreenshotInicio(codProgram));
                                    //
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
                                    Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    break;
                                }
                            }
                        }
                        //break;
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

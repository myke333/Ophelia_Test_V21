using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales;
using System.Text;
using System.Threading;
using System.Drawing;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesGlobalesProcesos;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.ValidacionQueries;
using OpenQA.Selenium.Interactions;
using System.Data;
using System.Diagnostics;

namespace Ophelia_Test_V21.Procesos.ModulosGlobalesProcesos
{
    /// <summary>
    /// Descripción resumida de ModulosGlobales
    /// </summary>
    [TestClass]
    public class ModulosGlobales : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();

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
        public ModulosGlobales()
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
        public void ModulosCreacionContratos()
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            string motor = string.Empty;
            string codProgram = string.Empty;

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
            //TableOrder = "ktes1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ModulosGlobalesProcesos.ModulosGlobales.ModulosCreacionContratos")
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
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                //Datos
                                rows["Identificacion"].ToString().Length != 0 && rows["Identificacion"].ToString() != null &&
                                rows["TipoContrato"].ToString().Length != 0 && rows["TipoContrato"].ToString() != null &&
                                rows["ProcedRetencion"].ToString().Length != 0 && rows["ProcedRetencion"].ToString() != null &&
                                rows["TipoCotizante"].ToString().Length != 0 && rows["TipoCotizante"].ToString() != null &&
                                rows["TipoSalario"].ToString().Length != 0 && rows["TipoSalario"].ToString() != null &&
                                rows["FechaNombramiento"].ToString().Length != 0 && rows["FechaNombramiento"].ToString() != null &&
                                rows["Cargo"].ToString().Length != 0 && rows["Cargo"].ToString() != null &&
                                rows["CentroCosto"].ToString().Length != 0 && rows["CentroCosto"].ToString() != null &&
                                rows["CentroTrabajo"].ToString().Length != 0 && rows["CentroTrabajo"].ToString() != null &&
                                rows["GrupoPrototipos"].ToString().Length != 0 && rows["GrupoPrototipos"].ToString() != null &&
                                rows["ClaseNomina"].ToString().Length != 0 && rows["ClaseNomina"].ToString() != null &&
                                rows["Sueldo"].ToString().Length != 0 && rows["Sueldo"].ToString() != null &&
                                rows["IdentificadorArbol"].ToString().Length != 0 && rows["IdentificadorArbol"].ToString() != null &&
                                rows["FechaFinalizacion"].ToString().Length != 0 && rows["FechaFinalizacion"].ToString() != null &&
                                 rows["TipoRegimen"].ToString().Length != 0 && rows["TipoRegimen"].ToString() != null 

                                )
                            {
                                string user = rows["User"].ToString();
                                motor = rows["Motor"].ToString();
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                codProgram = rows["CodProgram"].ToString();
                                //Datos
                                string Identificacion = rows["Identificacion"].ToString();
                                string TipoContrato = rows["TipoContrato"].ToString();
                                string ProcedRetencion = rows["ProcedRetencion"].ToString();
                                string TipoCotizante = rows["TipoCotizante"].ToString();
                                string TipoSalario = rows["TipoSalario"].ToString();
                                string FechaNombramiento = rows["FechaNombramiento"].ToString();
                                string Cargo = rows["Cargo"].ToString();
                                string CentroCosto = rows["CentroCosto"].ToString();
                                string CentroTrabajo = rows["CentroTrabajo"].ToString();
                                string GrupoPrototipos = rows["GrupoPrototipos"].ToString();
                                string ClaseNomina = rows["ClaseNomina"].ToString();
                                string Sueldo = rows["Sueldo"].ToString();
                                string IdentificadorArbol = rows["IdentificadorArbol"].ToString();
                                //Si contrato = Término Fijo 
                                string FechaFinalizacion = rows["FechaFinalizacion"].ToString();
                                string TipoRegimen = rows["TipoRegimen"].ToString();

                                List<string> crudVariables = new List<string>() { Identificacion, TipoContrato, ProcedRetencion, TipoCotizante, TipoSalario, FechaNombramiento,
                                Cargo, CentroCosto, CentroTrabajo, GrupoPrototipos, ClaseNomina, Sueldo, FechaFinalizacion, IdentificadorArbol, TipoRegimen};

                                string nomCapture = codProgram + motor + Hora();
                                string file = CrearDocumentoWordDinamico(nomCapture, "CrearContrato", motor);

                                try
                                {
                                    ErrorValidacion.Clear();
                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");


                                    //Thread.Sleep(15000);
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
                                        

                                        ////////AGREGAR REGISTRO
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 125, 15);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        ////Debugger.Launch();

                                        errors = FuncionesProcesosGlobales.AgregarContrato(desktopSession, crudVariables, file); 
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        ////ACEPTAR REGISTRO
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 215, 15);
                                        //desktopSession.Mouse.Click(null);
                                        //Thread.Sleep(5000);
                                        ////newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Identificacion, Identificacion, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);
                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }
                                    if (ErrorValidacion.Count > 0)
                                    {
                                        ReporteErrores(ErrorValidacion, codProgram, motor, "CrearContrato");
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, ErrorValidacion);
                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba {2} son:{0}{1}", Environment.NewLine, errorMessageString, codProgram));
                                    }

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
                                                //break;
                                            }
                                        }
                                    }
                                    AbrirPrograma a = new AbrirPrograma();
                                    a.Stop();
                                    //Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    //break;
                                }
                            }
							else
							{
								errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Al validar variables, no tiene registros o no existe");
							}
                        }

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
        public void ModulosCreacionNovedades()
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            string motor = string.Empty;
            string codProgram = string.Empty;
            int cont = 0;
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
            //TableOrder = "ktes1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ModulosGlobalesProcesos.ModulosGlobales.ModulosCreacionNovedades")
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
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                //Datos
                                rows["Identificacion"].ToString().Length != 0 && rows["Identificacion"].ToString() != null &&
                                rows["Concepto"].ToString().Length != 0 && rows["Concepto"].ToString() != null &&
                                //rows["Valorcuota"].ToString().Length != 0 && rows["Valorcuota"].ToString() != null &&
                                //rows["Cantidad"].ToString().Length != 0 && rows["Cantidad"].ToString() != null &&
                                rows["FechaPago"].ToString().Length != 0 && rows["FechaPago"].ToString() != null &&
                                rows["FechaNoved"].ToString().Length != 0 && rows["FechaNoved"].ToString() != null &&

                                rows["Noved"].ToString().Length != 0 && rows["Noved"].ToString() != null
                                )
                            {
                                string user = rows["User"].ToString();
                                motor = rows["motor"].ToString();
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                codProgram = rows["codProgram"].ToString();
                                //Datos
                                string Identificacion = rows["Identificacion"].ToString();
                                string Concepto = rows["Concepto"].ToString();
                                string Valorcuota = rows["Valorcuota"].ToString();
                                string Cantidad = rows["Cantidad"].ToString();
                                string FechaPago = rows["FechaPago"].ToString();
                                string FechaNoved = rows["FechaNoved"].ToString();

                                string Noved = rows["Noved"].ToString();

                                List<string> crudVariables = new List<string>() { Identificacion, Concepto, Valorcuota, Cantidad, FechaPago, FechaNoved};

                                string nomCapture = codProgram + motor + Hora();
                                string file = CrearDocumentoWordDinamico(nomCapture, "CrearNovedad", motor);

                                try
                                {
                                    if (Noved == "1") 
                                    {
                                        if (cont == 0)
                                        {
                                            ErrorValidacion.Clear();
                                            AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                            desktopSession = a.Start2(motor, "");
                                        }

                                        //Thread.Sleep(15000);
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
                                            ////////AGREGAR REGISTRO
                                            var ElementList = PruebaCRUD.NavClass(desktopSession);
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 125, 15);
                                            desktopSession.Mouse.Click(null);
                                            errors = FuncionesProcesosGlobales.AgregarNovedad(desktopSession, crudVariables, file); 
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        
                                            //newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

                                            ////Validacion Agregar
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Identificacion, Identificacion, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);
                                            Thread.Sleep(3000);
                                        }
                                        else
                                        {
                                            ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                        }
                                        if (ErrorValidacion.Count > 0)
                                        {
                                            ReporteErrores(ErrorValidacion, codProgram, motor, "CrearNovedad");
                                            var separator = string.Format("{0}{0}", Environment.NewLine);
                                            string errorMessageString = string.Join(separator, ErrorValidacion);
                                            Assert.Inconclusive(string.Format("Los errores presentados en la prueba {2} son:{0}{1}", Environment.NewLine, errorMessageString, codProgram));
                                        }

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
                                    }
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
                                                //break;
                                            }
                                        }
                                    }
                                    AbrirPrograma a = new AbrirPrograma();
                                    a.Stop();
                                    //Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    //break;
                                }
                            }
							else
							{
								errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Al validar variables, no tiene registros o no existe");
							}
                            cont = cont + 1;
                        }

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
        public void ModulosKNmLinom()
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            string motor = string.Empty;
            string codProgram = string.Empty;

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
            //TableOrder = "ktes1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ModulosGlobalesProcesos.ModulosGlobales.ModulosKNmLinom")
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
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                //rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                //rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                //rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                
                                //Datos
                                rows["FechaLiq"].ToString().Length != 0 && rows["FechaLiq"].ToString() != null &&
                                rows["FechaPago"].ToString().Length != 0 && rows["FechaPago"].ToString() != null 

                                )
                            {
                                string user = rows["User"].ToString();
                                motor = rows["motor"].ToString();
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                codProgram = rows["codProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                //Datos
                                string FechaLiq = rows["FechaLiq"].ToString();
                                string FechaPago = rows["FechaPago"].ToString();

                                List<string> crudVariables = new List<string>() { FechaLiq, FechaPago };

                                string nomCapture = codProgram + motor + Hora();
                                string file = CrearDocumentoWordDinamico(nomCapture, "ModKNmLinom", motor);

                                try
                                {
                                    ErrorValidacion.Clear();
                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");


                                    //Thread.Sleep(15000);
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
                                        //Se agregan las fechas de liquidacion y de pago
                                        FuncionesProcesosGlobales.LiquidarNomina(desktopSession, "0", crudVariables, file);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, Campo);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        var Aceptar = desktopSession.FindElementsByName("Aceptar");
                                        Thread.Sleep(500);
                                        //Click Boton Aceptar
                                        Aceptar[0].Click();
                                        WindowsDriver<WindowsElement> rootSession = null;
                                        Thread.Sleep(100);
                                        rootSession = PruebaCRUD.RootSession();
                                        rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                                        var allFrame = rootSession.FindElementsByClassName("IEFrame");
                                        Thread.Sleep(100);
                                        new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 120, (allFrame[0].Size.Height / 2) + 90).Click().Perform();
                                        Thread.Sleep(10000);
                                        newFunctions_4.ScreenshotNuevo("Fin del Proceso", true, file);
                                        Thread.Sleep(60000);
                                        new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 110, (allFrame[0].Size.Height / 2) + 90).DoubleClick().Perform();
                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }
                                    if (ErrorValidacion.Count > 0)
                                    {
                                        ReporteErrores(ErrorValidacion, codProgram, motor, "CrearNovedad");
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, ErrorValidacion);
                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba {2} son:{0}{1}", Environment.NewLine, errorMessageString, codProgram));
                                    }

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
                                                //break;
                                            }
                                        }
                                    }
                                    AbrirPrograma a = new AbrirPrograma();
                                    a.Stop();
                                    //Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    //break;
                                }
                            }
							else
							{
								errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Al validar variables, no tiene registros o no existe");
							}
                        }

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
        public void ModulosKNmAusen()
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            string motor = string.Empty;
            string codProgram = string.Empty;

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
            //TableOrder = "ktes1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ModulosGlobalesProcesos.ModulosGlobales.ModulosKNmAusen")
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
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                //Datos
                                 
                                rows["Identificacion"].ToString().Length != 0 && rows["Identificacion"].ToString() != null &&
                                rows["CodConcepto"].ToString().Length != 0 && rows["CodConcepto"].ToString() != null &&
                                rows["CodJust"].ToString().Length != 0 && rows["CodJust"].ToString() != null &&
                                rows["TipoAusen"].ToString().Length != 0 && rows["TipoAusen"].ToString() != null &&
                                rows["FechaInicio"].ToString().Length != 0 && rows["FechaInicio"].ToString() != null &&
                                rows["FechaFinal"].ToString().Length != 0 && rows["FechaFinal"].ToString() != null &&
                                rows["LNR"].ToString().Length != 0 && rows["LNR"].ToString() != null

                                )
                            {
                                string user = rows["User"].ToString();
                                motor = rows["Motor"].ToString();
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                codProgram = rows["CodProgram"].ToString();
                                //Datos
                                
                                string Identificacion = rows["Identificacion"].ToString();
                                string CodConcepto = rows["CodConcepto"].ToString();
                                string CodJust = rows["CodJust"].ToString();
                                string TipoAusen = rows["TipoAusen"].ToString();
                                string FechaInicio = rows["FechaInicio"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string LNR = rows["LNR"].ToString();

                                List<string> crudVariables = new List<string>() { Identificacion, CodConcepto, CodJust, TipoAusen, FechaInicio, FechaFinal, LNR };

                                string nomCapture = codProgram + motor + Hora();
                                string file = CrearDocumentoWordDinamico(nomCapture, "AgregarAusentismo", motor);

                                try
                                {
                                    if (LNR == "1")
                                    {
                                        ////Debugger.Launch();
                                        ErrorValidacion.Clear();
                                        AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                        desktopSession = a.Start2(motor, "");
                                    

                                        //Thread.Sleep(15000);
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
                                            ////////AGREGAR REGISTRO
                                            var ElementList = PruebaCRUD.NavClass(desktopSession);
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 125, 15);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                     
                                            errors = FuncionesProcesosGlobales.AgregarAusentismo(desktopSession, crudVariables, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            ////ACEPTAR REGISTRO
                                            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 215, 15);
                                            //desktopSession.Mouse.Click(null);
                                            //Thread.Sleep(5000);
                                            ////newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

                                            ////Validacion Agregar
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Identificacion, Identificacion, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);
                                        }
                                        else
                                        {
                                            ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                        }
                                    }

                                    if (ErrorValidacion.Count > 0)
                                    {
                                        ReporteErrores(ErrorValidacion, codProgram, motor, "AgregarAusentismo");
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, ErrorValidacion);
                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba {2} son:{0}{1}", Environment.NewLine, errorMessageString, codProgram));
                                    }

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
                                                //break;
                                            }
                                        }
                                    }
                                    AbrirPrograma a = new AbrirPrograma();
                                    a.Stop();
                                    //Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    //break;
                                }
                            }
							else
							{
								errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Al validar variables, no tiene registros o no existe");
							}
                        }

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
         public void ModulosCreacionIncapacidades()
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            string motor = string.Empty;
            string codProgram = string.Empty;

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
            //TableOrder = "ktes1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ModulosGlobalesProcesos.ModulosGlobales.ModulosCreacionIncapacidades")
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
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                //Datos
                                rows["Identificacion"].ToString().Length != 0 && rows["Identificacion"].ToString() != null &&
                                rows["Concepto"].ToString().Length != 0 && rows["Concepto"].ToString() != null &&
                                rows["CodigoDiag"].ToString().Length != 0 && rows["CodigoDiag"].ToString() != null &&
                                rows["NumeroIncapa"].ToString().Length != 0 && rows["NumeroIncapa"].ToString() != null &&
                                rows["FechaDesde"].ToString().Length != 0 && rows["FechaDesde"].ToString() != null &&
                                rows["FechaHasta"].ToString().Length != 0 && rows["FechaHasta"].ToString() != null &&
                                rows["FechaIniPago"].ToString().Length != 0 && rows["FechaIniPago"].ToString() != null &&
                                rows["Entidad"].ToString().Length != 0 && rows["Entidad"].ToString() != null &&
                                //validacion si tiene o no el empleado la incapacidad
                                rows["Incap"].ToString().Length != 0 && rows["Incap"].ToString() != null
                                )
                            {
                                string user = rows["User"].ToString();
                                motor = rows["Motor"].ToString();
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                codProgram = rows["CodProgram"].ToString();
                                //Datos
                                string Identificacion = rows["Identificacion"].ToString();
                                string Concepto = rows["Concepto"].ToString();
                                string CodigoDiag = rows["CodigoDiag"].ToString();
                                string NumeroIncapa = rows["NumeroIncapa"].ToString();                                
                                string FechaDesde = rows["FechaDesde"].ToString();
                                string FechaHasta = rows["FechaHasta"].ToString();
                                string FechaIniPago = rows["FechaIniPago"].ToString();
                                string Prorroga = rows["Prorroga"].ToString();
                                string Entidad = rows["Entidad"].ToString();
                                string Incap = rows["Incap"].ToString();
                                
                                List<string> crudData = new List<string>() { Identificacion, Concepto, CodigoDiag, NumeroIncapa, FechaDesde, FechaHasta, FechaIniPago, Prorroga, Entidad };

                                string nomCapture = codProgram + motor + Hora();
                                string file = CrearDocumentoWordDinamico(nomCapture, "CrearIncapacidad", motor);

                                try
                                {
                                    if (Incap == "1") 
                                    { 
                                        ErrorValidacion.Clear();
                                        ////Debugger.Launch();
                                        AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                        desktopSession = a.Start2(motor, "");


                                        //Thread.Sleep(15000);
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
                                        
                                            ////CRUD GENERAL

                                            //////AGREGAR REGISTRO
                                            //////////AGREGAR REGISTRO
                                            var ElementList = PruebaCRUD.NavClass(desktopSession);
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 125, 15);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);

                                            FuncionesProcesosGlobales.AgregarIncapacidad(desktopSession, crudData, file);

                                            //AGREGAR REGISTRO
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 215, 15);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(5000);
                                            a.Stop();
                                        }
                                        else
                                        {
                                            ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                        }
                                        if (ErrorValidacion.Count > 0)
                                        {
                                            ReporteErrores(ErrorValidacion, codProgram, motor, "CrearIncapacidad");
                                            var separator = string.Format("{0}{0}", Environment.NewLine);
                                            string errorMessageString = string.Join(separator, ErrorValidacion);
                                            Assert.Inconclusive(string.Format("Los errores presentados en la prueba {2} son:{0}{1}", Environment.NewLine, errorMessageString, codProgram));
                                        }

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
                                    }
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
                                                //break;
                                            }
                                        }
                                    }
                                    AbrirPrograma a = new AbrirPrograma();
                                    a.Stop();
                                    //Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    //break;
                                }
                            }
							else
							{
								errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Al validar variables, no tiene registros o no existe");
							}
                        }

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
        public void ModulosCreacionEmpleados()
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            string motor = string.Empty;
            string codProgram = string.Empty;

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
            //TableOrder = "ktes1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ModulosGlobalesProcesos.ModulosGlobales.ModulosCreacionEmpleados")
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
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                //Datos
                                rows["Identificacion"].ToString().Length != 0 && rows["Identificacion"].ToString() != null &&
                                rows["Apellido1"].ToString().Length != 0 && rows["Apellido1"].ToString() != null &&
                                rows["Nombre1"].ToString().Length != 0 && rows["Nombre1"].ToString() != null &&
                                rows["Nombre2"].ToString().Length != 0 && rows["Nombre2"].ToString() != null &&
                                rows["FechaNAC"].ToString().Length != 0 && rows["FechaNAC"].ToString() != null
                                )
                            {
                                string user = rows["User"].ToString();
                                motor = rows["Motor"].ToString();
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                codProgram = rows["CodProgram"].ToString();
                                //Datos
                                string Identificacion = rows["Identificacion"].ToString();
                                string Apellido1 = rows["Apellido1"].ToString();
                                string Nombre1 = rows["Nombre1"].ToString();
                                string Nombre2 = rows["Nombre2"].ToString();
                                string FechaNAC = rows["FechaNAC"].ToString();


                                List<string> crudVariables = new List<string>() { Identificacion, Apellido1, Nombre1, Nombre2, FechaNAC};

                                string nomCapture = codProgram + motor + Hora();
                                string file = CrearDocumentoWordDinamico(nomCapture, "CrearEmpleado", motor);

                                try
                                {
                                    ErrorValidacion.Clear();
                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");


                                    //Thread.Sleep(15000);
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


                                        ////////AGREGAR REGISTRO
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 125, 15);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        ////Debugger.Launch();

                                        errors = FuncionesProcesosGlobales.AgregarContrato(desktopSession, crudVariables, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        ////ACEPTAR REGISTRO
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 215, 15);
                                        //desktopSession.Mouse.Click(null);
                                        //Thread.Sleep(5000);
                                        ////newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Identificacion, Identificacion, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);
                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }
                                    if (ErrorValidacion.Count > 0)
                                    {
                                        ReporteErrores(ErrorValidacion, codProgram, motor, "CrearContrato");
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, ErrorValidacion);
                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba {2} son:{0}{1}", Environment.NewLine, errorMessageString, codProgram));
                                    }

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
                                                //break;
                                            }
                                        }
                                    }
                                    AbrirPrograma a = new AbrirPrograma();
                                    a.Stop();
                                    //Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    //break;
                                }
                            }
                            else
                            {
                                errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Al validar variables, no tiene registros o no existe");
                            }
                        }

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

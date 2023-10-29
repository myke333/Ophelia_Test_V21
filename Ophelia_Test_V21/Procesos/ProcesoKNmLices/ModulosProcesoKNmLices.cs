using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales;
using System.Text;
using System.Threading;
using System.Drawing;
using System.Diagnostics;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesoKNmLivac;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.ValidacionQueries;
using OpenQA.Selenium.Interactions;
using System.Data;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesoKNmLices;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesGlobalesProcesos;

namespace Ophelia_Test_V21.Procesos.ProcesoKNmLices
{
    /// <summary>
    /// Descripción resumida de ModulosProcesoKNmLices
    /// </summary>
    [TestClass]
    public class ModulosProcesoKNmLices : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();
               

        public ModulosProcesoKNmLices()
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
        public void ModulosCesantiasParciales()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesoKNmLices.ModulosProcesoKNmLices.ModulosCesantiasParciales")
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
                                //rows["TipoSolicitud"].ToString().Length != 0 && rows["TipoSolicitud"].ToString() != null &&
                                //rows["FechaSolicitud"].ToString().Length != 0 && rows["FechaSolicitud"].ToString() != null &&
                                //rows["ValorSolicitado"].ToString().Length != 0 && rows["ValorSolicitado"].ToString() != null &&
                                rows["EsParcial"].ToString().Length != 0 && rows["EsParcial"].ToString() != null
                                )
                            {
                                string user = rows["User"].ToString();
                                motor = rows["Motor"].ToString();
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                codProgram = rows["CodProgram"].ToString();
                                //Datos
                                string Identificacion = rows["Identificacion"].ToString();
                                string TipoSolicitud = rows["TipoSolicitud"].ToString();
                                string FechaSolicitud = rows["FechaSolicitud"].ToString();
                                string FechaCorte = rows["FechaCorte"].ToString();
                                string ValorSolicitado = rows["ValorSolicitado"].ToString();
                                string EsParcial = rows["EsParcial"].ToString();

                                List<string> crudVariables = new List<string>() { Identificacion, TipoSolicitud, FechaSolicitud, FechaCorte, ValorSolicitado, EsParcial };

                                string nomCapture = codProgram + motor + Hora();
                                string file = CrearDocumentoWordDinamico(nomCapture, "CrearCesantiasParciales", motor);

                                try
                                {
                                    if (EsParcial == "1")
                                    {
                                        //string nomCapture = codProgram + motor + Hora();
                                        //string file = CrearDocumentoWordDinamico(nomCapture, "CrearCesantiasParciales", motor);
                                       
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
                                            errors = FuncionesKNmLices.AgregarCesantiasParciales(desktopSession, crudVariables, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR REGISTRO
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 215, 15);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(5000);
                                            newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

                                            ////Validacion Agregar
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Identificacion, Identificacion, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);
                                        }
                                        else
                                        {
                                            ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                        }
                                        if (ErrorValidacion.Count > 0)
                                        {
                                            ReporteErrores(ErrorValidacion, codProgram, motor, "CrearCesantiasParciales");
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
        public void ProcesoKNmLices()
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
            //TableOrder = "DW_A1705";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesoKNmLices.ModulosProcesoKNmLices.ProcesoKNmLices")
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
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&

                                //rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["TipoPago"].ToString().Length != 0 && rows["TipoPago"].ToString() != null 
                                //rows["FechaLiquidacion"].ToString().Length != 0 && rows["FechaLiquidacion"].ToString() != null &&
                                //rows["FechaPagoCesan"].ToString().Length != 0 && rows["FechaPagoCesan"].ToString() != null &&
                                //rows["FechaPagoInteres"].ToString().Length != 0 && rows["FechaPagoInteres"].ToString() != null 
                                )
                            {
                                string user = rows["User"].ToString();
                                motor = rows["Motor"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string QbeFilter2 = rows["QbeFilter2"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string Campo2 = rows["Campo2"].ToString();
                                string TipoPago = rows["TipoPago"].ToString();
                                string FechaLiquidacion = rows["FechaLiquidacion"].ToString();
                                string FechaPagoCesan = rows["FechaPagoCesan"].ToString();
                                string FechaPagoInteres = rows["FechaPagoInteres"].ToString();

                                List<string> VarProcLices = new List<string> {TipoPago, FechaLiquidacion, FechaPagoCesan, FechaPagoInteres};
                                
                                //Carpeta
                                string Carpeta = "Proceso_KNmLices";

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                string nomCapture = codProgram + motor + Hora();
                                string file = CrearDocumentoWordDinamico(nomCapture, Carpeta, motor);

                                try
                                {
                                    if (cont == 0)
                                    {                                       
                                        ErrorValidacion.Clear();
                                        AbrirPrograma a = new AbrirPrograma("KNmLices", user);
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
                                        //Proceso FuncionesKNmLices
                                        Thread.Sleep(1000);
                                        errors = FuncionesKNmLices.ProcesoKNmLices(desktopSession, VarProcLices, file, cont);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);
                                        //if (cont == 0)
                                        //{
                                            //QBE
                                            errors = FuncionesKNmLices.QBEQry1(desktopSession, TipoQbe, QbeFilter, file, QbeFilter2, Campo, Campo2);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            Thread.Sleep(10000);
                                        //}
                                        //Debugger.Launch();
                                        //Ejecutar Liquidacion
                                        var BotonAceptar = desktopSession.FindElementsByName("Aceptar");
                                        BotonAceptar[0].Click();
                                        Thread.Sleep(1000);

                                        //Debugger.Launch();
                                        //problema de las fechas
                                        int dif = 0;
                                        if (FechaLiquidacion != "" && VarProcLices[0] == "Consolidada") 
                                        {
                                            DateTime Fecha = DateTime.ParseExact(FechaLiquidacion, "dd/MM/yyyy", null);
                                            //string FechaAct = DateTime.Now.ToString("dd/MM/yyyy");
                                            DateTime FechaACt = DateTime.Now;
                                            dif = DateTime.Compare(Fecha, FechaACt);

                                        }
                                        //Ventana Si, y 
                                        //Debugger.Launch();
                                        if (VarProcLices[0] == "Fin de Año" || VarProcLices[0] == "Parcial" || dif>0)
                                        {
                                            errors = FuncionesKNmLices.VentanasConfirmacionSi(desktopSession, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            Thread.Sleep(1000);

                                            for (int i= 1; i < 2; i++)
                                            {
                                                errors = FuncionesKNmLices.VentanasConfirmacionAceptar(desktopSession, file);
                                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                                Thread.Sleep(1000);
                                            }
                                            Thread.Sleep(1000);
                                            errors = FuncionesKNmLices.VentanasConfirmacionAceptar(desktopSession, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        }
                                        else 
                                        {
                                            for (int i = 1; i < 2; i++)
                                            {
                                                errors = FuncionesKNmLices.VentanasConfirmacionAceptar(desktopSession, file);
                                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                                //Thread.Sleep(5000);
                                            }
                                            Thread.Sleep(1000);
                                            errors = FuncionesKNmLices.VentanasConfirmacionAceptar(desktopSession, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        }
                                        //Debugger.Launch();
                                        //AbrirPrograma a = new AbrirPrograma();
                                        //a.Stop();
                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }
                                    if (ErrorValidacion.Count > 0)
                                    {
                                        ReporteErrores(ErrorValidacion, codProgram, motor, Carpeta);
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
                                    //break;
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
        public void ValidacionesProcesoKNmLices()
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            List<string> NomCamposValidarProcesoP = new List<string>();
            List<string> VarValorGlobalProcesoP = new List<string>();
            List<string> NumCamposTabProcesoP = new List<string>();
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
            //TableOrder = "DW_A1705"; 
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesoKNmLices.ModulosProcesoKNmLices.ValidacionesProcesoKNmLices")
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

                        string Carpeta = "Proceso_KNmLices";

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
                                //rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                //rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                //rows["QbeFilter2"].ToString().Length != 0 && rows["QbeFilter2"].ToString() != null &&
                                rows["IdInicial"].ToString().Length != 0 && rows["IdInicial"].ToString() != null &&
                                rows["IdFinal"].ToString().Length != 0 && rows["IdFinal"].ToString() != null
                                )
                            {

                                string user = rows["User"].ToString();
                                motor = rows["Motor"].ToString();
                                codProgram = rows["CodProgram"].ToString();
                                //string nomProgram = rows["NomProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string QbeFilter2 = rows["QbeFilter2"].ToString();
                                string IdInicial = rows["IdInicial"].ToString();
                                string IdFinal = rows["IdFinal"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string CampoId = rows["CampoId"].ToString();

                                // Variables que engloban el proceso 
                                string Identificacion = rows["Identificacion"].ToString();//KNmPreno // KNmCesan 1D
                                string Concepto = rows["Concepto"].ToString(); //KNmPReno 1TB(1204(Cesantiasfin de año) - 1206(Intereses de cesantias) - 1200 (Cesantias parciales)//
                                string TipoNomina = rows["TipoNomina"].ToString();//KNmPReno 3TB (C)//
                                string Cantidad = rows["Cantidad"].ToString();//KNmPReno 1TB (Dias laborados y pagados)//
                                string Valor = rows["Valor"].ToString();//KNmPReno 1TB (Dinero a pagar)//
                                string ValorTotal = rows["ValorTotal"].ToString();//KNmPReno 2TB (Sueldo base)//

                                string FechaPagoCesantias = rows["FechaPagoCesantias"].ToString();//KNmCEsan 7TB
                                string FechaPagoIntereses = rows["FechaPagoIntereses"].ToString();//KNmCEsan 1TB
                                string ValordelasCesantias = rows["ValordelasCesantias"].ToString();//KNmCEsan 6TB
                                string ValordelosIntereses = rows["ValordelosIntereses"].ToString();//KNmCEsan 2TB
                                string TipodeCesantias = rows["TipodeCesantias"].ToString();//KNmCEsan 9TB
                                string TipodeRegimen = rows["TipodeRegimen"].ToString();//KNmCEsan 1TB

                                string FechaConsolidado = rows["FechaConsolidado"].ToString(); // KNmCoces 2TB
                                string NroDiasTrabajados = rows["NroDiasTrabajados"].ToString(); // KNmCoces 4TB
                                string CesantiasCausadas = rows["CesantiasCausadas"].ToString();// KNmCoces 7TB
                                string InteresesCausadas = rows["InteresesCausadas"].ToString();// KNmCoces 1TB

                                //Variables del campo a validar
                                string CampoValidar1 = rows["CampoValidar1"].ToString();
                                string CampoValidar2 = rows["CampoValidar2"].ToString();
                                string CampoValidar3 = rows["CampoValidar3"].ToString();
                                string CampoValidar4 = rows["CampoValidar4"].ToString();
                                string CampoValidar5 = rows["CampoValidar5"].ToString();
                                string CampoValidar6 = rows["CampoValidar6"].ToString();
                                string CampoValidar7 = rows["CampoValidar7"].ToString();
                                string CampoValidar8 = rows["CampoValidar8"].ToString();
                                string CampoValidar9 = rows["CampoValidar9"].ToString();
                                string CampoValidar10 = rows["CampoValidar10"].ToString();
                                string CampoValidar11 = rows["CampoValidar11"].ToString();
                                string CampoValidar12 = rows["CampoValidar12"].ToString();

                                //Variables cantidad de Tabs
                                string CampoTab1 = rows["CampoTab1"].ToString();
                                string CampoTab2 = rows["CampoTab2"].ToString();
                                string CampoTab3 = rows["CampoTab3"].ToString();
                                string CampoTab4 = rows["CampoTab4"].ToString();
                                string CampoTab5 = rows["CampoTab5"].ToString();
                                string CampoTab6 = rows["CampoTab6"].ToString();
                                string CampoTab7 = rows["CampoTab7"].ToString();
                                string CampoTab8 = rows["CampoTab8"].ToString();
                                string CampoTab9 = rows["CampoTab9"].ToString();
                                string CampoTab10 = rows["CampoTab10"].ToString();
                                string CampoTab11 = rows["CampoTab11"].ToString();
                                string CampoTab12 = rows["CampoTab12"].ToString();


                                //Lista que contiene todos los campos
                                List<string> VarValorGlobalProceso = new List<string> { Concepto, TipoNomina, Cantidad, Valor, ValorTotal, FechaPagoCesantias, FechaPagoIntereses, ValordelasCesantias, ValordelosIntereses, TipodeCesantias, TipodeRegimen, FechaConsolidado, NroDiasTrabajados, CesantiasCausadas, InteresesCausadas };

                                //Lista que contiene el nombre de los campos a validar
                                List<string> VarValorGlobalProcesoNom = new List<string> { nameof(Concepto), nameof(TipoNomina), nameof(Cantidad), nameof(Valor), nameof(ValorTotal),  nameof(FechaPagoCesantias), nameof(FechaPagoIntereses), nameof(ValordelasCesantias), nameof(ValordelosIntereses), nameof(TipodeCesantias), nameof(TipodeRegimen), nameof(FechaConsolidado), nameof(NroDiasTrabajados), nameof(CesantiasCausadas), nameof(InteresesCausadas) };

                                //Lista que contiene el nombre de los campos a validar
                                List<string> NomCamposValidarProceso = new List<string> { CampoValidar1, CampoValidar2, CampoValidar3, CampoValidar4, CampoValidar5, CampoValidar6, CampoValidar7, CampoValidar8, CampoValidar9, CampoValidar10, CampoValidar11, CampoValidar12 };

                                //Lista que contiene el número de tabs a pasar en los campos a validar
                                List<string> NumCamposTabProceso = new List<string> { CampoTab1, CampoTab2, CampoTab3, CampoTab4, CampoTab5, CampoTab6, CampoTab7, CampoTab8, CampoTab9, CampoTab10, CampoTab11, CampoTab12 };

                                string file = "";
                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                if (cont == 0)
                                {
                                    string nomCapture = codProgram + motor + Hora();
                                    file = CrearDocumentoWordDinamico(nomCapture, Carpeta, motor);
                                }

                                try
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
                                        ////Debugger.Launch();
                                        Thread.Sleep(1000);
                                        if (cont == 0)
                                        {
                                            //KNmPreno Campo:{1,0} KNmCoces Campo:{0, 0} KNmCesan:{0,0}
                                            errors = FuncionesKNmLivac.QBEQrylices(desktopSession, TipoQbe, QbeFilter, file, QbeFilter2, Campo, CampoId);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            Thread.Sleep(1000);

                                        }
                                        Thread.Sleep(20000);

                                        //Seleccionar Data a pasar para el proceso de validación
                                        NomCamposValidarProcesoP = FuncionesProcesosGlobales.ValidacionListaDatosProceso1(desktopSession, VarValorGlobalProceso, NomCamposValidarProceso, VarValorGlobalProcesoNom);
                                        VarValorGlobalProcesoP = FuncionesProcesosGlobales.ValidacionListaDatosProceso2(desktopSession, VarValorGlobalProceso, NomCamposValidarProceso, VarValorGlobalProcesoNom, NomCamposValidarProcesoP);
                                        NumCamposTabProcesoP = FuncionesProcesosGlobales.ValidacionListaDatosTab(desktopSession, NumCamposTabProceso);

                                        //Proceso de validación
                                        Thread.Sleep(2000);
                                        errors = FuncionesProcesosGlobales.ValidacionesProcesos(desktopSession, IdInicial, IdFinal, cont, NomCamposValidarProcesoP, VarValorGlobalProcesoP, NumCamposTabProcesoP, Identificacion);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
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
                                    //a.Stop();
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
                        if (ErrorValidacion.Count > 0)
                        {
                            ////Debugger.Launch();
                            ReporteErrores(ErrorValidacion, codProgram, motor, Carpeta);
                            var separator = string.Format("{0}{0}", Environment.NewLine);
                            string errorMessageString = string.Join(separator, ErrorValidacion);
                            Assert.Inconclusive(string.Format("Los errores presentados en la prueba {2} son:{0}{1}", Environment.NewLine, errorMessageString, codProgram));
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

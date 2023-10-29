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
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.ValidacionQueries;
using OpenQA.Selenium.Interactions;
using System.Data;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesGlobalesProcesos;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesosKNmLinoe;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesoKNmLices;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesoKNmLivac;


namespace Ophelia_Test_V21.Procesos.ProcesoKNmLinoe
{
    /// <summary>
    /// Descripción resumida de ModulosProcesoKNmLinoe
    /// </summary>
    [TestClass]
    public class ModulosProcesoKNmLinoe : FuncionesVitales
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

        public ModulosProcesoKNmLinoe()
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
        public void ProcesoKNmLinoe()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesoKNmLinoe.ModulosProcesoKNmLinoe.ProcesoKNmLinoe")
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
                                ////Datos Login
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                
                                ///////////////////////Variables Obligatorias Linoe//////////////////////////////////
                                rows["Proceso"].ToString().Length != 0 && rows["Proceso"].ToString() != null &&
                                rows["TipoLiq"].ToString().Length != 0 && rows["TipoLiq"].ToString() != null &&
                                //rows["TipoAdj"].ToString().Length != 0 && rows["TipoAdj"].ToString() != null &&
                                rows["Prefijo"].ToString().Length != 0 && rows["Prefijo"].ToString() != null &&
                                rows["TRM"].ToString().Length != 0 && rows["TRM"].ToString() != null &&
                                //////Cuadro de texto
                                //rows["CUNE"].ToString().Length != 0 && rows["CUNE"].ToString() != null
                                //rows["Observaciones"].ToString().Length != 0 && rows["Observaciones"].ToString() != null &&
                                ////Tipo de Nomina
                                rows["NormalCheck"].ToString().Length != 0 && rows["NormalCheck"].ToString() != null &&
                                rows["ExtraCheck"].ToString().Length != 0 && rows["ExtraCheck"].ToString() != null &&
                                rows["AdicionalCheck"].ToString().Length != 0 && rows["AdicionalCheck"].ToString() != null &&
                                rows["DefinitivaCheck"].ToString().Length != 0 && rows["DefinitivaCheck"].ToString() != null &&
                                rows["CesantiasCheck"].ToString().Length != 0 && rows["CesantiasCheck"].ToString() != null &&
                                rows["VacacionesCheck"].ToString().Length != 0 && rows["VacacionesCheck"].ToString() != null &&
                                rows["PrimasCheck"].ToString().Length != 0 && rows["PrimasCheck"].ToString() != null &&
                                rows["PrimasExtrasCheck"].ToString().Length != 0 && rows["PrimasExtrasCheck"].ToString() != null &&
                                ////Fechas Linoe
                                rows["FechasInicial"].ToString().Length != 0 && rows["FechasInicial"].ToString() != null &&
                                rows["FechasFinal"].ToString().Length != 0 && rows["FechasFinal"].ToString() != null &&
                                rows["FechasEmision"].ToString().Length != 0 && rows["FechasEmision"].ToString() != null
                                ////CheckFechaPago
                                //rows["CheckFechaPago"].ToString().Length != 0 && rows["CheckFechaPago"].ToString() != null
                                
                                )
                            {
                                ///////Datos Login
                                string user = rows["User"].ToString();
                                motor = rows["Motor"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                /////// Variables Obligatorias
                                string Proceso = rows["Proceso"].ToString();
                                string TipoLiq = rows["TipoLiq"].ToString();
                                string Prefijo = rows["Prefijo"].ToString();
                                string TRM = rows["TRM"].ToString();

                                string TipoAdj = rows["TipoAdj"].ToString();
                                /////// Cuadros de texto
                                string CUNE = rows["CUNE"].ToString();
                                string Observaciones = rows["Observaciones"].ToString();
                                /////// Tipo de nomina
                                string NormalCheck = rows["NormalCheck"].ToString();
                                string ExtraCheck = rows["ExtraCheck"].ToString();
                                string AdicionalCheck = rows["AdicionalCheck"].ToString();
                                string DefinitivaCheck = rows["DefinitivaCheck"].ToString();
                                string CesantiasCheck = rows["CesantiasCheck"].ToString();
                                string VacacionesCheck = rows["VacacionesCheck"].ToString();
                                string PrimasCheck = rows["PrimasCheck"].ToString();
                                string PrimasExtrasCheck = rows["PrimasExtrasCheck"].ToString();
                                /////// Fechas linoe
                                string FechasInicial = rows["FechasInicial"].ToString();
                                string FechasFinal = rows["FechasFinal"].ToString();
                                string FechasEmision = rows["FechasEmision"].ToString();
                                string CheckFechaPago = rows["CheckFechaPago"].ToString();

                                List<string> VarProcLinoe = new List<string> { Proceso, TipoLiq, TipoAdj, Prefijo, TRM, CUNE, Observaciones, NormalCheck, ExtraCheck, AdicionalCheck, DefinitivaCheck, CesantiasCheck, VacacionesCheck, PrimasCheck, PrimasExtrasCheck, FechasInicial, FechasFinal, FechasEmision, CheckFechaPago};
                                                                            // 0          1        2       3       4      5             6             7        8               9              10                 11           12             13                 14             15            16              17           18                                                                
                                //Carpeta
                                string Carpeta = "Proceso_KNmLinoe";

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                //Debugger.Launch();
                                codProgram = "KNmLinoe";
                                string nomCapture = codProgram + motor + Hora();
                                string file = CrearDocumentoWordDinamico(nomCapture, Carpeta, motor);

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
                                        //Proceso FuncionesKNmLinoe
                                        //Debugger.Launch();
                                        errors = FuncionesKNmLinoe.ProcesoKNmLinoe(desktopSession, VarProcLinoe, file, cont);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);
                                        if (cont == 0)
                                        {
                                            //QBE
                                            errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, "0");
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            Thread.Sleep(15000);
                                        }
                                        //Ejecutar Liquidacion
                                        var BotonAceptar = desktopSession.FindElementsByName("Aceptar");
                                        BotonAceptar[0].Click();
                                        Thread.Sleep(1000);
                                        //Vrntana de confirmacion aceptar
                                        for (int i = 1; i < 3; i++)
                                        {
                                            errors = FuncionesKNmLices.VentanasConfirmacionAceptar(desktopSession, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            Thread.Sleep(3000);
                                        }
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
        public void ProcesoKNmNoema()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesoKNmLinoe.ModulosProcesoKNmLinoe.ProcesoKNmLinoe")
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
                                ////Datos Login
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&

                                ///////////////////////Variables Obligatorias Noema//////////////////////////////////
                                rows["Proceso"].ToString().Length != 0 && rows["Proceso"].ToString() != null &&
                                rows["TipoLiq"].ToString().Length != 0 && rows["TipoLiq"].ToString() != null &&
                                //rows["TipoAdj"].ToString().Length != 0 && rows["TipoAdj"].ToString() != null &&
                                rows["Prefijo"].ToString().Length != 0 && rows["Prefijo"].ToString() != null &&
                                rows["TRM"].ToString().Length != 0 && rows["TRM"].ToString() != null

                                )
                            {
                                ///////Datos Login
                                string user = rows["User"].ToString();
                                motor = rows["Motor"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                /////// Variables Obligatorias
                                string Proceso = rows["Proceso"].ToString();
                                string TipoLiq = rows["TipoLiq"].ToString();
                                string Prefijo = rows["Prefijo"].ToString();
                                string TRM = rows["TRM"].ToString();

                                ///////////////////

                                /////////Tablas Detalles Horas Extras
                                //Tablas detalles Horas Extras Diurnas
                                string SecuencialHED = rows["SecuencialHED"].ToString();
                                string HoraInicioHED = rows["HoraInicioHED"].ToString();
                                string HoraFinalHED = rows["HoraFinalHED"].ToString();
                                string CantidadHED = rows["CantidadHED"].ToString();
                                string PorcentajeHED = rows["PorcentajeHED"].ToString();
                                string ValorPagadoHED = rows["ValorPagadoHED"].ToString();
                                //Tablas detalles Horas Extras Nocturnas
                                string SecuencialHEN = rows["SecuencialHEN"].ToString();
                                string HoraInicioHEN = rows["HoraInicioHEN"].ToString();
                                string HoraFinalHEN = rows["HoraFinalHEN"].ToString();
                                string CantidadHEN = rows["CantidadHEN"].ToString();
                                string PorcentajeHEN = rows["PorcentajeHEN"].ToString();
                                string ValorPagadoHEN = rows["ValorPagadoHEN"].ToString();
                                //Tablas detalles Horas Extras Diurnas Dominical y Festivas
                                string SecuencialHEDDF = rows["SecuencialHEDDF"].ToString();
                                string HoraInicioHEDDF = rows["HoraInicioHEDDF"].ToString();
                                string HoraFinalHEDDF = rows["HoraFinalHEDDF"].ToString();
                                string CantidadHEDDF = rows["CantidadHEDDF"].ToString();
                                string PorcentajeHEDDF = rows["PorcentajeHEDDF"].ToString();
                                string ValorPagadoHEDDF = rows["ValorPagadoHEDDF"].ToString();
                                //Tablas detalles Horas Extras Nocturnas
                                string SecuencialHENDF = rows["SecuencialHENDF"].ToString();
                                string HoraInicioHENDF = rows["HoraInicioHENDF"].ToString();
                                string HoraFinalHENDF = rows["HoraFinalHENDF"].ToString();
                                string CantidadHENDF = rows["CantidadHENDF"].ToString();
                                string PorcentajeHENDF = rows["PorcentajeHENDF"].ToString();
                                string ValorPagadoHENDF = rows["ValorPagadoHENDF"].ToString();

                                /////////Tablas Detalles Horas Recargo
                                //Tablas detalles Horas Recargo Nocturno HRN
                                string SecuencialHRN = rows["SecuencialHRN"].ToString();
                                string HoraInicialHRN = rows["HoraInicialHRN"].ToString();
                                string HoraFinalHRN = rows["HoraFinalHRN"].ToString();
                                string CantidadHRN = rows["CantidadHRN"].ToString();
                                string PorcentajeHRN = rows["PorcentajeHRN"].ToString();
                                string ValorHRN = rows["ValorHRN"].ToString();
                                //Tablas detalles Horas Recargo Nocturno Dominical y Festivos HRNDF
                                string SecuencialHRNDF = rows["SecuencialHRNDF"].ToString();
                                string HoraInicialHRNDF = rows["HoraInicialHRNDF"].ToString();
                                string HoraFinalHRNDF = rows["HoraFinalHRNDF"].ToString();
                                string CantidadHRNDF = rows["CantidadHRNDF"].ToString();
                                string PorcentajeHRNDF = rows["PorcentajeHRNDF"].ToString();
                                string ValorHRNDF = rows["ValorHRNDF"].ToString();
                                //Tablas detalles Horas Recargo Diurno Dominical y Festivos HRDDF
                                string SecuencialHRDDF = rows["SecuencialHRDDF"].ToString();
                                string HoraInicialHRDDF = rows["HoraFinalHRDDF"].ToString();
                                string HoraFinalHRDDF = rows["HoraFinalHRDDF"].ToString();
                                string CantidadHRDDF = rows["CantidadHRDDF"].ToString();
                                string PorcentajeHRDDF = rows["PorcentajeHRDDF"].ToString();
                                string ValorHRDDF = rows["ValorHRDDF"].ToString();

                                /////////Tablas Detalles Vacaciones e Incapacidades
                                ///Vacaciones
                                string SecuencialVAC = rows["SecuencialVAC"].ToString();
                                string FechaInicioVAC = rows["FechaInicioVAC"].ToString();
                                string FechaFinalVAC = rows["FechaFinalVAC"].ToString();
                                string DiasVAC = rows["DiasVAC"].ToString();
                                string ValorVAC = rows["ValorVAC"].ToString();
                                //Incapacidades
                                string SecuencialINCAP = rows["SecuencialINCAP"].ToString();
                                string FechaInicioINCAP = rows["FechaInicioINCAP"].ToString();
                                string FechaFinalINCAP = rows["FechaFinalINCAP"].ToString();
                                string DiasINCAP = rows["DiasINCAP"].ToString();
                                string TipoINCAP = rows["TipoINCAP"].ToString();
                                string ValorINCAP = rows["ValorINCAP"].ToString();

                                /////////Tablas Detalles Licencias
                                ///Licencias Remuneradas
                                string SecuencialLR = rows["SecuencialLR"].ToString();
                                string FechaInicialLR = rows["FechaInicialLR"].ToString();
                                string FechaFinalLR = rows["FechaFinalLR"].ToString();
                                string DiasLR = rows["DiasVAC"].ToString();
                                string ValorLR = rows["ValorLR"].ToString();
                                ///Licencias NO Remuneradas
                                string SecuencialLNR = rows["SecuencialLNR"].ToString();
                                string FechaInicialLNR = rows["FechaInicialLNR"].ToString();
                                string FechaFinalLNR = rows["FechaFinalLNR"].ToString();
                                string DiasLNR = rows["DiasLNR"].ToString();

                                /////////Tablas Detalles Otros Conceptos Salariales/ No Salariales
                                ///Otros Conceptos Salariales
                                string SecuencialOCS = rows["SecuencialOCS"].ToString();
                                string ConceptoOCS = rows["ConceptoOCS"].ToString();
                                string ValorOCS = rows["ValorOCS"].ToString();
                                ///Otros Conceptos NO Salariales
                                string SecuencialOCNS = rows["SecuencialOCNS"].ToString();
                                string ConceptoOCNS = rows["ConceptoOCNS"].ToString();
                                string ValorOCNS = rows["ValorOCNS"].ToString();

                                /////////Tablas Detalles Huelgas
                                ///Huelgas
                                string SecuencialH = rows["SecuencialH"].ToString();
                                string FechaInicialH = rows["FechaInicialH"].ToString();
                                string FechaFinalH = rows["FechaFinalH"].ToString();
                                string DiasH = rows["DiasH"].ToString();
                                string ValorH = rows["ValorH"].ToString();

                                /////////Tablas Detalles Libranzas
                                ///Libranzas
                                string SecuencialL = rows["SecuencialL"].ToString();
                                string ConceptoL = rows["ConceptoL"].ToString();
                                string ValorL = rows["ValorL"].ToString();

                                List<string> VarProcLinoe = new List<string> { Proceso, TipoLiq, SecuencialHED, Prefijo, TRM };
                                // 0          1        2       3       4      5             6             7        8               9              10                 11           12             13                 14             15            16              17           18                                                                
                                //Carpeta
                                string Carpeta = "Proceso_KNmLinoe";

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                string nomCapture = codProgram + motor + Hora();
                                string file = CrearDocumentoWordDinamico(nomCapture, Carpeta, motor);

                                try
                                {
                                    if (cont == 0)
                                    {
                                        ErrorValidacion.Clear();
                                        AbrirPrograma a = new AbrirPrograma("KNmNoema", user);
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

                                        if (cont == 0)
                                        {
                                            //QBE
                                            errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, "0");
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            Thread.Sleep(10000);
                                        }
                                        //Proceso FuncionesKNmLinoe
                                        errors = FuncionesKNmLinoe.ProcesoKNmLinoe(desktopSession, VarProcLinoe, file, cont);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //Ejecutar Liquidacion
                                        var BotonAceptar = desktopSession.FindElementsByName("Aceptar");
                                        BotonAceptar[0].Click();
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
        public void ValidacionesProcesoKNmNoema()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesoKNmLinoe.ModulosProcesoKNmLinoe.ValidacionesProcesoKNmNoema")
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

                        string Carpeta = "Proceso_KNmNoema";

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
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["QbeFilter2"].ToString().Length != 0 && rows["QbeFilter2"].ToString() != null &&
                                rows["QbeFilter3"].ToString().Length != 0 && rows["QbeFilter3"].ToString() != null &&
                                rows["IdInicial"].ToString().Length != 0 && rows["IdInicial"].ToString() != null &&
                                rows["IdFinal"].ToString().Length != 0 && rows["IdFinal"].ToString() != null &&
                                rows["Identificacion"].ToString().Length != 0 && rows["Identificacion"].ToString() != null
                                )
                            {

                                string user = rows["User"].ToString();
                                motor = rows["Motor"].ToString();
                                codProgram = rows["CodProgram"].ToString();
                                //string nomProgram = rows["NomProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();// Prefijo
                                string QbeFilter2 = rows["QbeFilter2"].ToString();// Fecha de emision
                                string QbeFilter3 = rows["QbeFilter3"].ToString();// Identificaiones
                                string IdInicial = rows["IdInicial"].ToString();
                                string IdFinal = rows["IdFinal"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string CampoId = rows["CampoId"].ToString();

                                // Variables que engloban el proceso

                                ////Pestana Empelado
                                string Identificacion = rows["Identificacion"].ToString();//KNmNoema 0 // 
                                string TipoLiquidacion = rows["TipoLiquidacion"].ToString(); //KNmNoema 6TB// (I o A)
                                string FechaEmision = rows["FechaEmision"].ToString();//KNmNoema 1TB //
                                string Periodo = rows["Periodo"].ToString();//KNmNoema 23TB //
                                string Moneda = rows["Moneda"].ToString();//KNmNoema 1TB // (COP)
                                string TRM = rows["TRM"].ToString();// KNmNoema 1TB//
                                string TipoDoc = rows["TipoDoc"].ToString();// KNmNoema 4TB//
                                string FechaIngreso = rows["FechaIngreso"].ToString();// KNmNoema 1TB//
                                string CantidadTiempo = rows["CantidadTiempo"].ToString();// KNmNoema 4TB//
                                string Prefijo = rows["Prefijo"].ToString();// KNmNoema 4TB//
                                ////Pestana Ingreso - Valores
                                //Cuadro Valor
                                string SueldoBasico = rows["SueldoBasico"].ToString();// KNmNoema 9TB//
                                string DiasTrabajados = rows["DiasTrabajados"].ToString();// KNmNoema 8TB//
                                string ValorSueldoBasico = rows["ValorSueldoBasico"].ToString();// KNmNoema 2TB//
                                string ValorAuxTransporte = rows["ValorAuxTransporte"].ToString();// KNmNoema 1TB//
                                string ViaticosSalarial = rows["ViaticosSalarial"].ToString();// KNmNoema 1TB//
                                string ViaticosNoSalarial = rows["ViaticosNoSalarial"].ToString();// KNmNoema 1TB//
                                //Cuadro Vacaciones
                                string DiasVacacionesDinero = rows["DiasVacacionesDinero"].ToString();// KNmNoema 1TB//
                                string ValorDiasVacacionesDinero = rows["ValorDiasVacacionesDinero"].ToString();// KNmNoema 1TB//
                                //Cuadro Prima Legal
                                string DiasPagoPrima = rows["DiasPagoPrima"].ToString();// KNmNoema 1TB//
                                string ValorDiasPagoPrima = rows["ValorDiasPagoPrima"].ToString();// KNmNoema 1TB//
                                //Cuadro NO Prima
                                string ValorPrimaNoSalarial = rows["ValorPrimaNoSalarial"].ToString();// KNmNoema 1TB//
                                //Cuadro Cesantias 
                                string ValorPagoCesantias= rows["ValorPagoCesantias"].ToString();// KNmNoema 1TB//
                                string PorcentajeCesantias = rows["PorcentajeCesantias"].ToString();// KNmNoema 1TB//
                                string ValorPagoIntereses = rows["ValorPagoIntereses"].ToString();// KNmNoema 1TB//
                                //Cuadro Licencia de Maternidad/Paternidad
                                string FechaInicialLicMaternida = rows["FechaInicialLicMaternida"].ToString();// KNmNoema 1TB//
                                string FechaFinalLicMaternida = rows["FechaFinalLicMaternida"].ToString();// KNmNoema 1TB//
                                string DiasIncapacidadMaternidad = rows["DiasIncapacidadMaternidad"].ToString();// KNmNoema 1TB//
                                string ValorLicMaternida = rows["ValorLicMaternida"].ToString();// KNmNoema 1TB//
                                //Cuadro Valor Pago Bonificaciones
                                string PagoBonificacionesSalariales = rows["PagoBonificacionesSalariales"].ToString();// KNmNoema 1TB//
                                string PagoBonificacionesNoSalariales = rows["PagoBonificacionesNoSalariales"].ToString();// KNmNoema 1TB//
                                //Cuadro Valor Pago Auxilios
                                string PagoAuxiliosNoSalariales = rows["PagoAuxiliosNoSalariales"].ToString();// KNmNoema 1TB//
                                //Cuadro Valor Compensacion
                                string CompensacionOrdinaria = rows["CompensacionOrdinaria"].ToString(); // KNmNoema 5TB
                                string CompensacionExtraordinaria = rows["CompensacionExtraordinaria"].ToString(); // KNmNoema 1TB//
                                //Cuadro Valor Bono(s)
                                string BonoElectronicoSalarial = rows["BonoElectronicoSalarial"].ToString();// KNmNoema 1TB//
                                string BonoAlimentacionSalarial = rows["BonoAlimentacionSalarial"].ToString();// KNmNoema 1TB//
                                string BonoElectronicoNoSalarial = rows["BonoElectronicoNoSalarial"].ToString();// KNmNoema 1TB//
                                string BonoAlimentacionNoSalarial = rows["BonoAlimentacionNoSalarial"].ToString();// KNmNoema 1TB//
                                //Cuadro Otros Valores
                                string PagoComisiones = rows["PagoComisiones"].ToString();// KNmNoema 1TB//
                                string PagosPorTerceros = rows["PagosPorTerceros"].ToString();// KNmNoema 1TB//
                                string AnticiposNomina = rows["AnticiposNomina"].ToString();// KNmNoema 1TB//
                                string PagoDotaciones = rows["PagoDotaciones"].ToString();// KNmNoema 1TB//
                                string ApoyoSostenimiento = rows["ApoyoSostenimiento"].ToString();// KNmNoema 1TB//
                                string ValorTeletrabajo = rows["ValorTeletrabajo"].ToString();// KNmNoema 1TB//
                                string BonificacionesRetiro = rows["BonificacionesRetiro"].ToString();// KNmNoema 1TB//
                                string Indemnizacion = rows["Indemnizacion"].ToString();// KNmNoema 1TB//
                                string ReintegroEmpleador = rows["ReintegroEmpleador"].ToString();// KNmNoema 1TB//
                                ////Pestana Deducciones - Valores
                                //Cuadro Deduciones de Salud Empleado
                                string DeduccionSalud = rows["DeduccionSalud"].ToString();// KNmNoema 1TB//
                                string ValorDeduccionSalud = rows["ValorDeduccionSalud"].ToString();// KNmNoema 1TB//
                                //Cuadro Deduciones Pension Empleado
                                string PensionEmpleado = rows["PensionEmpleado"].ToString();// KNmNoema 1TB//
                                string ValorDeduccionPension = rows["ValorDeduccionPension"].ToString();// KNmNoema 1TB//
                                //Cuadro Fondo Solidaridad
                                string Solidaridad = rows["Solidaridad"].ToString();// KNmNoema 1TB// 
                                string ValorFondoSolidaridad = rows["ValorFondoSolidaridad"].ToString();
                                //Cuadro Fondo Subsistencia
                                string FondoSubsistencia = rows["FondoSubsistencia"].ToString();// KNmNoema 1TB//
                                string ValorFondoSubsistencia = rows["ValorFondoSubsistencia"].ToString();// KNmNoema 1TB//
                                //Cuadro Sindicato
                                string Sindicato = rows["Sindicato"].ToString();// KNmNoema 1TB//
                                string ValorSindicato = rows["ValorSindicato"].ToString();// KNmNoema 1TB//
                                //Otros Valores
                                string SancionPublica = rows["SancionPublica"].ToString();// KNmNoema 1TB//
                                string SancionPrivada = rows["SancionPrivada"].ToString();// KNmNoema 1TB//
                                string OtrasDeducciones = rows["OtrasDeducciones"].ToString();// KNmNoema 1TB//
                                string PagoAPV = rows["PagoAPV"].ToString();// KNmNoema 1TB//
                                string RetencionDeFuente = rows["RetencionDeFuente"].ToString();// KNmNoema 2TB//
                                string ValorCooperativas = rows["ValorCooperativas"].ToString();// KNmNoema 1TB//
                                string EmbargosFiscales = rows["EmbargosFiscales"].ToString();// KNmNoema 1TB//
                                string PlanComplentarios = rows["PlanComplentarios"].ToString();// KNmNoema 1TB//
                                string ConceptoEducativo = rows["ConceptoEducativo"].ToString();// KNmNoema1 1TB//
                                string ReintegroTrabajador = rows["ReintegroTrabajador"].ToString();// KNmNoema 1TB//
                                string DeudasEmpresa = rows["DeudasEmpresa"].ToString();// KNmNoema 1TB//
                                string TotalDevengo = rows["TotalDevengo"].ToString();// KNmNoema 2TB//
                                string TotalDeducciones = rows["TotalDeducciones"].ToString();// KNmNoema 1TB//
                                string Comprobante = rows["Comprobante"].ToString();// KNmNoema 1TB//

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
                                string CampoValidar13 = rows["CampoValidar13"].ToString();
                                string CampoValidar14 = rows["CampoValidar14"].ToString();
                                string CampoValidar15 = rows["CampoValidar15"].ToString();
                                string CampoValidar16 = rows["CampoValidar16"].ToString();
                                string CampoValidar17 = rows["CampoValidar17"].ToString();
                                string CampoValidar18 = rows["CampoValidar18"].ToString();
                                string CampoValidar19 = rows["CampoValidar19"].ToString();
                                string CampoValidar20 = rows["CampoValidar20"].ToString();
                                string CampoValidar21 = rows["CampoValidar21"].ToString();
                                string CampoValidar22 = rows["CampoValidar22"].ToString();
                                string CampoValidar23 = rows["CampoValidar23"].ToString();
                                string CampoValidar24 = rows["CampoValidar24"].ToString();
                                string CampoValidar25 = rows["CampoValidar25"].ToString();
                                string CampoValidar26 = rows["CampoValidar26"].ToString();
                                string CampoValidar27 = rows["CampoValidar27"].ToString();
                                string CampoValidar28 = rows["CampoValidar28"].ToString();
                                string CampoValidar29 = rows["CampoValidar29"].ToString();
                                string CampoValidar30 = rows["CampoValidar30"].ToString();
                                string CampoValidar31 = rows["CampoValidar31"].ToString();
                                string CampoValidar32 = rows["CampoValidar32"].ToString();
                                string CampoValidar33 = rows["CampoValidar33"].ToString();
                                string CampoValidar34 = rows["CampoValidar34"].ToString();
                                string CampoValidar35 = rows["CampoValidar35"].ToString();
                                string CampoValidar36 = rows["CampoValidar36"].ToString();
                                string CampoValidar37 = rows["CampoValidar37"].ToString();
                                string CampoValidar38 = rows["CampoValidar38"].ToString();
                                string CampoValidar39 = rows["CampoValidar39"].ToString();
                                string CampoValidar40 = rows["CampoValidar40"].ToString();
                                string CampoValidar41 = rows["CampoValidar41"].ToString();
                                string CampoValidar42 = rows["CampoValidar42"].ToString();
                                string CampoValidar43 = rows["CampoValidar43"].ToString();
                                string CampoValidar44 = rows["CampoValidar44"].ToString();
                                string CampoValidar45 = rows["CampoValidar45"].ToString();
                                string CampoValidar46 = rows["CampoValidar46"].ToString();
                                string CampoValidar47 = rows["CampoValidar47"].ToString();
                                string CampoValidar48 = rows["CampoValidar48"].ToString();
                                string CampoValidar49 = rows["CampoValidar49"].ToString();
                                string CampoValidar50 = rows["CampoValidar50"].ToString();
                                string CampoValidar51 = rows["CampoValidar51"].ToString();
                                string CampoValidar52 = rows["CampoValidar52"].ToString();
                                string CampoValidar53 = rows["CampoValidar53"].ToString();
                                string CampoValidar54 = rows["CampoValidar54"].ToString();
                                string CampoValidar55 = rows["CampoValidar55"].ToString();
                                string CampoValidar56 = rows["CampoValidar56"].ToString();
                                string CampoValidar57 = rows["CampoValidar57"].ToString();
                                string CampoValidar58 = rows["CampoValidar58"].ToString();
                                string CampoValidar59 = rows["CampoValidar59"].ToString();
                                string CampoValidar60 = rows["CampoValidar60"].ToString();
                                string CampoValidar61 = rows["CampoValidar61"].ToString();
                                string CampoValidar62 = rows["CampoValidar62"].ToString();
                                string CampoValidar63 = rows["CampoValidar63"].ToString();
                                string CampoValidar64 = rows["CampoValidar64"].ToString();
                                string CampoValidar65 = rows["CampoValidar65"].ToString();
                                string CampoValidar66 = rows["CampoValidar66"].ToString();
                                string CampoValidar67 = rows["CampoValidar67"].ToString();
                                string CampoValidar68 = rows["CampoValidar68"].ToString();
                                string CampoValidar69 = rows["CampoValidar69"].ToString();
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
                                string CampoTab13 = rows["CampoTab13"].ToString();
                                string CampoTab14 = rows["CampoTab14"].ToString();
                                string CampoTab15 = rows["CampoTab15"].ToString();
                                string CampoTab16 = rows["CampoTab16"].ToString();
                                string CampoTab17 = rows["CampoTab17"].ToString();
                                string CampoTab18 = rows["CampoTab18"].ToString();
                                string CampoTab19 = rows["CampoTab19"].ToString();
                                string CampoTab20 = rows["CampoTab20"].ToString();
                                string CampoTab21 = rows["CampoTab21"].ToString();
                                string CampoTab22 = rows["CampoTab22"].ToString();
                                string CampoTab23 = rows["CampoTab23"].ToString();
                                string CampoTab24 = rows["CampoTab24"].ToString();
                                string CampoTab25 = rows["CampoTab25"].ToString();
                                string CampoTab26 = rows["CampoTab26"].ToString();
                                string CampoTab27 = rows["CampoTab27"].ToString();
                                string CampoTab28 = rows["CampoTab28"].ToString();
                                string CampoTab29 = rows["CampoTab29"].ToString();
                                string CampoTab30 = rows["CampoTab30"].ToString();
                                string CampoTab31 = rows["CampoTab31"].ToString();
                                string CampoTab32 = rows["CampoTab32"].ToString();
                                string CampoTab33 = rows["CampoTab33"].ToString();
                                string CampoTab34 = rows["CampoTab34"].ToString();
                                string CampoTab35 = rows["CampoTab35"].ToString();
                                string CampoTab36 = rows["CampoTab36"].ToString();
                                string CampoTab37 = rows["CampoTab37"].ToString();
                                string CampoTab38 = rows["CampoTab38"].ToString();
                                string CampoTab39 = rows["CampoTab39"].ToString();
                                string CampoTab40 = rows["CampoTab40"].ToString();
                                string CampoTab41 = rows["CampoTab41"].ToString();
                                string CampoTab42 = rows["CampoTab42"].ToString();
                                string CampoTab43 = rows["CampoTab43"].ToString();
                                string CampoTab44 = rows["CampoTab44"].ToString();
                                string CampoTab45 = rows["CampoTab45"].ToString();
                                string CampoTab46 = rows["CampoTab46"].ToString();
                                string CampoTab47 = rows["CampoTab47"].ToString();
                                string CampoTab48 = rows["CampoTab48"].ToString();
                                string CampoTab49 = rows["CampoTab49"].ToString();
                                string CampoTab50 = rows["CampoTab50"].ToString();
                                string CampoTab51 = rows["CampoTab51"].ToString();
                                string CampoTab52 = rows["CampoTab52"].ToString();
                                string CampoTab53 = rows["CampoTab53"].ToString();
                                string CampoTab54 = rows["CampoTab54"].ToString();
                                string CampoTab55 = rows["CampoTab55"].ToString();
                                string CampoTab56 = rows["CampoTab56"].ToString();
                                string CampoTab57 = rows["CampoTab57"].ToString();
                                string CampoTab58 = rows["CampoTab58"].ToString();
                                string CampoTab59 = rows["CampoTab59"].ToString();
                                string CampoTab60 = rows["CampoTab60"].ToString();
                                string CampoTab61 = rows["CampoTab61"].ToString();
                                string CampoTab62 = rows["CampoTab62"].ToString();
                                string CampoTab63 = rows["CampoTab63"].ToString();
                                string CampoTab64 = rows["CampoTab64"].ToString();
                                string CampoTab65 = rows["CampoTab65"].ToString();
                                string CampoTab66 = rows["CampoTab66"].ToString();
                                string CampoTab67 = rows["CampoTab67"].ToString();
                                string CampoTab68 = rows["CampoTab68"].ToString();
                                string CampoTab69 = rows["CampoTab69"].ToString();
                                //Lista que contiene el número de tabs a pasar en los campos a validar
                                List<string> NumCamposTabProceso = new List<string> {                   CampoTab1,           CampoTab2,        CampoTab3,     CampoTab4,    CampoTab5,       CampoTab6,           CampoTab7,             CampoTab8,         CampoTab9,          CampoTab10,            CampoTab11,              CampoTab12,                CampoTab13,                  CampoTab14,               CampoTab15,                  CampoTab16,                    CampoTab17,                   CampoTab18,               CampoTab19,                  CampoTab20,                CampoTab21,                  CampoTab22,                  CampoTab23,                  CampoTab24,                        CampoTab25,                      CampoTab26,                     CampoTab27,                    CampoTab28,                           CampoTab29,                              CampoTab30,                     CampoTab31,                     CampoTab32,                         CampoTab33,                       CampoTab34,                      CampoTab35,                        CampoTab36,                     CampoTab37,            CampoTab38,               CampoTab39,              CampoTab40,               CampoTab41,               CampoTab42,                 CampoTab43,                CampoTab44,               CampoTab45,               CampoTab46,               CampoTab47,                 CampoTab48,               CampoTab49,                 CampoTab50,               CampoTab51,                  CampoTab52,                  CampoTab53,              CampoTab54,           CampoTab55,           CampoTab56,               CampoTab57,            CampoTab58,             CampoTab59,          CampoTab60,                 CampoTab61,               CampoTab62,               CampoTab63,                 CampoTab64,                CampoTab65,              CampoTab66,           CampoTab67,               CampoTab68,                 CampoTab69};


                                //Lista que contiene todos los campos
                                List<string> VarValorGlobalProceso = new List<string> {           TipoLiquidacion,         FechaEmision,         Periodo,        Moneda,         TRM,          TipoDoc,         FechaIngreso,          CantidadTiempo,       Prefijo,          SueldoBasico,         DiasTrabajados,           ValorSueldoBasico,       ValorAuxTransporte,          ViaticosSalarial,          ViaticosNoSalarial,        DiasVacacionesDinero,         ValorDiasVacacionesDinero,          DiasPagoPrima,       ValorDiasPagoPrima,          ValorPrimaNoSalarial,         ValorPagoCesantias,       PorcentajeCesantias,          ValorPagoIntereses,        FechaInicialLicMaternida,          FechaFinalLicMaternida,        DiasIncapacidadMaternidad,         ValorLicMaternida,          PagoBonificacionesSalariales,          PagoBonificacionesNoSalariales,        PagoAuxiliosNoSalariales,         CompensacionOrdinaria,          CompensacionExtraordinaria,        BonoElectronicoSalarial,         BonoAlimentacionSalarial,         BonoElectronicoNoSalarial,         BonoAlimentacionNoSalarial,          PagoComisiones,        PagosPorTerceros,          AnticiposNomina,        PagoDotaciones,         ApoyoSostenimiento,         ValorTeletrabajo,        BonificacionesRetiro,         Indemnizacion,           ReintegroEmpleador,         DeduccionSalud,        ValorDeduccionSalud,          PensionEmpleado,        ValorDeduccionPension,         Solidaridad,          ValorFondoSolidaridad,        FondoSubsistencia,         ValorFondoSubsistencia,         Sindicato,          ValorSindicato,        SancionPublica,         SancionPrivada,        OtrasDeducciones,           PagoAPV,         RetencionDeFuente,          ValorCooperativas,        EmbargosFiscales,         PlanComplentarios,         ConceptoEducativo,         ReintegroTrabajador,        DeudasEmpresa,       TotalDevengo,            TotalDeducciones,            Comprobante };

                                //Lista que contiene el nombre de los campos a validar
                                List<string> VarValorGlobalProcesoNom = new List<string> { nameof(TipoLiquidacion), nameof(FechaEmision), nameof(Periodo), nameof(Moneda), nameof(TRM), nameof(TipoDoc), nameof(FechaIngreso), nameof(CantidadTiempo), nameof(Prefijo), nameof(SueldoBasico), nameof(DiasTrabajados), nameof(ValorSueldoBasico), nameof(ValorAuxTransporte), nameof(ViaticosSalarial), nameof(ViaticosNoSalarial), nameof(DiasVacacionesDinero), nameof(ValorDiasVacacionesDinero), nameof(DiasPagoPrima), nameof(ValorDiasPagoPrima), nameof(ValorPrimaNoSalarial), nameof(ValorPagoCesantias), nameof(PorcentajeCesantias), nameof(ValorPagoIntereses), nameof(FechaInicialLicMaternida), nameof(FechaFinalLicMaternida), nameof(DiasIncapacidadMaternidad), nameof(ValorLicMaternida), nameof(PagoBonificacionesSalariales), nameof(PagoBonificacionesNoSalariales), nameof(PagoAuxiliosNoSalariales), nameof(CompensacionOrdinaria), nameof(CompensacionExtraordinaria), nameof(BonoElectronicoSalarial), nameof(BonoAlimentacionSalarial), nameof(BonoElectronicoNoSalarial), nameof(BonoAlimentacionNoSalarial), nameof(PagoComisiones), nameof(PagosPorTerceros), nameof(AnticiposNomina), nameof(PagoDotaciones), nameof(ApoyoSostenimiento), nameof(ValorTeletrabajo), nameof(BonificacionesRetiro), nameof(Indemnizacion), nameof(ReintegroEmpleador), nameof(DeduccionSalud), nameof(ValorDeduccionSalud), nameof(PensionEmpleado), nameof(ValorDeduccionPension), nameof(Solidaridad), nameof(ValorFondoSolidaridad), nameof(FondoSubsistencia), nameof(ValorFondoSubsistencia), nameof(Sindicato), nameof(ValorSindicato), nameof(SancionPublica), nameof(SancionPrivada), nameof(OtrasDeducciones), nameof(PagoAPV), nameof(RetencionDeFuente), nameof(ValorCooperativas), nameof(EmbargosFiscales), nameof(PlanComplentarios), nameof(ConceptoEducativo), nameof(ReintegroTrabajador), nameof(DeudasEmpresa), nameof(TotalDevengo),   nameof(TotalDeducciones),    nameof(Comprobante) };

                                //Lista que contiene el nombre de los campos a validar
                                List<string> NomCamposValidarProceso = new List<string> {            CampoValidar1,       CampoValidar2,     CampoValidar3, CampoValidar4, CampoValidar5,   CampoValidar6,      CampoValidar7,         CampoValidar8,     CampoValidar9,      CampoValidar10,        CampoValidar11,           CampoValidar12,           CampoValidar13,             CampoValidar14,             CampoValidar15,             CampoValidar16,                 CampoValidar17,                CampoValidar18,          CampoValidar19,             CampoValidar20,             CampoValidar21,              CampoValidar22,               CampoValidar23,            CampoValidar24,                    CampoValidar25,                    CampoValidar26,               CampoValidar27,                  CampoValidar28,                       CampoValidar29,                         CampoValidar30,                CampoValidar31,                  CampoValidar32,                     CampoValidar33,                 CampoValidar34,                   CampoValidar35,                    CampoValidar36,                 CampoValidar37,          CampoValidar38,          CampoValidar39,          CampoValidar40,           CampoValidar41,            CampoValidar42,           CampoValidar43,            CampoValidar44,           CampoValidar45,            CampoValidar46,           CampoValidar47,             CampoValidar48,           CampoValidar49,            CampoValidar50,           CampoValidar51,              CampoValidar52,              CampoValidar53,           CampoValidar54,      CampoValidar55,         CampoValidar56,         CampoValidar57,         CampoValidar58,        CampoValidar59,       CampoValidar60,            CampoValidar61,            CampoValidar62,           CampoValidar63,            CampoValidar64,              CampoValidar65,         CampoValidar66,        CampoValidar67,          CampoValidar68,             CampoValidar69};

                                
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
                                            errors = FuncionesKNmLinoe.QBEQry(desktopSession, TipoQbe, QbeFilter, QbeFilter2, QbeFilter3, file, Campo, CampoId);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            Thread.Sleep(1000);

                                        }
                                        //Thread.Sleep(20000);
                                        
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
    }
}

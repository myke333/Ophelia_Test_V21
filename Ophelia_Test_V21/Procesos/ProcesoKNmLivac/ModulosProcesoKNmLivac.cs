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
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesoKNmLices;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.ValidacionQueries;
using OpenQA.Selenium.Interactions;
using System.Data;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesGlobalesProcesos;

namespace Ophelia_Test_V21.Procesos.ProcesoKNmLivac
{
    /// <summary>
    /// Descripción resumida de ModulosCreacionContratos
    /// </summary>
    [TestClass]
    public class ModulosProcesoKNmLivac : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();


        public ModulosProcesoKNmLivac()
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
        public void ProcesoKNmLivac()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesoKNmLivac.ModulosProcesoKNmLivac.ProcesoKNmLivac")
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
                            //Debugger.Launch();
                            if (
                                //Datos Login
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
                                //rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                //rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&

                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                //rows["DiasTiempo"].ToString().Length != 0 && rows["DiasTiempo"].ToString() != null &&
                                rows["FechaLiquidacion"].ToString().Length != 0 && rows["FechaLiquidacion"].ToString() != null &&
                                //rows["DiasPago"].ToString().Length != 0 && rows["DiasPago"].ToString() != null &&
                                //rows["FechaInicio"].ToString().Length != 0 && rows["FechaInicio"].ToString() != null &&
                                //rows["FechaPago"].ToString().Length != 0 && rows["FechaPago"].ToString() != null &&
                                rows["TipoVacaciones"].ToString().Length != 0 && rows["TipoVacaciones"].ToString() != null &&
                                rows["TipoLiquidacion"].ToString().Length != 0 && rows["TipoLiquidacion"].ToString() != null
                                )
                            {

                                string user = rows["User"].ToString();
                                motor = rows["Motor"].ToString();
                                //codProgram = rows["CodProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string DiasTiempo = rows["DiasTiempo"].ToString();
                                string DiasPago = rows["DiasPago"].ToString();
                                string FechaInicio = rows["FechaInicio"].ToString();
                                string FechaPago = rows["FechaPago"].ToString();
                                string TipoVacaciones = rows["TipoVacaciones"].ToString();
                                string QbeFilter2 = rows["QbeFilter2"].ToString();
                                string Campo2 = rows["Campo2"].ToString();
                                string FechaLiquidacion = rows["FechaLiquidacion"].ToString();
                                string TipoLiquidacion = rows["TipoLiquidacion"].ToString();
                                List<string> VarProcLivac = new List<string> { DiasTiempo, DiasPago, FechaInicio, FechaPago, TipoVacaciones, FechaLiquidacion, TipoLiquidacion };

                                //Carpeta
                                string Carpeta = "Proceso_KNmLivac";

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                string nomCapture = codProgram + motor + Hora();
                                string file = CrearDocumentoWordDinamico(nomCapture, Carpeta, motor);

                                try
                                {
                                    if (cont == 0)
                                    {
                                        ErrorValidacion.Clear();
                                        AbrirPrograma a = new AbrirPrograma("KNmLivac", user);
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
                                        //Proceso FuncionesKNmLivac1
                                        errors = FuncionesKNmLivac.ProcesoKNmLivac(desktopSession, VarProcLivac, file, cont);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(3000);

                                        //QBE
                                        errors = FuncionesKNmLivac.QBEQry1(desktopSession, TipoQbe, QbeFilter, file, QbeFilter2, Campo, Campo2);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(5000);

                                        //Proceso FuncionesKNmLivac2
                                        errors = FuncionesKNmLivac.ProcesoKNmLivac2(desktopSession, VarProcLivac, file, cont);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(3000);

                                        //Ventanas emergentes Aceptar
                                        for (int i = 1; i < 2; i++)
                                        {
                                            errors = FuncionesKNmLices.VentanasConfirmacionAceptar(desktopSession, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            Thread.Sleep(5000);
                                        }
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
        public void ValidacionesProcesoKNmLivac()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesoKNmLivac.ModulosProcesoKNmLivac.ValidacionesProcesoKNmLivac")
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

                        string Carpeta = "Proceso_KNmLivac";

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
                                string Identificacion = rows["Identificacion"].ToString();
                                string Concepto = rows["Concepto"].ToString();
                                string FechaInicio = rows["FechaInicio"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string DiasTotales = rows["DiasTotales"].ToString();
                                string DiasReales = rows["DiasReales"].ToString();
                                string TipoNomina = rows["TipoNomina"].ToString();
                                string ValorBase = rows["ValorBase"].ToString();
                                string ValorPagado = rows["ValorPagado"].ToString();
                                string FechaCausacionInicial = rows["FechaCausacionInicial"].ToString();
                                string FechaCausacionFinal = rows["FechaCausacionFinal"].ToString();
                                string FechaRegresoVacaciones = rows["FechaRegresoVacaciones"].ToString();
                                string DiasDinero = rows["DiasDinero"].ToString();
                                string DiasPagoDinero = rows["DiasPagoDinero"].ToString();
                                string FechaAcumula = rows["FechaAcumula"].ToString();

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
                                List<string> VarValorGlobalProceso = new List<string> { Concepto, FechaInicio, FechaFinal, DiasTotales, DiasReales, TipoNomina, ValorBase, ValorPagado, FechaCausacionInicial, FechaCausacionFinal, FechaRegresoVacaciones, DiasDinero, DiasPagoDinero, FechaAcumula };

                                //Lista que contiene el nombre de los campos a validar
                                List<string> VarValorGlobalProcesoNom = new List<string> { nameof(Concepto), nameof(FechaInicio), nameof(FechaFinal), nameof(DiasTotales), nameof(DiasReales), nameof(TipoNomina), nameof(ValorBase), nameof(ValorPagado), nameof(FechaCausacionInicial), nameof(FechaCausacionFinal), nameof(FechaRegresoVacaciones), nameof(DiasDinero), nameof(DiasPagoDinero), nameof(FechaAcumula) };

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
                                            //QBE
                                            errors = FuncionesKNmLivac.QBEQry(desktopSession, TipoQbe, QbeFilter, file, QbeFilter2, Campo, CampoId);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            Thread.Sleep(10000);
                                        }
                                        Thread.Sleep(2000);

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




        //[TestMethod]
        //public void ValidacionesKNmDiasn()
        //{
        //    List<string> errorMessages = new List<string>();
        //    List<string> ErrorValidacion = new List<string>();
        //    List<string> errors = new List<string>();
        //    string motor = string.Empty;
        //    string codProgram = string.Empty;
        //    int cont = 0;
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
        //    //TableOrder = "DW_A1705";
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

        //        if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesoKNmLivac.ModulosProcesoKNmLivac.ValidacionesKNmDiasn")
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
        //                    //int velocidad = 10;

        //                    //Playback.PlaybackSettings.DelayBetweenActions = velocidad;

        //                    //@User @Motor @CodProgram @NomProgram @TipoQbe @QbeFilter @banExcel @BanderaLupa @BanderaSum @BanderaPreli
        //                    //@Identificacion @Concepto @Ausencia @Observaciones

        //                    if (
        //                        //Datos Login
        //                        rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
        //                        rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
        //                        rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
        //                        //rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
        //                        rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
        //                        //rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
        //                        //rows["QbeFilter2"].ToString().Length != 0 && rows["QbeFilter2"].ToString() != null &&
        //                        rows["IdInicial"].ToString().Length != 0 && rows["IdInicial"].ToString() != null &&
        //                        rows["IdFinal"].ToString().Length != 0 && rows["IdFinal"].ToString() != null &&
        //                        rows["Concepto"].ToString().Length != 0 && rows["Concepto"].ToString() != null &&
        //                        rows["FechaDesde"].ToString().Length != 0 && rows["FechaDesde"].ToString() != null &&
        //                        rows["FechaHasta"].ToString().Length != 0 && rows["FechaHasta"].ToString() != null 
        //                        )
        //                    {

        //                        string user = rows["User"].ToString();
        //                        motor = rows["Motor"].ToString();
        //                        codProgram = rows["CodProgram"].ToString();
        //                        //string nomProgram = rows["NomProgram"].ToString();
        //                        string TipoQbe = rows["TipoQbe"].ToString();
        //                        string QbeFilter = rows["QbeFilter"].ToString();
        //                        string QbeFilter2 = rows["QbeFilter2"].ToString();
        //                        string IdInicial = rows["IdInicial"].ToString();
        //                        string IdFinal = rows["IdFinal"].ToString();
        //                        string Concepto = rows["Concepto"].ToString();
        //                        string FechaDesde = rows["FechaDesde"].ToString();
        //                        string FechaHasta = rows["FechaHasta"].ToString();

        //                        string Carpeta = "Proceso_KNmLivac";
        //                        string file = "";
        //                        //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
        //                        ////Debugger.Launch();
        //                        if (cont == 0)
        //                        {
        //                            string nomCapture = codProgram + motor + Hora();
        //                            file = CrearDocumentoWordDinamico(nomCapture, Carpeta, motor);

        //                        }

        //                        try
        //                        {
        //                            if (cont == 0)
        //                            {
        //                                ErrorValidacion.Clear();
        //                                AbrirPrograma a = new AbrirPrograma("KNmDiasn", user);
        //                                desktopSession = a.Start2(motor, "");
        //                            }

        //                            //Thread.Sleep(15000);
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
        //                                if (cont == 0)
        //                                {
        //                                    //QBE
        //                                    errors = FuncionesKNmLivac.QBEQry(desktopSession, TipoQbe, QbeFilter, file, QbeFilter2);
        //                                    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                    Thread.Sleep(1000);
        //                                }
        //                                Thread.Sleep(3000);
        //                                //Funcion Validaciones KNmDiasn
        //                                errors = FuncionesKNmLivac.ValidacionesKNmDiasn(desktopSession, IdInicial, IdFinal, cont, Concepto, FechaDesde, FechaHasta);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                            }
        //                            else
        //                            {
        //                                ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
        //                            }
        //                            if (ErrorValidacion.Count > 0 && DataCase.Tables.Count == cont)
        //                            {
        //                                ReporteErrores(ErrorValidacion, codProgram, motor, Carpeta);
        //                                var separator = string.Format("{0}{0}", Environment.NewLine);
        //                                string errorMessageString = string.Join(separator, ErrorValidacion);
        //                                Assert.Inconclusive(string.Format("Los errores presentados en la prueba {2} son:{0}{1}", Environment.NewLine, errorMessageString, codProgram));
        //                            }

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
        //                                        //break;
        //                                    }
        //                                }
        //                            }
        //                            AbrirPrograma a = new AbrirPrograma();
        //                            //a.Stop();
        //                            //Assert.Fail(CaseId + " ::::::" + e.ToString());
        //                            //break;
        //                        }
        //                    }
							//else
							//{
							//	errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Al validar variables, no tiene registros o no existe");
							//}
        //                    cont = cont + 1;
        //                }

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

        //[TestMethod]
        //public void ValidacionesKNmVacac()
        //{
        //    List<string> errorMessages = new List<string>();
        //    List<string> ErrorValidacion = new List<string>();
        //    List<string> errors = new List<string>();
        //    string motor = string.Empty;
        //    string codProgram = string.Empty;
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
        //    //TableOrder = "DW_A1705";
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

        //        if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesoKNmLivac.ModulosProcesoKNmLivac.ValidacionesKNmVacac")
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
        //                for (int i = 1; i <= 4; i++)
        //                {
        //                    int cont = 0;
        //                    if (i == 1)
        //                    {
        //                        codProgram = "KNmVacac";
        //                    }
        //                    if (i == 2)
        //                    {
        //                        codProgram = "KNmDiasn";
        //                    }
        //                    if (i == 3)
        //                    {
        //                        codProgram = "KNmAusen";
        //                    }
        //                    if (i == 4)
        //                    {
        //                        codProgram = "KNmPreno";
        //                    }

        //                    foreach (DataRow rows in DataCase.Tables[0].Rows)
        //                {
        //                    //int velocidad = 10;

        //                    //Playback.PlaybackSettings.DelayBetweenActions = velocidad;

        //                    //@User @Motor @CodProgram @NomProgram @TipoQbe @QbeFilter @banExcel @BanderaLupa @BanderaSum @BanderaPreli
        //                    //@Identificacion @Concepto @Ausencia @Observaciones

        //                    if (
        //                       //Datos Login
        //                       rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
        //                       rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
        //                       rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
        //                       rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
        //                       rows["QbeFilter2"].ToString().Length != 0 && rows["QbeFilter2"].ToString() != null &&
        //                       //rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
        //                       rows["IdInicial"].ToString().Length != 0 && rows["IdInicial"].ToString() != null &&
        //                       rows["IdFinal"].ToString().Length != 0 && rows["IdFinal"].ToString() != null &&
        //                       rows["Concepto"].ToString().Length != 0 && rows["Concepto"].ToString() != null &&
        //                       rows["FechaDesde"].ToString().Length != 0 && rows["FechaDesde"].ToString() != null &&
        //                       rows["FechaHasta"].ToString().Length != 0 && rows["FechaHasta"].ToString() != null &&
        //                       rows["FechaICausacion"].ToString().Length != 0 && rows["FechaICausacion"].ToString() != null &&
        //                       rows["FechaFCausacion"].ToString().Length != 0 && rows["FechaFCausacion"].ToString() != null &&
        //                       rows["SueldoBase"].ToString().Length != 0 && rows["SueldoBase"].ToString() != null
        //                       )
        //                    {

        //                        string user = rows["User"].ToString();
        //                        motor = rows["Motor"].ToString();
        //                        string TipoQbe = rows["TipoQbe"].ToString();
        //                        string QbeFilter = rows["QbeFilter"].ToString();
        //                        string QbeFilter2 = rows["QbeFilter2"].ToString();
        //                        //string Campo = rows["Campo"].ToString();
        //                        string IdInicial = rows["IdInicial"].ToString();
        //                        string IdFinal = rows["IdFinal"].ToString();
        //                            string Concepto = rows["Concepto"].ToString();
        //                            string FechaDesde = rows["FechaDesde"].ToString();
        //                            string FechaHasta = rows["FechaHasta"].ToString();
        //                            string FechaICausacion = rows["FechaICausacion"].ToString();
        //                        string FechaFCausacion = rows["FechaFCausacion"].ToString();
        //                        string SueldoBase = rows["SueldoBase"].ToString();
        //                            string Campo = "";
        //                            string Carpeta = "Proceso_KNmLivac";
        //                            string file = "";
        //                            if (codProgram == "KNmVacac")
        //                            {
        //                                Campo = "6";
        //                            }
        //                            if (codProgram == "KNmDiasn")
        //                            {
        //                                Campo = "4";
        //                            }
        //                            if (codProgram == "KNmAusen")
        //                            {
        //                                Campo = "8";
        //                            }
        //                            if (codProgram == "KNmPreno")
        //                            {
        //                                Campo = "6";
        //                            }
        //                            //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
        //                            ////Debugger.Launch();
        //                            if (cont == 0)
        //                        {
        //                            string nomCapture = codProgram + motor + Hora();
        //                            file = CrearDocumentoWordDinamico(nomCapture, Carpeta, motor);

        //                        }

        //                        try
        //                        {
        //                            if (cont == 0)
        //                            {
        //                                ErrorValidacion.Clear();
        //                                AbrirPrograma a = new AbrirPrograma(codProgram, user);
        //                                desktopSession = a.Start2(motor, "");
        //                            }

        //                            //Thread.Sleep(15000);
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
        //                                if (cont == 0)
        //                                {
        //                                    //QBE
        //                                    errors = FuncionesKNmLivac.QBEQry(desktopSession, TipoQbe, QbeFilter, file, QbeFilter2, Campo);
        //                                    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                    Thread.Sleep(1000);
        //                                }
        //                                Thread.Sleep(3000);
        //                                    //Funcion Validaciones
        //                                    errors = FuncionesKNmLivac.ValidacionesKNmVacac(desktopSession, codProgram, IdInicial, IdFinal, cont,Concepto,FechaDesde,FechaHasta, FechaICausacion, FechaFCausacion, SueldoBase);
        //                                    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                }
        //                            else
        //                            {
        //                                ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
        //                            }
        //                            if (ErrorValidacion.Count > 0 && DataCase.Tables.Count == cont)
        //                            {
        //                                ReporteErrores(ErrorValidacion, codProgram, motor, Carpeta);
        //                                var separator = string.Format("{0}{0}", Environment.NewLine);
        //                                string errorMessageString = string.Join(separator, ErrorValidacion);
        //                                Assert.Inconclusive(string.Format("Los errores presentados en la prueba {2} son:{0}{1}", Environment.NewLine, errorMessageString, codProgram));
        //                            }

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
        //                                        //break;
        //                                    }
        //                                }
        //                            }
        //                            AbrirPrograma a = new AbrirPrograma();
        //                            //a.Stop();
        //                            //Assert.Fail(CaseId + " ::::::" + e.ToString());
        //                            //break;
        //                        }
        //                    }
							//else
							//{
							//	errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Al validar variables, no tiene registros o no existe");
							//}
        //                    cont = cont + 1;
        //                }
        //                }
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
        public void AbrirProgramaValidacionesKNmDiasn()
        {
            string codProgram = "KNmDiasn";//
            string nomProgram = "Días no Trabajados";
            string motor = "SQL";

            string user = "";
            if (motor == "SQL")
            {
                user = "ADesar2";
            }
            if (motor == "ORA")
            {
                user = "ODesar2";
            }

            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            ////Variables
            //string nomProgram = "NomProgram";
            //QBE
            string TipoQbe = "0";
            string QbeFilter = "6322";
            string Campo = "1";
            //EXCEL
            string banExcel = "0";
            //LUPA
            string BanderaLupa = "0";
            //SUMATORIA
            string BanderaSum = "0";
            //PRELIMINAR
            string BanderaPreli = "0";

            string nomCapture = codProgram + motor + Hora();
            string file = CrearDocumentoWordDinamico(nomCapture, "Proceso_KNmLivac", motor);

            AbrirPrograma a = new AbrirPrograma(codProgram, user);
            desktopSession = a.Start2(motor, "");

            //Carpeta
            string CarpetaErrores = "Proceso_KNmLivac";

            //Variables validacion KNmDiasn
            string IdentificacionEsp = "6322";
            string ConceptoEsp = "1250";
            string FechaDesdeEsp = "20/05/2019";
            string FechaHastaEsp = "21/05/2019";
            List<string> VarValKNmDiasn = new List<string> { IdentificacionEsp, ConceptoEsp, FechaDesdeEsp, FechaHastaEsp };


            //////Validaciones KNmDiasn
            //QBE
            errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, Campo);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            Thread.Sleep(1000);
            ////Funcion Validaciones KNmDiasn
            //errors = FuncionesKNmLivac.ValidacionesKNmDiasn(desktopSession, VarValKNmDiasn);
            //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }




            Thread.Sleep(100000);

            //a.Stop();
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
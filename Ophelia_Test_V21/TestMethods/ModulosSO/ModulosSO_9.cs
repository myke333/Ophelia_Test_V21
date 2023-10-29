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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo_2;
using OpenQA.Selenium.Interactions;

namespace Ophelia_Test_V21.TestMethods.ModulosSO
{
    /// <summary>
    /// Descripción resumida de ModulosSO_9
    /// </summary>
    [TestClass]
    public class ModulosSO_9 : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();
        public ModulosSO_9()
        {
            //
            // TODO: Agregar aquí la lógica del constructor
            //
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
        public void KSoRevotCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_9.KSoRevotCheckList")
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
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null
                                )
                            {
                                //VARIABLES GENERALES
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
                                string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);

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
                                        Screenshot("Abrir Programa", true, file, desktopSession);
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

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

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
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, QbeFilter, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));
                                        newFunctions_1.lapiz(desktopSession);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

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
							else
							{
								errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Al validar variables, no tiene registros o no existe");
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
        public void KSoIndicCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_9.KSoIndicCheckList")
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
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null
                                )
                            {
                                //VARIABLES GENERALES
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
                                string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);

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
                                        Screenshot("Abrir Programa", true, file, desktopSession);
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

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

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
                                        CrudKsoindic.PreliKSoIndic(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //RECTIFICACION DE CAMPOS DE EXCEL
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, QbeFilter, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));
                                        newFunctions_1.lapiz(desktopSession);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

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
							else
							{
								errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Al validar variables, no tiene registros o no existe");
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
        public void KSoPgramCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_9.KSoPgramCheckList")
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
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null
                                )
                            {
                                //VARIABLES GENERALES
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
                                string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);

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
                                        Screenshot("Abrir Programa", true, file, desktopSession);
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

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

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
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, QbeFilter, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));
                                        newFunctions_1.lapiz(desktopSession);
                                        Thread.Sleep(1000);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

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
							else
							{
								errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Al validar variables, no tiene registros o no existe");
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
        public void KSoRevacCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_9.KSoRevacCheckList")
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
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["Id"].ToString().Length != 0 && rows["Id"].ToString() != null &&
                                rows["Vacuna"].ToString().Length != 0 && rows["Vacuna"].ToString() != null &&
                                rows["Resp"].ToString().Length != 0 && rows["Resp"].ToString() != null &&
                                rows["EditResp"].ToString().Length != 0 && rows["EditResp"].ToString() != null &&



                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null
                                )
                            {
                                //VARIABLES GENERALES
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
                                string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();

                                string  Id = rows["Id"].ToString();
                                string  Vacuna = rows["Vacuna"].ToString();
                                string  Resp= rows["Resp"].ToString();
                                string   EditResp = rows["EditResp"].ToString();

                                List<string> crudPrincipal = new List<string>() { Id, Vacuna, Resp, EditResp };

                                string campo = "1";
                                


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);

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
                                        Screenshot("Abrir Programa", true, file, desktopSession);
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


                                        ///Agregar Datos

                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKsorevac.AgregarDatos(desktopSession, 0, crudPrincipal, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        ///Descartar edicion 

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);
                                        Thread.Sleep(1500);
                                        CrudKsorevac.AgregarDatos(desktopSession, 1, crudPrincipal, file);
                                        newFunctions_4.ScreenshotNuevo("Editar Datos Crud Principal", true, file);
                                        Thread.Sleep(1500);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ///Aceptar edicion 
                                        ///

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);
                                        CrudKsorevac.AgregarDatos(desktopSession, 1, crudPrincipal, file);
                                        newFunctions_4.ScreenshotNuevo("Editar Datos Crud Principal", true, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(3000);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);






                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

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
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, QbeFilter, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));
                                        newFunctions_1.lapiz(desktopSession);
                                        Thread.Sleep(1000);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ///

                                        ///ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
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
							else
							{
								errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Al validar variables, no tiene registros o no existe");
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
        public void KSoPanEvCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_9.KSoPanEvCheckList")
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
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&

                                rows["EvaRegSecu"].ToString().Length != 0 && rows["EvaRegSecu"].ToString() != null &&
                                rows["EvaRegFecha"].ToString().Length != 0 && rows["EvaRegFecha"].ToString() != null &&
                                rows["EvaRegDescri"].ToString().Length != 0 && rows["EvaRegDescri"].ToString() != null &&
                                rows["EvaRegSecuEditar"].ToString().Length != 0 && rows["EvaRegSecuEditar"].ToString() != null &&

                                rows["FechaSeg"].ToString().Length != 0 && rows["FechaSeg"].ToString() != null &&
                                rows["CronRespon"].ToString().Length != 0 && rows["CronRespon"].ToString() != null &&
                                rows["FechaSegEditar"].ToString().Length != 0 && rows["FechaSegEditar"].ToString() != null 
                                )
                            {
                                //VARIABLES GENERALES
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
                                string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();

                                //SQL
                                //VARIABLES CRUD
                                string Centrotrab = "";
                                string Area = "";
                                string AreaEditar = "";

                                //ORA
                                //VARIABLES CRUD
                                string Seccional = "";
                                string SeccionalEditar = "";

                                List<string> crudVariables = new List<string>() { };
                                List<string> crudVariablesdet1 = new List<string>() { };
                                List<string> crudVariablesdet2 = new List<string>() { };
                                List<string> crudVariablesdet3 = new List<string>() { };
                                List<string> crudVariablesdet4 = new List<string>() { };
                                List<string> crudVariablesdet5 = new List<string>() { };
                                List<string> crudVariablesdet6 = new List<string>() { };
                                List<string> crudVariablesdet7 = new List<string>() { };
                                List<string> crudVariablesdet8 = new List<string>() { };
                                List<string> crudVariablesdet9 = new List<string>() { };
                                List<string> crudVariablesdet10 = new List<string>() { };

                                if (motor=="SQL")
                                {
                                    //SQL
                                    //VARIABLES CRUD
                                    Centrotrab = rows["Centrotrab"].ToString();
                                    Area = rows["Area"].ToString();
                                    string Tarea = rows["Tarea"].ToString();
                                    AreaEditar = rows["AreaEditar"].ToString();
                                    //VARIABLES DETALLE 1
                                    string IdPeCodigo = rows["IdPeCodigo"].ToString();
                                    string IdPeCodigoEditar = rows["IdPeCodigoEditar"].ToString();
                                    //VARIABLES DETALLE 2
                                    string FuenCodigo = rows["FuenCodigo"].ToString();
                                    //VARIABLES DETALLE 3
                                    string MediCodigo = rows["MediCodigo"].ToString();
                                    //VARIABLES DETALLE 4
                                    string IndiCodigo = rows["IndiCodigo"].ToString();
                                    string IndiCodigoEditar = rows["IndiCodigoEditar"].ToString();
                                    //VARIABLES DETALLE 5
                                    string EvaRegSecu = rows["EvaRegSecu"].ToString();
                                    string EvaRegFecha = rows["EvaRegFecha"].ToString();
                                    string EvaRegDescri = rows["EvaRegDescri"].ToString();
                                    string EvaRegSecuEditar = rows["EvaRegSecuEditar"].ToString();
                                    //VARIABLES DETALLE 6
                                    string CriContCod = rows["CriContCod"].ToString();
                                    string CriContCodEditar = rows["CriContCodEditar"].ToString();
                                    //VARIABLES DETALLE 7
                                    string EfeProCodigo = rows["EfeProCodigo"].ToString();
                                    string EfeProNombreEditar = rows["EfeProNombreEditar"].ToString();
                                    //VARIABLES DETALLE 8
                                    string RecTecFuen = rows["RecTecFuen"].ToString();
                                    string RecTecFuenEditar = rows["RecTecFuenEditar"].ToString();
                                    //VARIABLES DETALLE 9
                                    string PlanAConju = rows["PlanAConju"].ToString();
                                    string TipoMContr = rows["TipoMContr"].ToString();
                                    string PlanAAct = rows["PlanAAct"].ToString();
                                    string PlanAConjuEditar = rows["PlanAConjuEditar"].ToString();
                                    //VARIABLES DETALLE 10
                                    string FechaSeg = rows["FechaSeg"].ToString();
                                    string CronRespon = rows["CronRespon"].ToString();
                                    string FechaSegEditar = rows["FechaSegEditar"].ToString();

                                    crudVariables = new List<string>() { Centrotrab, Area, Tarea, AreaEditar };
                                    crudVariablesdet1 = new List<string>() { IdPeCodigo, IdPeCodigoEditar };
                                    crudVariablesdet2 = new List<string>() { FuenCodigo };
                                    crudVariablesdet3 = new List<string>() { MediCodigo };
                                    crudVariablesdet4 = new List<string>() { IndiCodigo, IndiCodigoEditar };
                                    crudVariablesdet5 = new List<string>() { EvaRegSecu, EvaRegFecha, EvaRegDescri, EvaRegSecuEditar };
                                    crudVariablesdet6 = new List<string>() { CriContCod, CriContCodEditar };
                                    crudVariablesdet7 = new List<string>() { EfeProCodigo, EfeProNombreEditar };
                                    crudVariablesdet8 = new List<string>() { RecTecFuen, RecTecFuenEditar };
                                    crudVariablesdet9 = new List<string>() { PlanAConju, TipoMContr, PlanAAct, PlanAConjuEditar };
                                    crudVariablesdet10 = new List<string>() { FechaSeg, CronRespon, FechaSegEditar };

                                }

                                if (motor=="ORA")
                                {
                                    //ORA
                                    //VARIABLES CRUD
                                    Seccional = rows["Seccional"].ToString();
                                    string Dependencia = rows["Dependencia"].ToString();
                                    string CentroCost = rows["CentroCost"].ToString();
                                    string Zona = rows["Zona"].ToString();
                                    string Pais = rows["Pais"].ToString();
                                    SeccionalEditar = rows["SeccionalEditar"].ToString();
                                    //VARIABLES DETALLE 1 No tiene por falta de parametrizacion
                                    //VARIABLES DETALLE 2 No tiene por falta de parametrizacion
                                    //VARIABLES DETALLE 3 No tiene por falta de parametrizacion
                                    //VARIABLES DETALLE 4 No tiene por falta de parametrizacion
                                    //VARIABLES DETALLE 5
                                    string EvaRegSecu = rows["EvaRegSecu"].ToString();
                                    string EvaRegFecha = rows["EvaRegFecha"].ToString();
                                    string EvaRegDescri = rows["EvaRegDescri"].ToString();
                                    string EvaRegSecuEditar = rows["EvaRegSecuEditar"].ToString();
                                    //VARIABLES DETALLE 6 No tiene por falta de parametrizacion
                                    //VARIABLES DETALLE 7 No tiene
                                    //VARIABLES DETALLE 8 No tiene
                                    //VARIABLES DETALLE 10
                                    string FechaSeg = rows["FechaSeg"].ToString();
                                    string CronRespon = rows["CronRespon"].ToString();
                                    string FechaSegEditar = rows["FechaSegEditar"].ToString();

                                    crudVariables = new List<string>() { Seccional, Dependencia, CentroCost, Zona, Pais, SeccionalEditar };
                                    crudVariablesdet5 = new List<string>() { EvaRegSecu, EvaRegFecha, EvaRegDescri, EvaRegSecuEditar };
                                    crudVariablesdet10 = new List<string>() { FechaSeg, CronRespon, FechaSegEditar };
                                }


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);

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
                                        Screenshot("Abrir Programa", true, file, desktopSession);
                                        /*//VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }*/

                                        //ELIMINAR REGISTRO PEGADO
                                        if (motor == "SQL")
                                        {
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Centrotrab, Centrotrab, "COD_CENT", c, ErrorValidacion, "No se agregó correctamente", 0, "1");
                                        }
                                        else
                                        {
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Seccional, Seccional, "COD_NIVE", c, ErrorValidacion, "No se agregó correctamente", 0,"1");
                                        }

                                        /*//VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);*/

                                        var ElementList9 = PruebaCRUD.NavClass(desktopSession);
                                        var ElementList8 = PruebaCRUD.NavClass(desktopSession);
                                        var ElementList7 = PruebaCRUD.NavClass(desktopSession);
                                        var ElementList6 = PruebaCRUD.NavClass(desktopSession);
                                        var ElementList1 = PruebaCRUD.NavClass(desktopSession);

                                        //////////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        CrudKsopanev.CRUDKSoPanEv(desktopSession, 1, crudVariables, file,motor);
                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        //Debugger.Launch();
                                        WindowsDriver<WindowsElement> rootSession = null;
                                        rootSession = PruebaCRUD.RootSession();
                                        //rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                                        var allFrame = rootSession.FindElementsByClassName("IEFrame");
                                        new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) - 30, (allFrame[0].Size.Height / 2) + 60).DoubleClick().Click().Perform();
                                        Thread.Sleep(5000);
                                        //desktopSession.Mouse.Click(null);

                                        Thread.Sleep(3000);
                                        newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);
                                        Thread.Sleep(1000);
                                        /*if (motor=="SQL")
                                        {
                                            //Validacion Agregar
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Centrotrab, Centrotrab, "COD_CENT", c, ErrorValidacion, "No se agregó correctamente", 0);
                                        }
                                        else
                                        {
                                            //Validacion Agregar
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Seccional, Seccional, "COD_NIVE", c, ErrorValidacion, "No se agregó correctamente", 0);
                                        }*/

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        CrudKsopanev.CRUDKSoPanEv(desktopSession, 2, crudVariables, file, motor);
                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);
                                        /*if (motor == "SQL")
                                        {
                                            //Validacion Editar Descartar
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Centrotrab, Area, "COD_AREA", c, ErrorValidacion, "No se edito correctamente", 1);
                                        }
                                        else
                                        {
                                            //Validacion Editar Descartar
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Seccional, Seccional, "COD_NIVE", c, ErrorValidacion, "No se edito correctamente", 1);
                                        }*/

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        CrudKsopanev.CRUDKSoPanEv(desktopSession, 2, crudVariables, file,motor);
                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        rootSession = PruebaCRUD.RootSession();
                                        rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                                        allFrame = rootSession.FindElementsByClassName("IEFrame");
                                        new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) - 30, (allFrame[0].Size.Height / 2) + 60).DoubleClick().Click().Perform();
                                        desktopSession.Mouse.Click(null);
                                        new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 90, (allFrame[0].Size.Height / 2) + 150).DoubleClick().Click().Perform();
                                        for (int i = 0; i < 8; i++)
                                        {
                                            desktopSession.Mouse.Click(null);
                                        }
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);
                                        /*if (motor == "SQL")
                                        {
                                            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Centrotrab, "E", c, ErrorValidacion);

                                            //Validacion Editar Aceptar
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Centrotrab, AreaEditar, "COD_AREA", c, ErrorValidacion, "El registro no se editó correctamente", 1);
                                        }
                                        else
                                        {
                                            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, SeccionalEditar, "E", c, ErrorValidacion);

                                            //Validacion Editar Aceptar
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, SeccionalEditar, SeccionalEditar, "COD_NIVE", c, ErrorValidacion, "El registro no se editó correctamente", 1);
                                        }*/

                                        //PASAR Pestaña 1
                                        var TPageControl = desktopSession.FindElementsByClassName("TPageControl");
                                        desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 130, 13);
                                        desktopSession.Mouse.Click(null);

                                        if (motor=="SQL")
                                        {
                                        //PASAR DETALLE 1
                                        TPageControl = desktopSession.FindElementsByClassName("TPageControl");
                                        desktopSession.Mouse.MouseMove(TPageControl[1].Coordinates, 50, 13);
                                        desktopSession.Mouse.Click(null);

                                        ////////AGREGAR DETALLE 1
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopanev.CRUDDetalleKSoPanEv(desktopSession, 1, crudVariablesdet1, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 1", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopanev.CRUDDetalleKSoPanEv(desktopSession, 2, crudVariablesdet1, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 1", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopanev.CRUDDetalleKSoPanEv(desktopSession, 2, crudVariablesdet1, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 1", true, file);

                                        //PASAR DETALLE 2
                                        CrudKsopanev.PasarPestaña(desktopSession, 1, 1, 1);
                                        TPageControl = desktopSession.FindElementsByClassName("TPageControl");
                                        desktopSession.Mouse.MouseMove(TPageControl[2].Coordinates, 30, 13);
                                        desktopSession.Mouse.Click(null);

                                        ////////AGREGAR DETALLE 2
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                        ElementList1 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopanev.CRUDDetalleKSoPanEv(desktopSession, 1, crudVariablesdet2, file);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 2", true, file);

                                        //PASAR DETALLE 3
                                        CrudKsopanev.PasarPestaña(desktopSession, 1, 1, 2);

                                        ////////AGREGAR DETALLE 3
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                        ElementList1 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopanev.CRUDDetalleKSoPanEv(desktopSession, 1, crudVariablesdet3, file);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 3", true, file);

                                        //PASAR DETALLE 4
                                        CrudKsopanev.PasarPestaña(desktopSession, 2, 1, 2);

                                        ////////AGREGAR DETALLE 4
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                        ElementList1 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopanev.CRUDDetalleKSoPanEv(desktopSession, 1, crudVariablesdet4, file);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 4", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopanev.CRUDDetalleKSoPanEv(desktopSession, 2, crudVariablesdet4, file);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 4", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopanev.CRUDDetalleKSoPanEv(desktopSession, 2, crudVariablesdet4, file);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 4", true, file);
                                        }
                                        //PASAR DETALLE 5
                                        CrudKsopanev.PasarPestaña(desktopSession, 2, 1, 1);

                                        ////////AGREGAR DETALLE 5
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 1);
                                        var ElementList5 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopanev.CRUDDetalle5KSoPanEv(desktopSession, 1, crudVariablesdet5, file);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);                                        
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 5", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopanev.CRUDDetalle5KSoPanEv(desktopSession, 2, crudVariablesdet5, file);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 5", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopanev.CRUDDetalle5KSoPanEv(desktopSession, 2, crudVariablesdet5, file);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000); 
                                        if (motor == "ORA")
                                        {
                                            bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                            if (valEditOra == true)
                                            {
                                                //LUPA
                                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                                //ACEPTAR EDICION
                                                desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                CrudKsopanev.CRUDDetalle5KSoPanEv(desktopSession, 2, crudVariablesdet5, file);
                                                desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                Thread.Sleep(1000);
                                            }
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 5", true, file);

                                        if (motor == "SQL")
                                        {
                                            //PASAR DETALLE 6
                                            CrudKsopanev.PasarPestaña(desktopSession, 3, 1, 1);

                                            ////////AGREGAR DETALLE 6
                                            botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 1);
                                            ElementList6 = PruebaCRUD.NavClass(desktopSession);
                                            desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsopanev.CRUDDetalleKSoPanEv(desktopSession, 1, crudVariablesdet6, file);
                                            desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(200);
                                            newFunctions_4.ScreenshotNuevo("Agregar Detalle 5", true, file);

                                            //DESCARTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsopanev.CRUDDetalleKSoPanEv(desktopSession, 2, crudVariablesdet6, file);
                                            desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                            newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 5", true, file);

                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsopanev.CRUDDetalleKSoPanEv(desktopSession, 2, crudVariablesdet6, file);
                                            desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(200);
                                            bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                            if (valEditOra == true)
                                            {
                                                //LUPA
                                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                                //ACEPTAR EDICION
                                                desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                CrudKsopanev.CRUDDetalleKSoPanEv(desktopSession, 2, crudVariablesdet6, file);
                                                desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                Thread.Sleep(200);
                                            }
                                            newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 5", true, file);

                                            //PASAR DETALLE 7
                                            CrudKsopanev.PasarPestaña(desktopSession, 2, 1, 0);
                                            CrudKsopanev.PasarPestaña(desktopSession, 3, 1, 1);

                                            ////////AGREGAR DETALLE 7
                                            botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                            ElementList7 = PruebaCRUD.NavClass(desktopSession);
                                            desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsopanev.CRUDDetalle7KSoPanEv(desktopSession, 1, crudVariablesdet7, file);
                                            desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(200);
                                            newFunctions_4.ScreenshotNuevo("Agregar Detalle 7", true, file);

                                            //DESCARTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsopanev.CRUDDetalle7KSoPanEv(desktopSession, 2, crudVariablesdet7, file);
                                            desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                            newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 7", true, file);

                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsopanev.CRUDDetalle7KSoPanEv(desktopSession, 2, crudVariablesdet7, file);
                                            desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                            newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 7", true, file);

                                            //PASAR DETALLE 8
                                            CrudKsopanev.PasarPestaña(desktopSession, 5, 1, 1);

                                            ////////AGREGAR DETALLE 8
                                            botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                            ElementList8 = PruebaCRUD.NavClass(desktopSession);
                                            desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsopanev.CRUDDetalleKSoPanEv(desktopSession, 1, crudVariablesdet8, file);
                                            desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(200);
                                            valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                            if (valEditOra == true)
                                            {
                                                //LUPA
                                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                                //AGREGAR DETALLE 8
                                                desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                CrudKsopanev.CRUDDetalleKSoPanEv(desktopSession, 1, crudVariablesdet8, file);
                                                desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                Thread.Sleep(200);
                                            }
                                            newFunctions_4.ScreenshotNuevo("Agregar Detalle 8", true, file);

                                            //DESCARTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsopanev.CRUDDetalleKSoPanEv(desktopSession, 2, crudVariablesdet8, file);
                                            desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                            newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 8", true, file);

                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsopanev.CRUDDetalleKSoPanEv(desktopSession, 2, crudVariablesdet8, file);
                                            desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                            valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                            if (valEditOra == true)
                                            {
                                                //LUPA
                                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                                //ACEPTAR EDICION
                                                desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                CrudKsopanev.CRUDDetalleKSoPanEv(desktopSession, 2, crudVariablesdet8, file);
                                                desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                Thread.Sleep(1000);
                                            }
                                            newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 8", true, file);

                                            //PASAR DETALLE 9
                                            CrudKsopanev.PasarPestaña(desktopSession, 6, 1, 1);

                                            ////////AGREGAR DETALLE 9
                                            botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                            ElementList9 = PruebaCRUD.NavClass(desktopSession);
                                            desktopSession.Mouse.MouseMove(ElementList9[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsopanev.CRUDDetalle9KSoPanEv(desktopSession, 1, crudVariablesdet9, file);
                                            desktopSession.Mouse.MouseMove(ElementList9[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(200);
                                            valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                            if (valEditOra == true)
                                            {
                                                //LUPA
                                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                                //AGREGAR DETALLE 9
                                                desktopSession.Mouse.MouseMove(ElementList9[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                CrudKsopanev.CRUDDetalle9KSoPanEv(desktopSession, 1, crudVariablesdet9, file);
                                                desktopSession.Mouse.MouseMove(ElementList9[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                Thread.Sleep(200);
                                            }
                                            newFunctions_4.ScreenshotNuevo("Agregar Detalle 9", true, file);

                                            //DESCARTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList9[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsopanev.CRUDDetalle9KSoPanEv(desktopSession, 2, crudVariablesdet9, file);
                                            desktopSession.Mouse.MouseMove(ElementList9[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                            newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 9", true, file);

                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList9[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsopanev.CRUDDetalle9KSoPanEv(desktopSession, 2, crudVariablesdet9, file);
                                            desktopSession.Mouse.MouseMove(ElementList9[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                            newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 9", true, file);
                                        }

                                        //PASAR DETALLE 10
                                        CrudKsopanev.PasarPestaña(desktopSession, 3, 1, 0);

                                        ////////AGREGAR DETALLE 10
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                        var ElementList10 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList10[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopanev.CRUDDetalle10KSoPanEv(desktopSession, 1, crudVariablesdet10, file);
                                        desktopSession.Mouse.MouseMove(ElementList10[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 9", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList10[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopanev.CRUDDetalle10KSoPanEv(desktopSession, 2, crudVariablesdet10, file);
                                        desktopSession.Mouse.MouseMove(ElementList10[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 9", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList10[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopanev.CRUDDetalle10KSoPanEv(desktopSession, 2, crudVariablesdet10, file);
                                        desktopSession.Mouse.MouseMove(ElementList10[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000); 
                                        if (motor == "ORA")
                                        {
                                            bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                            if (valEditOra == true)
                                            {
                                                //LUPA
                                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                                //ACEPTAR EDICION
                                                desktopSession.Mouse.MouseMove(ElementList10[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                CrudKsopanev.CRUDDetalle10KSoPanEv(desktopSession, 2, crudVariablesdet10, file);
                                                desktopSession.Mouse.MouseMove(ElementList10[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                Thread.Sleep(1000);
                                            }
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 9", true, file);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

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
                                        CrudKsopanev.PreliKSoPanEv(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli,motor);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        if (motor == "SQL")
                                        {
                                            //RECTIFICACION DE CAMPOS DE EXCEL
                                            string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, Centrotrab, c);
                                            string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                            Celda = Celda_Excel(ruta, nom_empr);
                                            Celda.ForEach(u => ErrorValidacion.Add(u));
                                        }
                                        else
                                        {
                                            //RECTIFICACION DE CAMPOS DE EXCEL
                                            string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, SeccionalEditar, c);
                                            string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                            Celda = Celda_Excel(ruta, nom_empr);
                                            Celda.ForEach(u => ErrorValidacion.Add(u));
                                        }
                                        newFunctions_1.lapiz(desktopSession);
                                        Thread.Sleep(1000);
                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR DETALLE 10
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList10[1], botonesNavegador, file);
                                        Thread.Sleep(500);

                                        if (motor == "SQL")
                                        {
                                            //PASAR DETALLE 9
                                            CrudKsopanev.PasarPestaña(desktopSession, 2, 1, 0);
                                            //ELIMINAR DETALLE 9
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList9[1], botonesNavegador, file);
                                            Thread.Sleep(500);

                                            //PASAR DETALLE 8
                                            CrudKsopanev.PasarPestaña(desktopSession, 5, 1, 1);
                                            Thread.Sleep(500);
                                            //ELIMINAR DETALLE 8
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList8[1], botonesNavegador, file);
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList8[1], botonesNavegador, file);
                                            Thread.Sleep(500);

                                            //PASAR DETALLE 7
                                            CrudKsopanev.PasarPestaña(desktopSession, 3, 1, 1);
                                            Thread.Sleep(500);
                                            //ELIMINAR DETALLE 7
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList7[1], botonesNavegador, file);
                                            Thread.Sleep(500);

                                            //PASAR DETALLE 6
                                            CrudKsopanev.PasarPestaña(desktopSession, 1, 1, 0);
                                            Thread.Sleep(500);
                                            //ELIMINAR DETALLE 6
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList6[1], botonesNavegador, file);
                                            rootSession = PruebaCRUD.RootSession();
                                            rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                                            allFrame = rootSession.FindElementsByClassName("IEFrame");
                                            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) - 30, (allFrame[0].Size.Height / 2) + 60).DoubleClick().Click().Perform();
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(500);

                                            //PASAR DETALLE 5
                                            CrudKsopanev.PasarPestaña(desktopSession, 2, 1, 1);
                                            Thread.Sleep(500);
                                        }
                                        else
                                        {
                                            CrudKsopanev.PasarPestaña(desktopSession, 1, 1, 0);
                                        }
                                        //ELIMINAR DETALLE 5
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList5[1], botonesNavegador, file);
                                        Thread.Sleep(500);

                                        if (motor == "SQL")
                                        {
                                            //PASAR DETALLE 4
                                            CrudKsopanev.PasarPestaña(desktopSession, 1, 1, 1);
                                            CrudKsopanev.PasarPestaña(desktopSession, 2, 1, 2);
                                            Thread.Sleep(500);
                                            //ELIMINAR DETALLE 4
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[1], botonesNavegador, file);
                                            Thread.Sleep(500);

                                            //PASAR DETALLE 3
                                            CrudKsopanev.PasarPestaña(desktopSession, 1, 1, 2);
                                            Thread.Sleep(500);
                                            //ELIMINAR DETALLE 3
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[1], botonesNavegador, file);
                                            Thread.Sleep(500);

                                            //PASAR DETALLE 2
                                            TPageControl = desktopSession.FindElementsByClassName("TPageControl");
                                            desktopSession.Mouse.MouseMove(TPageControl[2].Coordinates, 30, 13);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(500);
                                            //ELIMINAR DETALLE 2
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[1], botonesNavegador, file);
                                            Thread.Sleep(500);

                                            //PASAR DETALLE 1
                                            TPageControl = desktopSession.FindElementsByClassName("TPageControl");
                                            desktopSession.Mouse.MouseMove(TPageControl[1].Coordinates, 50, 13);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(500);
                                            //ELIMINAR DETALLE 1
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador, file);
                                            Thread.Sleep(500);
                                        }

                                        //ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        rootSession = PruebaCRUD.RootSession();
                                        rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                                        allFrame = rootSession.FindElementsByClassName("IEFrame");
                                        new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) - 30, (allFrame[0].Size.Height / 2) + 80).DoubleClick().Click().Perform();
                                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);

                                        //Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Centrotrab, "", "COD_CENT", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Centrotrab, "D", c, ErrorValidacion);

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
							else
							{
								errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Al validar variables, no tiene registros o no existe");
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
}

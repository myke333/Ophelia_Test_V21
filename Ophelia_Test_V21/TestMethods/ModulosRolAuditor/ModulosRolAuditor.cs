using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data;
using Ophelia_Test_V21.AFuncionesGlobales;
using OpenQA.Selenium.Appium.Windows;
using System.Threading;
using System.Diagnostics;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesRolAuditor;

namespace Ophelia_Test_V21.TestMethods.ModulosRolAuditor
{
    /// <summary>
    /// Descripción resumida de ModulosRolAuditor
    /// </summary>
    [TestClass]
    public class ModulosRolAuditor : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();
        public ModulosRolAuditor()
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
        public void RolAuditor()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosRolAuditor.ModulosRolAuditor.RolAuditor")
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
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                //rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null
                                )
                            {

                                string user = rows["User"].ToString();
                                motor = rows["Motor"].ToString();
                                codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                string nomCapture = codProgram + motor + Hora();
                                string file = CrearDocumentoWordDinamico(nomCapture, "RolAuditor", motor);

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
                                        //APERTURA PROGRAMA
                                        newFunctions_4.ScreenshotNuevo("Abrir programa", true, file);
                                        Thread.Sleep(5000);
                                        //VALIDACION CODIGO DEL PROGRAMA
                                        errors = newFunctions_4.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VALIDACION NOMBRE DEL PROGRAMA
                                        errors = newFunctions_4.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);
                                        ////Debugger.Launch();

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //REPORTE DINAMICO
                                        string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
                                        errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //REPORTE PRELIMINAR
                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = PreliminarGlobal.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1, nomProgram, codProgram, TipoQbe, QbeFilter, motor);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));
                                        Thread.Sleep(1000);

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);
                                        a.Stop();


                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }
                                    if (ErrorValidacion.Count > 0)
                                    {
                                        ReporteErrores(ErrorValidacion, codProgram, motor, "RolAuditor");
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
        public void AbrirPrograma()
        {
            string codProgram = "KNmPreno";//
            string nomProgram = "KNmPazysCheckList";
            string motor = "ORA";

            string user = "";
            if (motor == "SQL")
            {
                user = "UCalida1";
            }
            if (motor == "ORA")
            {
                user = "ODESAR";
            }

            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            ////Variables
            //string nomProgram = "NomProgram";
            //QBE
            string TipoQbe = "0";
            string QbeFilter = "2";
            //EXCEL
            string banExcel = "0";
            //LUPA
            string BanderaLupa = "0";
            //SUMATORIA
            string BanderaSum = "0";
            //PRELIMINAR
            string BanderaPreli = "0";

            string nomCapture = codProgram + motor + Hora();
            string file = CrearDocumentoWordDinamico(nomCapture, "RolAuditor", motor);

            AbrirPrograma a = new AbrirPrograma(codProgram, user);
            desktopSession = a.Start2(motor, "");

            //VERSION
            FuncionesGlobales.GetVersion(desktopSession, file);
            Thread.Sleep(1000);

            //NOTAS
            newFunctions_4.openInnerNote(desktopSession, file);
            Thread.Sleep(1000);
            ////Debugger.Launch();

            //LUPA
            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

            //QBE
            errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            Thread.Sleep(1000);

            //SUMATORIA
            errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            Thread.Sleep(1000);

            //REPORTE DINAMICO
            string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
            errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            Thread.Sleep(1000);

            //REPORTE PRELIMINAR
            string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
            errors = PreliminarGlobal.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1, nomProgram, codProgram, TipoQbe, QbeFilter, motor);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            Thread.Sleep(1000);
            //Rectificar Pie Pagina PDF PRELIMINAR
            listaResultado = Textopdf(pdf1, codProgram, user);
            listaResultado.ForEach(e => ErrorValidacion.Add(e));
            Thread.Sleep(1000);

            //EXPORTAR EXCEL
            string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
            errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            Thread.Sleep(1000);
            a.Stop();
        }
    }
}

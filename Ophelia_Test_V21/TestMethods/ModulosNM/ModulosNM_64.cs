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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosPn;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_7;

namespace Ophelia_Test_V21.TestMethods.ModulosNM
{
    /// <summary>
    /// Descripción resumida de ModulosNM_64
    /// </summary>
    [TestClass]
    public class ModulosNM_64 : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();

        public ModulosNM_64(){}
        [TestMethod]
        public void KNmLaumtCheckList()
        {
            List<string> listaResultado;
            List<string> Celda = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_64.KNmLaumtCheckList")
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
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null
                                
                                )
                            {
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                

                                List<string> crudVariables = new List<string>() { };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);
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
                                        ////VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);


                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }
                                    //CONTADOR DE ERRORES
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
        public void KNmGopsiCheckList()
        {
            List<string> listaResultado;
            List<string> Celda = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_64.KNmGopsiCheckList")
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
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null 
                                
                                )
                            {
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                

                                List<string> crudVariables = new List<string>() { };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);
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
                                        ////VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);


                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }
                                    //CONTADOR DE ERRORES
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
        public void KNmAdmenCheckList()
        {
            //////ARRANCAR EN SQL
            //string motor = "SQL";
            //AbrirPrograma a = new AbrirPrograma("KNmAdmen", "UCalida1");
            //string file = CrearDocumentoWordDinamico("KNmAdmen", "PruebasCrud", motor);
            //List<string> crudVariables = new List<string> { };

            //////ARRANCAR EN ORA
            string motor = "ORA";
            AbrirPrograma a = new AbrirPrograma("KNmAdmen", "ODESAR");
            string file = CrearDocumentoWordDinamico("KNmAdmen", "PruebasCrud", motor);
            List<string> crudVariables = new List<string> { };


            ////EJECUTAR EN EL NAVEGADOR
            desktopSession = a.Start2(motor, "");
            Thread.Sleep(20000000);
            //CrudKpnpcinf.CRUD(desktopSession, "0", crudVariables, file);
            Thread.Sleep(2000);
            //Thread.Sleep(20000000);
        }
        [TestMethod]
        public void KNmMcparCheckList()
        {
            List<string> listaResultado;
            List<string> Celda = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_64.KNmMcparCheckList")
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
                                rows["Tabla1"].ToString().Length != 0 && rows["Tabla1"].ToString() != null &&
                                rows["Campo1"].ToString().Length != 0 && rows["Campo1"].ToString() != null &&
                                rows["EdCampo1"].ToString().Length != 0 && rows["EdCampo1"].ToString() != null &&
                                rows["Tabla2"].ToString().Length != 0 && rows["Tabla2"].ToString() != null &&
                                rows["Campo2"].ToString().Length != 0 && rows["Campo2"].ToString() != null &&
                                rows["EdCampo2"].ToString().Length != 0 && rows["EdCampo2"].ToString() != null &&
                                rows["Tabla3"].ToString().Length != 0 && rows["Tabla3"].ToString() != null &&
                                rows["Campo3"].ToString().Length != 0 && rows["Campo3"].ToString() != null &&
                                rows["EdCampo3"].ToString().Length != 0 && rows["EdCampo3"].ToString() != null &&
                                rows["Tabla4"].ToString().Length != 0 && rows["Tabla4"].ToString() != null &&
                                rows["Campo4"].ToString().Length != 0 && rows["Campo4"].ToString() != null &&
                                rows["EdCampo4"].ToString().Length != 0 && rows["EdCampo4"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&

                                //VARIABLES CRUD 1
                                rows["DirRed"].ToString().Length != 0 && rows["DirRed"].ToString() != null &&
                                rows["DirRedACT"].ToString().Length != 0 && rows["DirRedACT"].ToString() != null &&
                                //VARIABLES CRUD 2
                                rows["Nom"].ToString().Length != 0 && rows["Nom"].ToString() != null &&
                                rows["FecIni"].ToString().Length != 0 && rows["FecIni"].ToString() != null &&
                                rows["FecFin"].ToString().Length != 0 && rows["FecFin"].ToString() != null &&
                                rows["NomACT"].ToString().Length != 0 && rows["NomACT"].ToString() != null &&
                                //VARIABLES CRUD 3
                                rows["Nomb"].ToString().Length != 0 && rows["Nomb"].ToString() != null &&
                                rows["NombACT"].ToString().Length != 0 && rows["NombACT"].ToString() != null &&
                                //VARIABLES CRUD 4
                                rows["Cod"].ToString().Length != 0 && rows["Cod"].ToString() != null &&
                                rows["Nombr"].ToString().Length != 0 && rows["Nombr"].ToString() != null &&
                                rows["TipoDato"].ToString().Length != 0 && rows["TipoDato"].ToString() != null &&
                                rows["Val"].ToString().Length != 0 && rows["Val"].ToString() != null &&
                                rows["NombrACT"].ToString().Length != 0 && rows["NombrACT"].ToString() != null
                                )
                            {
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string Tabla1 = rows["Tabla1"].ToString();
                                string Campo1 = rows["Campo1"].ToString();
                                string EdCampo1 = rows["EdCampo1"].ToString();
                                string Tabla2 = rows["Tabla2"].ToString();
                                string Campo2 = rows["Campo2"].ToString();
                                string EdCampo2 = rows["EdCampo2"].ToString();
                                string Tabla3 = rows["Tabla3"].ToString();
                                string Campo3 = rows["Campo3"].ToString();
                                string EdCampo3 = rows["EdCampo3"].ToString();
                                string Tabla4 = rows["Tabla4"].ToString();
                                string Campo4 = rows["Campo4"].ToString();
                                string EdCampo4 = rows["EdCampo4"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();

                                string DirRed = rows["DirRed"].ToString();
                                string DirRedACT = rows["DirRedACT"].ToString();
                                string Nom = rows["Nom"].ToString();
                                string FecIni = rows["FecIni"].ToString();
                                string FecFin = rows["FecFin"].ToString();
                                string NomACT = rows["NomACT"].ToString();
                                string Nomb = rows["Nomb"].ToString();
                                string NombACT = rows["NombACT"].ToString();
                                string Cod = rows["Cod"].ToString();
                                string Nombr = rows["Nombr"].ToString();
                                string TipDato = rows["TipoDato"].ToString();
                                string Val = rows["Val"].ToString();
                                string NombrACT = rows["NombrACT"].ToString();

                                List<string> crudVariables = new List<string>() { DirRed, DirRedACT, Nom, FecIni, FecFin, NomACT, Nomb, NombACT, Cod, Nombr, TipDato, Val, NombrACT };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);
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
                                        ////VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 27, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //INSERTAR
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        Thread.Sleep(2000);
                                        CrudKnmmcpar.CRUD_DETALLE1(desktopSession, "0", crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Registro", true, file);

                                        ////APLICAR
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        //VALIDACIÓN INSERTAR
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, DirRed, DirRed, Campo1, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        CrudKnmmcpar.OPCIONES(desktopSession, "Temporalidad", crudVariables, file);

                                        ////////AGREGAR REGISTRO DETALLE 1
                                        Dictionary<string, Point> botonesNavegador1 = new Dictionary<string, Point>();
                                        botonesNavegador1 = PruebaCRUD.findname(desktopSession, 29, 1);
                                        var ElementList1 = PruebaCRUD.NavClass(desktopSession);

                                        //////INSERTAR DETALLE 1
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Insertar"].X, botonesNavegador1["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKnmmcpar.CRUD_DETALLE2(desktopSession, "0", crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Registro DETALLE 2", true, file);
                                        Thread.Sleep(1000);

                                        ////////////APLICAR DETALLE 1
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        //VALIDACIÓN INSERTAR DETALLE 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, Campo2, Nom, Nom, Campo2, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //////////////EDITAR CAMBIOS DETALLE 1 (EDITAR - CANCELAR)
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro DETALLE 2", true, file);

                                        CrudKnmmcpar.CRUD_DETALLE2(desktopSession, "1", crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados DETALLE 2", true, file);
                                        Thread.Sleep(1000);

                                        ////////////CANCELAR EDITAR DETALLE 1
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Cancelar"].X, botonesNavegador1["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada DETALLE 2", true, file);

                                        //VALIDACIÓN CANCELAR-EDITAR DETALLE 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, Campo2, Nom, Nom, EdCampo2, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////////////EDITAR CAMBIOS DETALLE 1 (EDITAR - APLICAR)
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro DETALLE 2", true, file);

                                        CrudKnmmcpar.CRUD_DETALLE2(desktopSession, "1", crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados DETALLE 2", true, file);
                                        Thread.Sleep(1000);

                                        ////////////APLICAR DETALLE 1
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Insertados DETALLE 2", true, file);

                                        Thread.Sleep(1000);
                                        bool valEditOra2 = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra2 == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList1[0].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar Registro DETALLE 2", true, file);

                                            CrudKnmmcpar.CRUD_DETALLE2(desktopSession, "1", crudVariables, file);

                                            desktopSession.Mouse.MouseMove(ElementList1[0].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados Insertados DETALLE 2", true, file);
                                        }
                                        ////VALIDACIÓN APLICAR-EDITAR
                                        Thread.Sleep(1500);
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla2, motor, user, Campo2, Nom, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, Campo2, Nom, NomACT, EdCampo2, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);

                                        CrudKnmmcpar.OPCIONES(desktopSession, "SubComisiones", crudVariables, file);

                                        ////////AGREGAR REGISTRO DETALLE 2
                                        Dictionary<string, Point> botonesNavegador2 = new Dictionary<string, Point>();
                                        botonesNavegador2 = PruebaCRUD.findname(desktopSession, 29, 1);
                                        var ElementList2 = PruebaCRUD.NavClass(desktopSession);

                                        //////INSERTAR DETALLE 2
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Insertar"].X, botonesNavegador2["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKnmmcpar.CRUD_DETALLE3(desktopSession, "0", crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Registro DETALLE 3", true, file);
                                        Thread.Sleep(1000);

                                        ////////////APLICAR DETALLE 2
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        //VALIDACIÓN INSERTAR DETALLE 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla3, user, motor, Campo3, Nomb, Nomb, Campo3, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //////////////EDITAR CAMBIOS DETALLE 2 (EDITAR - CANCELAR)
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro DETALLE 3", true, file);

                                        CrudKnmmcpar.CRUD_DETALLE3(desktopSession, "1", crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados DETALLE 3", true, file);
                                        Thread.Sleep(1000);

                                        ////////////CANCELAR EDITAR DETALLE 2
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Cancelar"].X, botonesNavegador2["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        //VALIDACIÓN CANCELAR-EDITAR DETALLE 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla3, user, motor, Campo3, Nomb, Nomb, EdCampo3, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////////////EDITAR CAMBIOS DETALLE 2 (EDITAR - APLICAR)
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro DETALLE 3", true, file);

                                        CrudKnmmcpar.CRUD_DETALLE3(desktopSession, "1", crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados DETALLE 3", true, file);
                                        Thread.Sleep(1000);

                                        ////////////APLICAR DETALLE 2
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Insertados DETALLE 3", true, file);

                                        Thread.Sleep(1000);
                                        bool valEditOra3 = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra3 == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList2[0].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar Registro DETALLE 3", true, file);

                                            CrudKnmmcpar.CRUD_DETALLE3(desktopSession, "1", crudVariables, file);

                                            desktopSession.Mouse.MouseMove(ElementList2[0].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados Insertados DETALLE 3", true, file);
                                        }
                                        ////VALIDACIÓN APLICAR-EDITAR
                                        Thread.Sleep(1500);
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla3, motor, user, Campo3, Nomb, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla3, user, motor, Campo3, Nomb, NombACT, EdCampo3, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);
                                        CrudKnmmcpar.OPCIONES(desktopSession, "Variables", crudVariables, file);

                                        ///////AGREGAR REGISTRO DETALLE 3
                                        Dictionary<string, Point> botonesNavegador3 = new Dictionary<string, Point>();
                                        botonesNavegador3 = PruebaCRUD.findname(desktopSession, 29, 1);
                                        var ElementList3 = PruebaCRUD.NavClass(desktopSession);

                                        //////INSERTAR DETALLE 3
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Insertar"].X, botonesNavegador3["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKnmmcpar.CRUD_DETALLE4(desktopSession, "0", crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Registro DETALLE 4", true, file);
                                        Thread.Sleep(1000);

                                        ////////////APLICAR DETALLE 3
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Aplicar"].X, botonesNavegador3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        //VALIDACIÓN INSERTAR DETALLE 3
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla4, user, motor, Campo4, Cod, Cod, Campo4, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //////////////EDITAR CAMBIOS DETALLE 3 (EDITAR - CANCELAR)
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Editar"].X, botonesNavegador3["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro DETALLE 4", true, file);

                                        CrudKnmmcpar.CRUD_DETALLE4(desktopSession, "1", crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados DETALLE 4", true, file);
                                        Thread.Sleep(1000);

                                        ////////////CANCELAR EDITAR DETALLE 3
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Cancelar"].X, botonesNavegador3["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        //VALIDACIÓN CANCELAR-EDITAR DETALLE 3
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla4, user, motor, Campo4, Cod, Nombr, EdCampo4, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////////////EDITAR CAMBIOS DETALLE 3 (EDITAR - APLICAR)
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Editar"].X, botonesNavegador3["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro DETALLE 4", true, file);

                                        CrudKnmmcpar.CRUD_DETALLE4(desktopSession, "1", crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados DETALLE 4", true, file);
                                        Thread.Sleep(1000);

                                        ////////////APLICAR DETALLE 3
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Aplicar"].X, botonesNavegador3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Insertados DETALLE 4", true, file);

                                        Thread.Sleep(1000);
                                        bool valEditOra4 = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra4 == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList3[0].Coordinates, botonesNavegador3["Editar"].X, botonesNavegador3["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar Registro DETALLE 4", true, file);

                                            CrudKnmmcpar.CRUD_DETALLE4(desktopSession, "1", crudVariables, file);

                                            desktopSession.Mouse.MouseMove(ElementList3[0].Coordinates, botonesNavegador3["Aplicar"].X, botonesNavegador3["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados Insertados DETALLE 4", true, file);
                                        }
                                        ////VALIDACIÓN APLICAR-EDITAR
                                        Thread.Sleep(1500);
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla4, motor, user, Campo4, Cod, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla4, user, motor, Campo4, Cod, NombrACT, EdCampo4, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);

                                        CrudKnmmcpar.OPCIONES(desktopSession, "Generales", crudVariables, file);

                                        //////EDITAR CAMBIOS (EDITAR - CANCELAR)
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);

                                        CrudKnmmcpar.CRUD_DETALLE1(desktopSession, "1", crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);
                                        Thread.Sleep(2000);

                                        ////CANCELAR EDITAR
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada", true, file);

                                        //VALIDACIÓN CANCELAR-EDITAR
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, DirRed, DirRed, EdCampo1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        //////EDITAR CAMBIOS (EDITAR - APLICAR)
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);

                                        CrudKnmmcpar.CRUD_DETALLE1(desktopSession, "1", crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);
                                        Thread.Sleep(2000);

                                        ////APLICAR EDITAR
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Insertados", true, file);

                                        Thread.Sleep(1000);
                                        bool valEditOra1 = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra1 == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);

                                            CrudKnmmcpar.CRUD_DETALLE1(desktopSession, "1", crudVariables, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados Insertados", true, file);
                                        }
                                        ////VALIDACIÓN APLICAR-EDITAR
                                        Thread.Sleep(1500);
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla1, motor, user, Campo1, DirRed, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, DirRed, DirRedACT, EdCampo1, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);

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

                                        //RECTIFICACION DE CAMPOS DE EXCEL  
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla1, user, motor, Campo1, DirRed, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        Thread.Sleep(1000);
                                        CrudKnmmcpar.OPCIONES(desktopSession, "Variables", crudVariables, file);
                                        /////////////// ELIMINAR REGISTRO DETALLE 3
                                        CrudKnmmcpar.EliminarRegistro(desktopSession, ElementList3[1], botonesNavegador3, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados DETALLE 3", true, file);
                                        //VALIDACIÓN ELIMINAR REGISTRO 3
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla4, motor, user, Campo4, Cod, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla4, user, motor, Campo4, Cod, "", Campo4, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        Thread.Sleep(1000);
                                        CrudKnmmcpar.OPCIONES(desktopSession, "SubComisiones", crudVariables, file);
                                        /////////////// ELIMINAR REGISTRO DETALLE 2
                                        CrudKnmmcpar.EliminarRegistro(desktopSession, ElementList2[1], botonesNavegador2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados DETALLE 2", true, file);
                                        //VALIDACIÓN ELIMINAR REGISTRO 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla3, motor, user, Campo3, Nomb, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla3, user, motor, Campo3, Nomb, "", Campo3, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        Thread.Sleep(1000);
                                        CrudKnmmcpar.OPCIONES(desktopSession, "Temporalidad", crudVariables, file);
                                        /////////////// ELIMINAR REGISTRO DETALLE 1
                                        CrudKnmmcpar.EliminarRegistro(desktopSession, ElementList1[1], botonesNavegador1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados DETALLE 1", true, file);
                                        //VALIDACIÓN ELIMINAR REGISTRO 1 
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla2, motor, user, Campo2, Nom, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, Campo2, Nom, "", Campo2, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        Thread.Sleep(1000);
                                        CrudKnmmcpar.OPCIONES(desktopSession, "Generales", crudVariables, file);
                                        newFunctions_3.LupaAud(desktopSession, "0", file);
                                        /////////// ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados", true, file);
                                        //VALIDACIÓN ELIMINAR REGISTRO
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla1, motor, user, Campo1, DirRed, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, DirRed, "", Campo1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);
                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }
                                    //CONTADOR DE ERRORES
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
        public void KNmMcmacCheckList()
        {
            List<string> listaResultado;
            List<string> Celda = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_64.KNmMcmacCheckList")
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
                                rows["Tabla1"].ToString().Length != 0 && rows["Tabla1"].ToString() != null &&
                                rows["Campo1"].ToString().Length != 0 && rows["Campo1"].ToString() != null &&
                                rows["EdCampo1"].ToString().Length != 0 && rows["EdCampo1"].ToString() != null &&
                                rows["Tabla2"].ToString().Length != 0 && rows["Tabla2"].ToString() != null &&
                                rows["Campo2"].ToString().Length != 0 && rows["Campo2"].ToString() != null &&
                                rows["EdCampo2"].ToString().Length != 0 && rows["EdCampo2"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&

                                //VARIABLES CRUD 1
                                rows["Contrato"].ToString().Length != 0 && rows["Contrato"].ToString() != null &&
                                rows["Identificacion"].ToString().Length != 0 && rows["Identificacion"].ToString() != null &&
                                rows["TipComi"].ToString().Length != 0 && rows["TipComi"].ToString() != null &&
                                rows["SubComiTemp"].ToString().Length != 0 && rows["SubComiTemp"].ToString() != null &&
                                rows["LiquiConcp"].ToString().Length != 0 && rows["LiquiConcp"].ToString() != null &&
                                rows["Val"].ToString().Length != 0 && rows["Val"].ToString() != null &&
                                rows["FecNov"].ToString().Length != 0 && rows["FecNov"].ToString() != null &&
                                rows["ValACT"].ToString().Length != 0 && rows["ValACT"].ToString() != null &&
                                //VARIABLES CRUD 2
                                rows["Descrip"].ToString().Length != 0 && rows["Descrip"].ToString() != null &&
                                rows["Valor"].ToString().Length != 0 && rows["Valor"].ToString() != null &&
                                rows["ValorACT"].ToString().Length != 0 && rows["ValorACT"].ToString() != null 
                                )
                            {
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string Tabla1 = rows["Tabla1"].ToString();
                                string Campo1 = rows["Campo1"].ToString();
                                string EdCampo1 = rows["EdCampo1"].ToString();
                                string Tabla2 = rows["Tabla2"].ToString();
                                string Campo2 = rows["Campo2"].ToString();
                                string EdCampo2 = rows["EdCampo2"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();

                                string Contrato = rows["Contrato"].ToString();
                                string Identificacion = rows["Identificacion"].ToString();
                                string TipComi = rows["TipComi"].ToString();
                                string SubComiTemp = rows["SubComiTemp"].ToString();
                                string LiquiConcp = rows["LiquiConcp"].ToString();
                                string Val = rows["Val"].ToString();
                                string FecNov = rows["FecNov"].ToString();
                                string ValACT = rows["ValACT"].ToString();
                                string Descrip = rows["Descrip"].ToString();
                                string Valor = rows["Valor"].ToString();
                                string ValorACT = rows["ValorACT"].ToString();

                                List<string> crudVariables = new List<string>() { Contrato, Identificacion, TipComi, SubComiTemp, LiquiConcp, Val, FecNov, ValACT };
                                List<string> crudVariablesDETALLE = new List<string> { Descrip, Valor, ValorACT };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);
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
                                        ////VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 27, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //INSERTAR
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKnmmcmac.CRUD(desktopSession, "0", crudVariables, file);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Insertar Registro", true, file);

                                        ////APLICAR
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        Thread.Sleep(5000);
                                        //VALIDACIÓN INSERTAR
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Identificacion, Identificacion, Campo1, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        CrudKnmmcmac.CRUD_DETALLE(desktopSession, "0", crudVariablesDETALLE, file);

                                        ////////AGREGAR REGISTRO DETALLE
                                        Dictionary<string, Point> botonesNavegador1 = new Dictionary<string, Point>();
                                        botonesNavegador1 = PruebaCRUD.findname(desktopSession, 29, 1);
                                        var ElementList1 = PruebaCRUD.NavClass(desktopSession);

                                        //////INSERTAR DETALLE
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Insertar"].X, botonesNavegador1["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKnmmcmac.CRUD_DETALLE(desktopSession, "1", crudVariablesDETALLE, file);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Insertar Registro DETALLE", true, file);

                                        //////////APLICAR DETALLE
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        ////VALIDACIÓN INSERTAR DETALLE
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, Campo2, Descrip, Descrip, Campo2, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////////////EDITAR CAMBIOS DETALLE (EDITAR - CANCELAR)
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro DETALLE", true, file);

                                        CrudKnmmcmac.CRUD_DETALLE(desktopSession, "2", crudVariablesDETALLE, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados DETALLE", true, file);
                                        Thread.Sleep(2000);

                                        //////////CANCELAR EDITAR DETALLE
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Cancelar"].X, botonesNavegador1["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada DETALLE", true, file);

                                        ////VALIDACIÓN CANCELAR-EDITAR
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, Campo2, Descrip, Valor, EdCampo2, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ///////////EDITAR CAMBIOS DETALLE (EDITAR - APLICAR)
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro DETALLE", true, file);

                                        CrudKnmmcmac.CRUD_DETALLE(desktopSession, "2", crudVariablesDETALLE, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados DETALLE", true, file);
                                        Thread.Sleep(2000);

                                        ////APLICAR EDITAR DETALLE
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Insertados DETALLE", true, file);

                                        Thread.Sleep(1000);
                                        bool valEditOra1 = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra1 == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar Registro DETALLE", true, file);

                                            CrudKnmmcmac.CRUD_DETALLE(desktopSession, "2", crudVariablesDETALLE, file);

                                            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados Insertados DETALLE", true, file);
                                        }
                                        ////VALIDACIÓN APLICAR-EDITAR
                                        Thread.Sleep(1500);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla2, motor, user, Campo2, Descrip, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, Campo2, Descrip, ValorACT, EdCampo2, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //////EDITAR CAMBIOS (EDITAR - CANCELAR)
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);

                                        CrudKnmmcmac.CRUD(desktopSession, "1", crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);
                                        Thread.Sleep(2000);

                                        ////CANCELAR EDITAR
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada", true, file);

                                        //VALIDACIÓN CANCELAR-EDITAR
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Identificacion, Val, EdCampo1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        //////EDITAR CAMBIOS (EDITAR - APLICAR)
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);

                                        CrudKnmmcmac.CRUD(desktopSession, "1", crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);
                                        Thread.Sleep(2000);

                                        ////APLICAR EDITAR
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Insertados", true, file);

                                        Thread.Sleep(1000);
                                        bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);

                                            CrudKnmmcmac.CRUD(desktopSession, "1", crudVariables, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados Insertados", true, file);
                                        }
                                        ////VALIDACIÓN APLICAR-EDITAR
                                        Thread.Sleep(1500);
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla1, motor, user, Campo1, Identificacion, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Identificacion, ValACT, EdCampo1, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);

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

                                        //RECTIFICACION DE CAMPOS DE EXCEL  
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla1, user, motor, Campo1, Identificacion, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        Thread.Sleep(1000);
                                        //ELIMINAR REGISTRO DETALLE
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[0], botonesNavegador1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados DETALLE", true, file);
                                        //VALIDACIÓN ELIMINAR REGISTRO DETALLE
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla2, motor, user, Campo2, Descrip, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, Campo2, Descrip, "", Campo2, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        Thread.Sleep(1000);
                                        //ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados", true, file);
                                        //VALIDACIÓN ELIMINAR REGISTRO
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla1, motor, user, Campo1, Identificacion, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Identificacion, "", Campo1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);
                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }
                                    //CONTADOR DE ERRORES
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
        public void KNmDifinCheckList()
        {
            List<string> listaResultado;
            List<string> Celda = new List<string>();
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_64.KNmDifinCheckList")
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
                                rows["Tabla1"].ToString().Length != 0 && rows["Tabla1"].ToString() != null &&
                                rows["Campo1"].ToString().Length != 0 && rows["Campo1"].ToString() != null &&
                                rows["EdCampo1"].ToString().Length != 0 && rows["EdCampo1"].ToString() != null &&
                                rows["Tabla2"].ToString().Length != 0 && rows["Tabla2"].ToString() != null &&
                                rows["Campo2"].ToString().Length != 0 && rows["Campo2"].ToString() != null &&
                                rows["EdCampo2"].ToString().Length != 0 && rows["EdCampo2"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                //rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&

                                //VARIABLES CRUD 2
                                rows["CodVar"].ToString().Length != 0 && rows["CodVar"].ToString() != null &&
                                rows["NomVar"].ToString().Length != 0 && rows["NomVar"].ToString() != null &&
                                rows["FecCrea"].ToString().Length != 0 && rows["FecCrea"].ToString() != null &&
                                rows["CodVarACT"].ToString().Length != 0 && rows["CodVarACT"].ToString() != null &&
                                //VARIABLES CRUD 1
                                rows["Codigo"].ToString().Length != 0 && rows["Codigo"].ToString() != null &&
                                rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
                                rows["NombreACT"].ToString().Length != 0 && rows["NombreACT"].ToString() != null
                                )
                            {
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string Tabla1 = rows["Tabla1"].ToString();
                                string Campo1 = rows["Campo1"].ToString();
                                string EdCampo1 = rows["EdCampo1"].ToString();
                                string Tabla2 = rows["Tabla2"].ToString();
                                string Campo2 = rows["Campo2"].ToString();
                                string EdCampo2 = rows["EdCampo2"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();

                                string Codigo = rows["Codigo"].ToString();
                                string Nombre = rows["Nombre"].ToString();
                                string NombreACT = rows["NombreACT"].ToString();
                                string CodVar = rows["CodVar"].ToString();
                                string NomVar = rows["NomVar"].ToString();
                                string FecCrea = rows["FecCrea"].ToString();
                                string CodVarACT = rows["CodVarACT"].ToString();
                                
                                List<string> crudVariables = new List<string>() { Codigo, Nombre, NombreACT };
                                List<string> crudVariablesDETALLE = new List<string> { CodVar, NomVar, FecCrea, CodVarACT };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);
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
                                        ////VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 27, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //INSERTAR
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKnmdifin.CRUD(desktopSession, "0", crudVariables, file);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Insertar Registro", true, file);

                                        ////APLICAR
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        //VALIDACIÓN INSERTAR
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Codigo, Codigo, Campo1, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////////AGREGAR REGISTRO DETALLE
                                        Dictionary<string, Point> botonesNavegador1 = new Dictionary<string, Point>();
                                        botonesNavegador1 = PruebaCRUD.findname(desktopSession, 29, 1);
                                        var ElementList1 = PruebaCRUD.NavClass(desktopSession);

                                        //////INSERTAR DETALLE
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Insertar"].X, botonesNavegador1["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKnmdifin.CRUD_DETALLE(desktopSession, "0", crudVariablesDETALLE, file);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Insertar Registro DETALLE", true, file);

                                        //////////APLICAR DETALLE
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        ////VALIDACIÓN INSERTAR DETALLE
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, Campo2, NomVar, NomVar, Campo2, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////////////EDITAR CAMBIOS DETALLE (EDITAR - CANCELAR)
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro DETALLE", true, file);

                                        CrudKnmdifin.CRUD_DETALLE(desktopSession, "1", crudVariablesDETALLE, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados DETALLE", true, file);
                                        Thread.Sleep(2000);

                                        //////////CANCELAR EDITAR DETALLE
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Cancelar"].X, botonesNavegador1["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada DETALLE", true, file);

                                        ////VALIDACIÓN CANCELAR-EDITAR
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, Campo2, NomVar, CodVar, EdCampo2, c, ErrorValidacion, "No se edito correctamente", 1);

                                        newFunctions_3.LupaAud(desktopSession, "0", file);

                                        ///////////EDITAR CAMBIOS DETALLE (EDITAR - APLICAR)
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro DETALLE", true, file);

                                        CrudKnmdifin.CRUD_DETALLE(desktopSession, "1", crudVariablesDETALLE, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados DETALLE", true, file);
                                        Thread.Sleep(2000);

                                        ////APLICAR EDITAR DETALLE
                                        desktopSession.Mouse.MouseMove(ElementList1[0].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Insertados DETALLE", true, file);

                                        Thread.Sleep(1000);
                                        bool valEditOra1 = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra1 == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList1[0].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar Registro DETALLE", true, file);

                                            CrudKnmdifin.CRUD_DETALLE(desktopSession, "1", crudVariablesDETALLE, file);
                                            Thread.Sleep(1000);

                                            desktopSession.Mouse.MouseMove(ElementList1[0].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados Insertados DETALLE", true, file);
                                        }
                                        ////VALIDACIÓN APLICAR-EDITAR
                                        Thread.Sleep(1500);
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla2, motor, user, Campo2, NomVar, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, Campo2, NomVar, CodVarACT, EdCampo2, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //////EDITAR CAMBIOS (EDITAR - CANCELAR)
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);

                                        CrudKnmdifin.CRUD(desktopSession, "1", crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);
                                        Thread.Sleep(2000);

                                        ////CANCELAR EDITAR
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada", true, file);

                                        //VALIDACIÓN CANCELAR-EDITAR
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Codigo, Nombre, EdCampo1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        //////EDITAR CAMBIOS (EDITAR - APLICAR)
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);

                                        CrudKnmdifin.CRUD(desktopSession, "1", crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);
                                        Thread.Sleep(2000);

                                        ////APLICAR EDITAR
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Insertados", true, file);

                                        Thread.Sleep(1000);
                                        bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);

                                            CrudKnmdifin.CRUD(desktopSession, "1", crudVariables, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados Insertados", true, file);
                                        }
                                        ////VALIDACIÓN APLICAR-EDITAR
                                        Thread.Sleep(1500);
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla1, motor, user, Campo1, Codigo, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Codigo, NombreACT, EdCampo1, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(4000);

                                        //SUMATORIA
                                        //errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //newFunctions_1.lapiz(desktopSession);


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
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla1, user, motor, Campo1, Codigo, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        Thread.Sleep(1000);
                                        //ELIMINAR REGISTRO DETALLE
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[1], botonesNavegador1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados DETALLE", true, file);
                                        //VALIDACIÓN ELIMINAR REGISTRO DETALLE
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla2, motor, user, Campo2, NomVar, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, Campo2, NomVar, "", Campo2, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        Thread.Sleep(1000);
                                        //ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados", true, file);
                                        //VALIDACIÓN ELIMINAR REGISTRO
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla1, motor, user, Campo1, Codigo, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Codigo, "", Campo1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);
                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }
                                    //CONTADOR DE ERRORES
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

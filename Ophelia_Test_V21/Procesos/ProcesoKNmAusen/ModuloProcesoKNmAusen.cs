﻿using System;
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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_11;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesosIncap;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesosKNmAusen;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_8;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_10;
using OpenQA.Selenium.Appium;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosEd;

namespace Ophelia_Test_V21.Procesos.ProcesoKNmAusen
{
    /// <summary>
    /// Descripción resumida de ModuloProcesoKNmAusen
    /// </summary>
    [TestClass]
    public class ModuloProcesoKNmAusen : FuncionesVitales
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
        public ModuloProcesoKNmAusen()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesoKNmAusen.ModuloProcesoKNmAusen.RegistroKNmAusen")
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

                               
                                //string delete2 = $"delete from nm_noved where cod_empl= '{Id}'";
                                //FuncionesVitales.UpdateDeleteInsert(delete2, motor, user);
                                //string delete3 = $"delete from nm_ausen where cod_empl= '{Id}'";
                                //FuncionesVitales.UpdateDeleteInsert(delete3, motor, user);
                                //string delete4 = $"delete from nm_diasn where cod_empl= '{Id}'";
                                //FuncionesVitales.UpdateDeleteInsert(delete4, motor, user);
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
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
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
        //[TestMethod]
        //public void CorrerKNmLinomAusen()
        //{

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

        //        if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesoKNmAusen.ModuloProcesoKNmAusen.CorrerKNmLinomAusen")
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

        //                        //Variables Validaciones Crud principal
        //                        rows["Id"].ToString().Length != 0 && rows["Id"].ToString() != null)

        //                    {
        //                        //Variables obligatorias
        //                        string user = rows["User"].ToString();
        //                        string motor = rows["Motor"].ToString();
        //                        string codProgram = rows["CodProgram"].ToString();
        //                        string nomProgram = rows["NomProgram"].ToString();

        //                        //Variables generales
        //                        string TipoQbe = rows["TipoQbe"].ToString();
        //                        string QbeFilter = rows["QbeFilter"].ToString();

        //                        //Variables Crud principal
        //                        string Id = rows["Id"].ToString();

        //                        //String Codigo = "9";



        //                        //LISTA CRUD PRINCIPAL
        //                        List<string> crudPrincipal = new List<string>() { Id };

        //                        //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
        //                        string file = CrearDocumentoWordDinamico(codProgram, "LiquidacionNominaAusentismos", motor);

        //                        try
        //                        {
        //                            List<string> ErrorValidacion = new List<string>();
        //                            List<string> errors = new List<string>();

        //                            AbrirPrograma a = new AbrirPrograma(codProgram, user);
        //                            desktopSession = a.Start2(motor, "");
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


        //                                //Proyectar Fechas de liquidacion
        //                                FuncionesProcesosKNmAusen.CorrerNomina(desktopSession, 0, crudPrincipal);
        //                                newFunctions_4.ScreenshotNuevo("Datos Proceso", true, file);
        //                                Thread.Sleep(1500);

        //                                //QBE
        //                                errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                //Iniciar Proceso
        //                                newFunctions_4.ScreenshotNuevo("Empleado a Liquidar", true, file);
        //                                FuncionesProcesosKNmAusen.CorrerNomina(desktopSession, 1, crudPrincipal);
        //                                Thread.Sleep(10000);
        //                                newFunctions_4.ScreenshotNuevo("Proceso Iniciado", true, file);
        //                                WindowsDriver<WindowsElement> RootSession2 = null;
        //                                RootSession2 = PruebaCRUD.RootSession();
        //                                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //                                Thread.Sleep(10000);
        //                                newFunctions_4.ScreenshotNuevo("Proceso Finalizado", true, file);
        //                                WindowsDriver<WindowsElement> RootSession3 = null;
        //                                RootSession3 = PruebaCRUD.RootSession();
        //                                RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //                                Thread.Sleep(2000);

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
        //                            a.Stop();
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
        //                            //break;

        //                        }
        //                        catch (Exception e)
        //                        {
        //                            Thread.Sleep(4000);

        //                            //Add Juan
        //                            newFunctions_4.ScreenshotNuevo("Error en la prueba", true, file);

        //                            //Add Daniel
        //                            TestContext.AddResultFile(newFunctions_4.ScreenshotInicio(codProgram));
        //                            //
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
        //                            Assert.Fail(CaseId + " ::::::" + e.ToString());
        //                            break;
        //                        }
        //                    }
        //                }
        //                //break;
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
        //public void VerificarRegistrosPrenomina()
        //{

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

        //        if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesoKNmAusen.ModuloProcesoKNmAusen.VerificarRegistrosPrenomina")
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
        //                        rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null)



        //                    {
        //                        //Variables obligatorias
        //                        string user = rows["User"].ToString();
        //                        string motor = rows["Motor"].ToString();
        //                        string codProgram = rows["CodProgram"].ToString();
        //                        string nomProgram = rows["NomProgram"].ToString();

        //                        //Variables generales
        //                        string TipoQbe = rows["TipoQbe"].ToString();
        //                        string QbeFilter = rows["QbeFilter"].ToString();


        //                        String Codigo = "1";



        //                        //LISTA CRUD PRINCIPAL
        //                        List<string> crudPrincipal = new List<string>() { };

        //                        //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
        //                        string file = CrearDocumentoWordDinamico(codProgram, "RegistrosAusentismoPrenomina", motor);


        //                        try
        //                        {
        //                            List<string> ErrorValidacion = new List<string>();
        //                            List<string> errors = new List<string>();

        //                            AbrirPrograma a = new AbrirPrograma(codProgram, user);
        //                            desktopSession = a.Start2(motor, "");
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

        //                                //QBE
        //                                errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, Codigo);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                //Iniciar Proceso
        //                                newFunctions_4.ScreenshotNuevo("Registros a Validar", true, file);
        //                                FuncionesProcesosKNmAusen.VerificarPrenomina(desktopSession, 0, crudPrincipal);
        //                                Thread.Sleep(2000);
        //                                newFunctions_4.ScreenshotNuevo("Registros", true, file);

        //                                //Verificar Registros 1X1
        //                                for (int i = 0; i < 4; i++)
        //                                {
        //                                    FuncionesProcesosKNmAusen.VerificarPrenomina(desktopSession, 1, crudPrincipal);
        //                                    Thread.Sleep(2000);
        //                                    newFunctions_4.ScreenshotNuevo("Registros # " + i++, true, file);
        //                                }

        //                                string delete1 = $"delete from nm_preno where cod_empl= '{QbeFilter}'";
        //                                FuncionesVitales.UpdateDeleteInsert(delete1, motor, user);

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
        //                            a.Stop();
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
        //                            //break;

        //                        }
        //                        catch (Exception e)
        //                        {
        //                            Thread.Sleep(4000);

        //                            //Add Juan
        //                            newFunctions_4.ScreenshotNuevo("Error en la prueba", true, file);

        //                            //Add Daniel
        //                            TestContext.AddResultFile(newFunctions_4.ScreenshotInicio(codProgram));
        //                            //
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
        //                            Assert.Fail(CaseId + " ::::::" + e.ToString());
        //                            break;
        //                        }
        //                    }
        //                }
        //                //break;
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

        //[TestCleanup]
        //public void Limpiar()
        //{
        //    AbrirPrograma a = new AbrirPrograma();
        //    if (desktopSession != null)
        //    {
        //        desktopSession.Close();
        //        desktopSession.Dispose();
        //    }
        //    //    //a.Stop();
        //}
    }
}

using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosIG;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.ValidacionQueries;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading;

namespace Ophelia_Test_V21.TestMethods.ModulosIG
{
    /// <summary>
    /// Summary description for UnitTest1
        [TestClass]
        public class ModulosIG_2 : FuncionesVitales
        {
            protected static WindowsDriver<WindowsElement> session;
            protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
            protected static WindowsDriver<WindowsElement> desktopSession;
            protected static WindowsDriver<WindowsElement> rootSession;
            List<string> listaResultado;
            List<string> Celda = new List<string>();

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

            [TestMethod]
            public void KIgPindmChecklist()
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

                    if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosIG.ModulosIG_2.KIgPindmChecklist")
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
                                    rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                    
                                    rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                    rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                    rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                    rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                    rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                    rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null)
                                                                                                          
                                {
                                    string user = rows["User"].ToString();
                                    string motor = rows["Motor"].ToString();
                                    /*string Tabla = rows["Tabla"].ToString()*/;
                                    //string Campo = rows["Campo"].ToString();
                                    string codProgram = rows["CodProgram"].ToString();
                                    string nomProgram = rows["NomProgram"].ToString();

                                    string TipoQbe = rows["TipoQbe"].ToString();
                                    string QbeFilter = rows["QbeFilter"].ToString();
                                    string banExcel = rows["banExcel"].ToString();
                                    string BanderaLupa = rows["BanderaLupa"].ToString();
                                    string BanderaPreli = rows["BanderaPreli"].ToString();
                                    string BanderaSum = rows["BanderaSum"].ToString();

                                    //string fecha = rows["fecha"].ToString();
                                    //string indiIni = rows["indiIni"].ToString();
                                    //string indiFin = rows["indiFin"].ToString();
                                    //List<string> data = new List<string> { fecha, indiIni, indiFin };

                                    //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                    string file = CrearDocumentoWordDinamico(codProgram, "IG", motor);

                                    try
                                    {
                                        List<string> ErrorValidacion = new List<string>();
                                        List<string> errors = new List<string>();

                                        AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                        desktopSession = a.Start2(motor, "");
                                        Thread.Sleep(2000);

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
                                            newFunctions_1.lapiz(desktopSession);
                                            Thread.Sleep(1000);




                                            // pasar a un lado                                            
                                            CrudKigpindm.Crud(desktopSession, 1, file);
                                            Thread.Sleep(2500);


                                            //REPORTE DINAMICO
                                            string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
                                            errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                            //REPORTE PRELIMINAR
                                            string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                            errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                            //EXPORTAR EXCEL
                                            string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                            errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            Thread.Sleep(2000);


                                            // pasar al otro lado
                                            CrudKigpindm.Crud(desktopSession, 2, file);
                                            Thread.Sleep(1000);

                                            newFunctions_4.ScreenshotNuevo("Programa finalizado", true, file);




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

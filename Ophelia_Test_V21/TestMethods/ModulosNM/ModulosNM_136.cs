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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_12;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_9;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_10;
using OpenQA.Selenium.Appium;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosEd;

namespace Ophelia_Test_V21.TestMethods.ModulosNM
{
    [TestClass]
    public class ModulosNM_136 : FuncionesVitales
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
        public ModulosNM_136()
        { }

        private TestContext testContextInstance;

        public TestContext TestContext
        {
            get
            { return testContextInstance; }
            set
            { testContextInstance = value; }
        }

        //[TestMethod]
        //public void KNmRanckChecKlist()
        //{
        //    //Debugger.Launch();
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

        //        if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_136.KNmRanckhecKlist")
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
        //                        rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
        //                        rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
        //                        rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
        //                        rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&

        //                        //Variables Validaciones Crud principal
        //                        rows["Cod"].ToString().Length != 0 && rows["Cod"].ToString() != null &&
        //                        rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
        //                        rows["Valori"].ToString().Length != 0 && rows["Valori"].ToString() != null &&


        //                        //Variables Validaciones detalle 1
        //                        rows["Kmi"].ToString().Length != 0 && rows["Kmi"].ToString() != null &&
        //                        rows["Kmf"].ToString().Length != 0 && rows["Kmf"].ToString() != null &&
        //                        rows["Valor"].ToString().Length != 0 && rows["Valor"].ToString() != null &&
        //                        rows["Kmi2"].ToString().Length != 0 && rows["Kmi2"].ToString() != null &&

        //                        //Variables Validaciones 
        //                        rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
        //                        rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null



        //                        )

        //                    {
        //                        //Variables obligatorias
        //                        string user = rows["User"].ToString();
        //                        string motor = rows["Motor"].ToString();
        //                        string codProgram = rows["CodProgram"].ToString();
        //                        string nomProgram = rows["NomProgram"].ToString();

        //                        //Variables generales
        //                        string TipoQbe = rows["TipoQbe"].ToString();
        //                        string QbeFilter = rows["QbeFilter"].ToString();
        //                        string banExcel = rows["banExcel"].ToString();
        //                        string BanderaLupa = rows["BanderaLupa"].ToString();
        //                        string BanderaPreli = rows["BanderaPreli"].ToString();
        //                        string BanderaSum = rows["BanderaSum"].ToString();

        //                        //Variables validacion  crud Principal
        //                        string Tabla = rows["Tabla"].ToString();
        //                        string Campo = rows["Campo"].ToString();



        //                        //Variables validacion  crud Principal
        //                        string Cod = rows["Cod"].ToString();
        //                        string Nombre = rows["Nombre"].ToString();
        //                        string Valori = rows["Valori"].ToString();


        //                        //Variables validacion DETALLE 1
        //                        string Kmi = rows["Kmi"].ToString();
        //                        string Kmf = rows["Kmf"].ToString();
        //                        string Valor = rows["Valor"].ToString();
        //                        string Kmi2 = rows["Kmi2"].ToString();





        //                        //LISTA CRUD PRINCIPAL
        //                        List<string> crudPrincipal = new List<string>() { Cod, Nombre, Valori };
        //                        //LISTA DETALLE 1
        //                        List<string> crudDet1 = new List<string>() { Kmi, Kmf, Valor, Kmi2 };


        //                        //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
        //                        string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);

        //                        try
        //                        {
        //                            List<string> ErrorValidacion = new List<string>();
        //                            List<string> errors = new List<string>();

        //                            AbrirPrograma a = new AbrirPrograma(codProgram, user);
        //                            desktopSession = a.Start2(motor, "");

        //                            Thread.Sleep(2500);
        //                            ////Debugger.Launch();
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
        //                                //VALIDACION CODIGO DEL PROGRAMA
        //                                errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                //VALIDACION NOMBRE DEL PROGRAMA
        //                                errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                //VERSION
        //                                FuncionesGlobales.GetVersion(desktopSession, file);
        //                                Thread.Sleep(1000);
        //                                //NOTAS
        //                                newFunctions_4.openInnerNote(desktopSession, file);

        //                                Thread.Sleep(1500);

        //                                //MODIFICAMOS
        //                                //AGREGAR REGISTRO Principal
        //                                Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
        //                                botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 0);
        //                                var ElementList = PruebaCRUD.NavClass(desktopSession);
        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
        //                                desktopSession.Mouse.Click(null);

        //                                CrudKnmRanck.AgregarCrud(desktopSession, 0, crudPrincipal, file);
        //                                newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

        //                                Thread.Sleep(1500);



        //                                ////DESCARTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Editar", true, file);
        //                                Thread.Sleep(1500);

        //                                CrudKnmRanck.AgregarCrud(desktopSession, 1, crudPrincipal, file);
        //                                newFunctions_4.ScreenshotNuevo("Editar Datos Crud Principal", true, file);
        //                                Thread.Sleep(1500);

        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

        //                                Thread.Sleep(1000);

        //                                //LUPA
        //                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                Thread.Sleep(1000);

        //                                // ACEPTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                Thread.Sleep(2000);

        //                                newFunctions_4.ScreenshotNuevo("Editar", true, file);
        //                                CrudKnmRanck.AgregarCrud(desktopSession, 1, crudPrincipal, file);
        //                                newFunctions_4.ScreenshotNuevo("Editar Datos Crud Principal", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                Thread.Sleep(3000);
        //                                newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);



        //                                ////AGREGAR REGISTRO DETALLE 1                                        
        //                                Dictionary<string, Point> botonesNavegador2 = new Dictionary<string, Point>();
        //                                botonesNavegador2 = PruebaCRUD.findname(desktopSession, 23, 1);
        //                                var ElementList2 = PruebaCRUD.NavClass(desktopSession);
        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Insertar"].X, botonesNavegador2["Insertar"].Y);
        //                                desktopSession.Mouse.Click(null);

        //                                CrudKnmRanck.CrudDet(desktopSession, 0, crudDet1, file);
        //                                newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle", true, file);

        //                                Thread.Sleep(1500);



        //                                //LUPA
        //                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                Thread.Sleep(1000);

        //                                //DESCARTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Editar detalle", true, file);

        //                                CrudKnmRanck.CrudDet(desktopSession, 1, crudDet1, file);
        //                                newFunctions_4.ScreenshotNuevo("Datos Editados detalle", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Cancelar"].X, botonesNavegador2["Cancelar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados detalle", true, file);

        //                                Thread.Sleep(1000);



        //                                ////ACEPTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Editar detalle", true, file);

        //                                CrudKnmRanck.CrudDet(desktopSession, 1, crudDet1, file);
        //                                newFunctions_4.ScreenshotNuevo("Datos Editados detalle", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Edición Aplicada detalle", true, file);








        //                                //QBE
        //                                errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                //LUPA
        //                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                //SUMATORIA
        //                                errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                newFunctions_1.lapiz(desktopSession);

        //                                //REPORTE DINAMICO
        //                                string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
        //                                errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                //Rectificar Pie Pagina PDF DINAMICO
        //                                listaResultado = Textopdf(pdf, codProgram, user);
        //                                listaResultado.ForEach(e => ErrorValidacion.Add(e));

        //                                //REPORTE PRELIMINAR
        //                                string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
        //                                errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
        //                                //Rectificar Pie Pagina PDF PRELIMINAR
        //                                listaResultado = Textopdf(pdf1, codProgram, user);
        //                                listaResultado.ForEach(e => ErrorValidacion.Add(e));

        //                                //EXPORTAR EXCEL
        //                                string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
        //                                errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                //LUPA
        //                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }



        //                                //PruebaCRUD.EliminarRegistro(desktopSession, ElementList2[1], botonesNavegador2, file);
        //                                //newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 1", true, file);

        //                                //Thread.Sleep(2000);

        //                                ////LUPA
        //                                //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                //PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
        //                                //newFunctions_4.ScreenshotNuevo("Datos Eliminados Principal", true, file);

        //                                Thread.Sleep(2000);
        //                                newFunctions_4.ScreenshotNuevo("Fin del programa", true, file);
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
        //                            break;

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
        //                                        break;
        //                                    }
        //                                }
        //                            }
        //                            Assert.Fail(CaseId + " ::::::" + e.ToString());
        //                            break;
        //                        }
        //                    }
        //                }
        //                break;
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
        public void KNmAbempChecKlist()
        {
            //Debugger.Launch();
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_136.KNmAbempChecKlist")
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
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                               


                                //Variables Validaciones detalle 1

                                rows["Benefi"].ToString().Length != 0 && rows["Benefi"].ToString() != null &&
                                rows["Enti"].ToString().Length != 0 && rows["Enti"].ToString() != null &&
                                rows["Sucur"].ToString().Length != 0 && rows["Sucur"].ToString() != null &&
                                rows["NumC"].ToString().Length != 0 && rows["NumC"].ToString() != null &&

                                rows["NumC2"].ToString().Length != 0 && rows["NumC2"].ToString() != null &&

                                //Variables Validaciones 
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null



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
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();
                                
                                //Variables validacion  crud Principal
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();



                                //Variables validacion DETALLE 1
                                string Benefi = rows["Benefi"].ToString();
                                string Enti = rows["Enti"].ToString();
                                string Sucur = rows["Sucur"].ToString();
                                string NumC = rows["NumC"].ToString();

                                string NumC2 = rows["NumC2"].ToString();




                                //LISTA DETALLE 1
                                List<string> crudPrincipal = new List<string>();
                                //LISTA DETALLE 1
                                List<string> crudDet1 = new List<string>() { Benefi, Enti, Sucur, NumC, NumC2 };


                                string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);

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
                                        //newFunctions_4.ScreenshotNuevo("Abrir Programa", true, file);
                                        ////VALIDACION CODIGO DEL PROGRAMA
                                        //errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////VALIDACION NOMBRE DEL PROGRAMA
                                        //errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////VERSION
                                        //FuncionesGlobales.GetVersion(desktopSession, file);
                                        //Thread.Sleep(1000);
                                        ////NOTAS
                                        //newFunctions_4.openInnerNote(desktopSession, file);

                                        //Thread.Sleep(1500);
                                       


                                        //MODIFICAMOS
                                        //AGREGAR REGISTRO Principal
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKNmAbemp.AgregarCrud(desktopSession, 0, crudPrincipal, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);
                                        Thread.Sleep(1500);
                                        CrudKNmAbemp.AgregarCrud(desktopSession, 1, crudPrincipal, file);
                                        newFunctions_4.ScreenshotNuevo("Editar Datos Crud Principal", true, file);
                                        Thread.Sleep(1500);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);
                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        // ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);

                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);
                                        CrudKNmAbemp.AgregarCrud(desktopSession, 1, crudPrincipal, file);

                                        newFunctions_4.ScreenshotNuevo("Editar Datos Crud Principal", true, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(3000);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);



                                        ////AGREGAR REGISTRO DETALLE 1                                        
                                        Dictionary<string, Point> botonesNavegador2 = new Dictionary<string, Point>();
                                        botonesNavegador2 = PruebaCRUD.findname(desktopSession, 25, 1);
                                        var ElementList2 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Insertar"].X, botonesNavegador2["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKNmAbemp.CrudDet(desktopSession, 0, crudDet1, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle", true, file);

                                        Thread.Sleep(1500);



                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar detalle", true, file);

                                        CrudKNmAbemp.CrudDet(desktopSession, 1, crudDet1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados detalle", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Cancelar"].X, botonesNavegador2["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados detalle", true, file);

                                        Thread.Sleep(1000);



                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar detalle", true, file);

                                        CrudKNmAbemp.CrudDet(desktopSession, 1, crudDet1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados detalle", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada detalle", true, file);








                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //REPORTE DINAMICO
                                        string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
                                        errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF DINAMICO
                                        listaResultado = Textopdf(pdf, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        ////REPORTE PRELIMINAR
                                        //string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        //errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        ////Rectificar Pie Pagina PDF PRELIMINAR
                                        //listaResultado = Textopdf(pdf1, codProgram, user);
                                        //listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        ////EXPORTAR EXCEL
                                        //string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        //errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_1.lapiz(desktopSession);


                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList2[1], botonesNavegador2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 1", true, file);
                                        ////VALIDACIÓN ELIMINAR REGISTRO
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle1, motor, user, CampoDetalle1, Codigo, "D", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, "", CampoDetalle1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        Thread.Sleep(2000);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Principal", true, file);
                                        ////VALIDACIÓN ELIMINAR REGISTRO
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "D", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Fin del programa", true, file);
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
        [TestMethod]
        public void KNmCaccfChecKlist()
        {
            //Debugger.Launch();
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_136.KNmCaccfChecKlist")
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
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&


                                //Variables Validaciones Crud principal

                                rows["Cod"].ToString().Length != 0 && rows["Cod"].ToString() != null &&
                                rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
                                rows["Numero"].ToString().Length != 0 && rows["Numero"].ToString() != null &&
                                


                                //Variables Validaciones detalle 1

                                rows["Posf"].ToString().Length != 0 && rows["Posf"].ToString() != null &&
                                rows["Tipo"].ToString().Length != 0 && rows["Tipo"].ToString() != null &&
                                rows["Jus"].ToString().Length != 0 && rows["Jus"].ToString() != null &&
                                rows["Rell"].ToString().Length != 0 && rows["Rell"].ToString() != null &&
                                rows["Posf2"].ToString().Length != 0 && rows["Posf2"].ToString() != null &&


                                //Variables Validaciones 
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null



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
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();

                                //Variables validacion 
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();

                                //Variables validacion  crud Principal
                                string Cod = rows["Cod"].ToString();
                                string Nombre = rows["Nombre"].ToString();
                                string Numero = rows["Numero"].ToString();


                                //Variables validacion DETALLE 1
                                string Posf = rows["Posf"].ToString();
                                string Tipo = rows["Tipo"].ToString();
                                string Jus = rows["Jus"].ToString();
                                string Rell = rows["Rell"].ToString();
                                string Posf2 = rows["Posf2"].ToString();



                                //LISTA DETALLE 1
                                List<string> crudPrincipal = new List<string>() { Cod, Nombre, Numero};
                                //LISTA DETALLE 1
                                List<string> crudDet1 = new List<string>() { Posf, Tipo, Jus, Rell, Posf2 };


                                string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);

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
                                        //newFunctions_4.ScreenshotNuevo("Abrir Programa", true, file);
                                        ////VALIDACION CODIGO DEL PROGRAMA
                                        //errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////VALIDACION NOMBRE DEL PROGRAMA
                                        //errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////VERSION
                                        //FuncionesGlobales.GetVersion(desktopSession, file);
                                        //Thread.Sleep(1000);
                                        ////NOTAS
                                        //newFunctions_4.openInnerNote(desktopSession, file);

                                        //Thread.Sleep(1500);

                                        //LAPIZ
                                        newFunctions_1.lapiz(desktopSession);

                                        //MODIFICAMOS
                                        //AGREGAR REGISTRO Principal
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKNmCaccf.AgregarCrud(desktopSession, 0, crudPrincipal, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);
                                        Thread.Sleep(1500);
                                        CrudKNmCaccf.AgregarCrud(desktopSession, 1, crudPrincipal, file);
                                        newFunctions_4.ScreenshotNuevo("Editar Datos Crud Principal", true, file);
                                        Thread.Sleep(1500);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);
                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        // ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);

                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);
                                        CrudKNmCaccf.AgregarCrud(desktopSession, 1, crudPrincipal, file);

                                        newFunctions_4.ScreenshotNuevo("Editar Datos Crud Principal", true, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(3000);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);



                                        ////AGREGAR REGISTRO DETALLE 1                                        
                                        Dictionary<string, Point> botonesNavegador2 = new Dictionary<string, Point>();
                                        botonesNavegador2 = PruebaCRUD.findname(desktopSession, 25, 1);
                                        var ElementList2 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Insertar"].X, botonesNavegador2["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKNmCaccf.CrudDet(desktopSession, 0, crudDet1, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle", true, file);

                                        Thread.Sleep(1500); 



                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar detalle", true, file);

                                        CrudKNmCaccf.CrudDet(desktopSession, 1, crudDet1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados detalle", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Cancelar"].X, botonesNavegador2["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados detalle", true, file);

                                        Thread.Sleep(1000);



                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar detalle", true, file);

                                        CrudKNmCaccf.CrudDet(desktopSession, 1, crudDet1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados detalle", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada detalle", true, file);








                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
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

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_1.lapiz(desktopSession);


                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList2[1], botonesNavegador2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 1", true, file);
                                        ////VALIDACIÓN ELIMINAR REGISTRO
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle1, motor, user, CampoDetalle1, Codigo, "D", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, "", CampoDetalle1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        Thread.Sleep(2000);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Principal", true, file);
                                        ////VALIDACIÓN ELIMINAR REGISTRO
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "D", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Fin del programa", true, file);
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


        [TestMethod]
        public void KNmHmdemChecklist()
        {
            //Debugger.Launch();
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_136.KNmHmdemChecklist")
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
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&


                                //Variables Validaciones Crud principal

                                rows["EmpreO"].ToString().Length != 0 && rows["EmpreO"].ToString() != null &&
                                rows["Id"].ToString().Length != 0 && rows["Id"].ToString() != null &&
                                rows["EmpreD"].ToString().Length != 0 && rows["EmpreD"].ToString() != null &&
                                rows["CenC"].ToString().Length != 0 && rows["CenC"].ToString() != null &&
                                rows["CenE"].ToString().Length != 0 && rows["CenE"].ToString() != null &&
                                rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
                                rows["NumC"].ToString().Length != 0 && rows["NumC"].ToString() != null &&
                                rows["Valor"].ToString().Length != 0 && rows["Valor"].ToString() != null &&
                                rows["Valor2"].ToString().Length != 0 && rows["Valor2"].ToString() != null &&






                                //Variables Validaciones 
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null



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
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();

                                //Variables validacion 
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();

                                //Variables validacion  crud Principal
                                string EmpreO = rows["EmpreO"].ToString();
                                string Id = rows["Id"].ToString();
                                string EmpreD = rows["EmpreD"].ToString();
                                string CenC = rows["CenC"].ToString();
                                string CenE = rows["CenE"].ToString();
                                string Fecha = rows["Fecha"].ToString();
                                string NumC = rows["NumC"].ToString();
                                string Valor = rows["Valor"].ToString();
                                string Valor2 = rows["Valor2"].ToString();





                                //LISTA DETALLE PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { EmpreO, Id, EmpreD, CenC, CenE, Fecha, NumC, Valor, Valor2 };
                               


                                string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);

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
                                        //newFunctions_4.ScreenshotNuevo("Abrir Programa", true, file);
                                        ////VALIDACION CODIGO DEL PROGRAMA
                                        //errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////VALIDACION NOMBRE DEL PROGRAMA
                                        //errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////VERSION
                                        //FuncionesGlobales.GetVersion(desktopSession, file);
                                        //Thread.Sleep(1000);
                                        ////NOTAS
                                        //newFunctions_4.openInnerNote(desktopSession, file);

                                        //Thread.Sleep(1500);



                                        ////MODIFICAMOS
                                        ////AGREGAR REGISTRO Principal
                                        ////Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        ////botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 0);
                                        ////var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        ////desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        ////desktopSession.Mouse.Click(null);

                                        ////CrudKNmHmdem.AgregarCrud(desktopSession, 0, crudPrincipal, file);
                                        ////newFunctions_4.Screenshot
                                        ///eMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        ////desktopSession.Mouse.Click(null);
                                        ////newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        ////Thread.Sleep(1500);

                                        //////DESCARTAR EDICION
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //newFunctions_4.ScreenshotNuevo("Editar", true, file);
                                        //Thread.Sleep(1500);

                                        //CrudKNmHmdem.AgregarCrud(desktopSession, 1, crudPrincipal, file);
                                        //newFunctions_4.ScreenshotNuevo("Editar Datos Crud Principal", true, file);
                                        //Thread.Sleep(1500);

                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        //Thread.Sleep(1000);
                                        ////LUPA
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(1000);

                                        //// ACEPTAR EDICION
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //Thread.Sleep(2000);

                                        //newFunctions_4.ScreenshotNuevo("Editar", true, file);
                                        //CrudKNmHmdem.AgregarCrud(desktopSession, 1, crudPrincipal, file);

                                        //newFunctions_4.ScreenshotNuevo("Editar Datos Crud Principal", true, file);
                                        //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //Thread.Sleep(3000);
                                        //newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);








                                        ////LAPIZ
                                        //newFunctions_1.lapiz(desktopSession);



                                        ////QBE
                                        //errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////REPORTE DINAMICO
                                        //string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
                                        //errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        ////Rectificar Pie Pagina PDF DINAMICO
                                        //listaResultado = Textopdf(pdf, codProgram, user);
                                        //listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        ////REPORTE PRELIMINAR
                                        //string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        //errors = CrudKNmHmdem.ReportePreliminar1(desktopSession, BanderaPreli, file, pdf1);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        ////Rectificar Pie Pagina PDF PRELIMINAR
                                        //listaResultado = Textopdf(pdf1, codProgram, user);
                                        //listaResultado.ForEach(e => ErrorValidacion.Add(e));
                                        //Thread.Sleep(1500);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //REPORTE PRELIMINAR 2 
                                        string pdf2 = @"C:\Reportes\ReportePDF2_" + "Preliminar" + "_" + Hora();
                                        errors = CrudKNmHmdem.ReportePreliminar2(desktopSession, BanderaPreli, file,codProgram, pdf2);
                                        Thread.Sleep(2000);


                                        newFunctions_4.ScreenshotNuevo("Ventana Exportar Excel", true, file);
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        CrudKNmHmdem.ExpExel(desktopSession, ruta);
                                        Thread.Sleep(2000);




                                        ////EXPORTAR EXCEL
                                        //string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        //errors = newFunctions_2.generarExcel(desktopSession, file, codProgram, ruta);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        ////Rectificar Pie Pagina PDF PRELIMINAR
                                        //listaResultado = Textopdf(pdf2, codProgram, user);
                                        //listaResultado.ForEach(e => ErrorValidacion.Add(e));
                                        //Thread.Sleep(1500);

                                        //EXPORTAR EXCEL
                                        string ruta2 = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta2);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_1.lapiz(desktopSession);


                                        
                                        ////VALIDACIÓN ELIMINAR REGISTRO
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle1, motor, user, CampoDetalle1, Codigo, "D", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Codigo, "", CampoDetalle1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        //newFunctions_4.ScreenshotNuevo("Datos Eliminados Principal", true, file);
                                        ////VALIDACIÓN ELIMINAR REGISTRO
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "D", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Fin del programa", true, file);
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

        [TestMethod]
        public void AbrirPrograma()
        {

            string codProgram = "KNmLinom";
            string user = "ONATALIA";
            string motor = "ORA";

            AbrirPrograma a = new AbrirPrograma(codProgram, user);
            desktopSession = a.Start2(motor, "");


        }
        //[TestCleanup]
        //public void Limpiar()
        //{
        //    AbrirPrograma a = new AbrirPrograma();
        //    if (desktopSession != null)
        //    {
        //        desktopSession.Close();
        //        desktopSession.Dispose();
        //    }
        //    a.Stop();
        //}

    }
 }

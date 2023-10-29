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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_9;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_10;
using OpenQA.Selenium.Appium;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosEd;

namespace Ophelia_Test_V21.TestMethods.ModulosSO
{
    [TestClass]
    public class ModulosSO_16 : FuncionesVitales
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

        public ModulosSO_16()
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

        //[TestMethod]
        //public void KSoGiateChecKlist()
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

        //        if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_15.KSoGiateChecKlist")
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

        //                        //Variables Validaciones
        //                        rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
        //                        rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
        //                        rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&

        //                        //Datos Cruds
        //                        rows["Codigo"].ToString().Length != 0 && rows["Codigo"].ToString() != null &&
        //                        rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
        //                        rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
        //                        rows["ActNombre"].ToString().Length != 0 && rows["ActNombre"].ToString() != null)

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

        //                        //Variables validacion
        //                        string Tabla = rows["Tabla"].ToString();
        //                        string Campo = rows["Campo"].ToString();
        //                        string EditCampo = rows["EditCampo"].ToString();

        //                        //Variables crud principal
        //                        string Codigo = rows["Codigo"].ToString();
        //                        string Nombre = rows["Nombre"].ToString();
        //                        string Fecha = rows["Fecha"].ToString();
        //                        string ActNombre = rows["ActNombre"].ToString();

        //                        //LISTA CRUD PRINCIPAL
        //                        List<string> crudPrincipal = new List<string>() { Codigo, Nombre, Fecha, ActNombre };

        //                        //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
        //                        string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);

        //                        try
        //                        {
        //                            List<string> ErrorValidacion = new List<string>();
        //                            List<string> errors = new List<string>();

        //                            AbrirPrograma a = new AbrirPrograma(codProgram, user);
        //                            desktopSession = a.Start2(motor, "");
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
        //                                Thread.Sleep(1000);

        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, Campo, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

        //                                //MODIFICAMOS
        //                                ////AGREGAR REGISTRO Principal
        //                                Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
        //                                botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
        //                                var ElementList = PruebaCRUD.NavClass(desktopSession);
        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
        //                                desktopSession.Mouse.Click(null);

        //                                CrudKsogiate.AgregarCrud(desktopSession, 0, crudPrincipal);
        //                                newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

        //                                Thread.Sleep(1500);

        //                                ////Validacion Agregar
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

        //                                ////DESCARTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                CrudKsogiate.AgregarCrud(desktopSession, 1, crudPrincipal);
        //                                newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

        //                                Thread.Sleep(1000);

        //                                ////Validacion Editar Descartar
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

        //                                ////ACEPTAR EDICION
        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                CrudKsogiate.AgregarCrud(desktopSession, 1, crudPrincipal);
        //                                newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
        //                                desktopSession.Mouse.Click(null);
        //                                newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

        //                                Thread.Sleep(3000);
        //                                ////validacion error al editar
        //                                bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
        //                                //Debugger.Launch();
        //                                if (valEditOra == true)
        //                                {
        //                                    ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
        //                                    //LUPA
        //                                    errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                    newFunctions_3.LupaAud(desktopSession, "1", file);
        //                                    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                    //ACEPTAR EDICION
        //                                    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
        //                                    desktopSession.Mouse.Click(null);
        //                                    newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //                                    CrudKsogiate.AgregarCrud(desktopSession, 1, crudPrincipal);
        //                                    newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //                                    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
        //                                    desktopSession.Mouse.Click(null);
        //                                    Thread.Sleep(3000);
        //                                }
        //                                newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

        //                                //////Validacion Editar
        //                                Thread.Sleep(1500);
        //                                ////// //Debugger.Launch();
        //                                ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "E", c, ErrorValidacion);
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, ActNombre, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

        //                                Thread.Sleep(2000);

        //                                //QBE
        //                                errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
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

        //                                Thread.Sleep(2000);

        //                                //LUPA
        //                                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
        //                                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

        //                                PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
        //                                newFunctions_4.ScreenshotNuevo("Datos Eliminados Principal", true, file);
        //                                //VALIDACIÓN ELIMINAR REGISTRO
        //                                ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "D", c, ErrorValidacion);
        //                                ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

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
        public void KSoGiateChecKlist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_16.KSoGiateChecKlist")
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

                                //Variables Validaciones principal    
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&

                                //Variables Validaciones Detalle   @TablaDet @CampoDet @EditCampoDet
                                rows["TablaDet"].ToString().Length != 0 && rows["TablaDet"].ToString() != null &&
                                rows["CampoDet"].ToString().Length != 0 && rows["CampoDet"].ToString() != null &&
                                rows["EditCampoDet"].ToString().Length != 0 && rows["EditCampoDet"].ToString() != null &&



                                //Datos Crud Detalle  @Diag @ActDiag
                                rows["Diag"].ToString().Length != 0 && rows["Diag"].ToString() != null &&
                                rows["ActDiag"].ToString().Length != 0 && rows["ActDiag"].ToString() != null &&


                                //Datos Crud Principal   @Fecha @ActFecha  ide
                                rows["ide"].ToString().Length != 0 && rows["ide"].ToString() != null &&
                                rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&                                
                                rows["ActFecha"].ToString().Length != 0 && rows["ActFecha"].ToString() != null )

                                


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
                                string EditCampo = rows["EditCampo"].ToString();


                                //Variables validacion  Detalle     @TablaDet @CampoDet @EditCampoDet
                                string TablaDet = rows["TablaDet"].ToString();
                                string CampoDet = rows["CampoDet"].ToString();
                                string EditCampoDet = rows["EditCampoDet"].ToString();


                                //Variables crud principal   @Fecha @ActFecha
                                string ide = rows["ide"].ToString();
                                string Fecha = rows["Fecha"].ToString();
                                string ActFecha = rows["ActFecha"].ToString();


                                //Variables Detalle   @Diag @ActDiag
                                string Diag = rows["Diag"].ToString();
                                string ActDiag = rows["ActDiag"].ToString();


                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { ide, Fecha, ActFecha };
  
                                List<string> crudDet1 = new List<string>() { Diag, ActDiag};


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
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


                                        // Botones Crud PRINCIPAL
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 28, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);


                                        //MODIFICAMOS
                                        //AGREGAR REGISTRO Principal
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);

                                        newFunctions_4.ScreenshotNuevo(" Datos CRUD PRINCIPAL  ", true, file);
                                        CrudKsogiate.AggCrudPri(desktopSession, 0, crudPrincipal);
                                        Thread.Sleep(1500);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);
                                        Thread.Sleep(1500);


                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        CrudKsogiate.ChoiceVentana(desktopSession);
                                        Thread.Sleep(1500);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKsogiate.AggCrudPri(desktopSession, 1, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);
                                        Thread.Sleep(1000);


                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);


                                        CrudKsogiate.AggCrudPri(desktopSession, 1, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);


                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);
                                        Thread.Sleep(2000);


                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        //// CRUD DETALLE 
                                        Dictionary<string, Point> botonesNavegador2 = new Dictionary<string, Point>();
                                        botonesNavegador2 = PruebaCRUD.findname(desktopSession, 28, 1);
                                        var ElementList2 = PruebaCRUD.NavClass(desktopSession);


                                        ////AGREGAR REGISTRO DETALLE 1                        
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Insertar"].X, botonesNavegador2["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKsogiate.AggCrudDet(desktopSession, 0, crudDet1);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados CRUD DETALLE", true, file);
                                        Thread.Sleep(1500);


                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        ///DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar CRUD DETALLE", true, file);

                                        CrudKsogiate.AggCrudDet(desktopSession, 1, crudDet1);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Cancelar"].X, botonesNavegador2["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados CRUD DETALLE", true, file);
                                        Thread.Sleep(1000);


                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        /////////////////////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar CRUD DETALLE", true, file);

                                        CrudKsogiate.AggCrudDet(desktopSession, 1, crudDet1);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada CRUD DETALLE", true, file);
                                        Thread.Sleep(3000);


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


                                        //REPORTE PRELIMINAR
                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(2000);


                                        newFunctions_1.lapiz(desktopSession);
                                        Thread.Sleep(2000);


                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList2[1], botonesNavegador2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados CRUD DETALLE", true, file);
                                        Thread.Sleep(1500);


                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Principal", true, file);
                                        Thread.Sleep(1500);


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
        public void KSoRiincChecKlist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_16.KSoRiincChecKlist")
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
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null)

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

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
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

                                        //MODIFICAMOS                                     
                                        CrudKsoriinc.AgregarCrud(desktopSession);
                                        newFunctions_4.ScreenshotNuevo("Fechas", true, file);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        
                                        CrudKsoriinc.Boton(desktopSession);
                                        newFunctions_4.ScreenshotNuevo("Boton Aceptar", true, file);
                                        Thread.Sleep(2000);

                                        newFunctions_4.ScreenshotNuevo("Ventana Exportar Excel", true, file);
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        CrudKsoriinc.ExpExel(desktopSession, ruta);
                                        Thread.Sleep(2000);

                                        newFunctions_4.ScreenshotNuevo("Cerrar Ventana", true, file);
                                        CrudKsoriinc.VentanaClose(desktopSession);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Programa finalizado", true, file);
                                        Thread.Sleep(2000);
                                  
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
            string programa = "kgedoepr";
            AbrirPrograma a = new AbrirPrograma(programa, "UNatalia"); //KFABILU
            desktopSession = a.Start2("SQL", "");
            Thread.Sleep(2000);
        }


        //[TestMethod]
        //public void AbrirPrograma()
        //{
        //    //VARIABLES
        //    string Rol = "502";
        //    string DescripcionRol = "prueba";


        //    //LISTA CRUD PRINCIPAL
        //    List<string> crudPrincipal = new List<string>() { Rol, DescripcionRol };

        //    string programa = "KSoRcaus";
        //    //Debugger.Launch();
        //    //AbrirPrograma a = new AbrirPrograma(programa, "UNatalia");
        //    //desktopSession = a.Start2("SQL", "");
        //    //AbrirPrograma a = new AbrirPrograma(programa, "ONATALIA");
        //    //desktopSession = a.Start2("ORA", "");
        //    AbrirPrograma a = new AbrirPrograma(programa, "UCalida1");
        //    desktopSession = a.Start2("SQL", "");
        //    //AbrirPrograma a = new AbrirPrograma(programa, "ODESAR");
        //    //desktopSession = a.Start2("ORA", "");
        //    //Thread.Sleep(2000000);
        //    List<string> crudData = new List<string>() { };
        //    string file = CrearDocumentoWordDinamico("prueba", "Pruebas", "SQL");

        //    ////AGREGAR REGISTRO
        //    Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
        //    botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
        //    var ElementList = PruebaCRUD.NavClass(desktopSession);
        //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
        //    desktopSession.Mouse.Click(null);

        //    ClasePrueba.AgregarCrud(desktopSession, 0, crudPrincipal);
        //    //Thread.Sleep(2000000);

        //    //newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

        //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
        //    desktopSession.Mouse.Click(null);
        //    //newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

        //    //Thread.Sleep(1500);

        //    //////Validacion Agregar
        //    ////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, RolNum, RolNum, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

        //    //////DESCARTAR EDICION
        //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
        //    desktopSession.Mouse.Click(null);
        //    //newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //    ClasePrueba.AgregarCrud(desktopSession, 1, crudPrincipal);

        //    //CrudKed3crol.CRUD(desktopSession, 2, crudVariables, file);
        //    //newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
        //    desktopSession.Mouse.Click(null);
        //    //newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

        //    //Thread.Sleep(1000);

        //    //////Validacion Editar Descartar
        //    //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, RolNum, RolPal, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

        //    //////ACEPTAR EDICION
        //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
        //    desktopSession.Mouse.Click(null);
        //    //newFunctions_4.ScreenshotNuevo("Editar", true, file);

        //    ClasePrueba.AgregarCrud(desktopSession, 1, crudPrincipal);

        //    //CrudKed3crol.CRUD(desktopSession, 2, crudVariables, file);
        //    ////newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

        //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
        //    desktopSession.Mouse.Click(null);
        //    //newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

        //    newFunctions_3.LupaAud(desktopSession, "1", file);

        //    //ELIMINAR REGISTRO CRUD

        //    newFunctions_3.LupaAud(desktopSession, "1", file);

        //    PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);


        //    //newFunctions_4.ScreenshotNuevo("Datos Eliminados", true, file);
        //    //VALIDACIÓN ELIMINAR REGISTRO
        //    //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla1, motor, user, Campo1, Codigo, "D", c, ErrorValidacion);
        //    //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Codigo, "", Campo1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

        //}

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

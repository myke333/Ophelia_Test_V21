using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosEd;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.ValidacionQueries;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Threading;


namespace Ophelia_Test_V21.TestMethods.ModulosED
{
    /// <summary>
    /// Descripción resumida de ModulosED_6
    /// </summary>  
    [TestClass]
    public class ModulosED_6 : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        protected static WindowsDriver<WindowsElement> desktopSession;

        //protected static WindowsDriver<WindowsElement> rootSession;
        //List<string> listaResultado;
        //List<string> Celda = new List<string>();

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
        public ModulosED_6()
        { }
        private TestContext testContextInstance;

        public TestContext TestContext
        {
            get
            { return testContextInstance; }
            set
            { testContextInstance = value; }
        }

        [TestMethod]
        public void KEd3CrolChecklist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosED.ModulosED_6.KEd3CrolChecklist")
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

                                //Variables Validaciones
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&

                                //Variables Validaciones detalle 1
                                rows["TablaDetalle1"].ToString().Length != 0 && rows["TablaDetalle1"].ToString() != null &&
                                rows["CampoDetalle1"].ToString().Length != 0 && rows["CampoDetalle1"].ToString() != null &&
                                rows["EditCampoDetalle1"].ToString().Length != 0 && rows["EditCampoDetalle1"].ToString() != null &&

                                //Variables Validaciones detalle 2
                                rows["TablaDetalle2"].ToString().Length != 0 && rows["TablaDetalle2"].ToString() != null &&
                                rows["CampoDetalle2"].ToString().Length != 0 && rows["CampoDetalle2"].ToString() != null &&
                                rows["EditCampoDetalle2"].ToString().Length != 0 && rows["EditCampoDetalle2"].ToString() != null &&

                                //Variables Validaciones detalle 3
                                rows["TablaDetalle3"].ToString().Length != 0 && rows["TablaDetalle3"].ToString() != null &&
                                rows["CampoDetalle3"].ToString().Length != 0 && rows["CampoDetalle3"].ToString() != null &&
                                rows["EditCampoDetalle3"].ToString().Length != 0 && rows["EditCampoDetalle3"].ToString() != null &&

                                //Variables Validaciones detalle 4
                                rows["TablaDetalle4"].ToString().Length != 0 && rows["TablaDetalle4"].ToString() != null &&
                                rows["CampoDetalle4"].ToString().Length != 0 && rows["CampoDetalle4"].ToString() != null &&
                                rows["EditCampoDetalle4"].ToString().Length != 0 && rows["EditCampoDetalle4"].ToString() != null &&

                                //Datos Cruds
                                rows["Rol"].ToString().Length != 0 && rows["Rol"].ToString() != null &&
                                rows["DescripcionRol"].ToString().Length != 0 && rows["DescripcionRol"].ToString() != null &&
                                rows["ActDescripcionRol"].ToString().Length != 0 && rows["ActDescripcionRol"].ToString() != null &&

                                rows["CargoDetalle1"].ToString().Length != 0 && rows["CargoDetalle1"].ToString() != null &&
                                rows["ActCargoDetalle1"].ToString().Length != 0 && rows["ActCargoDetalle1"].ToString() != null &&

                                rows["CargoSiguienteDet2"].ToString().Length != 0 && rows["CargoSiguienteDet2"].ToString() != null &&
                                rows["ActCargoSiguienteDet2"].ToString().Length != 0 && rows["ActCargoSiguienteDet2"].ToString() != null &&

                                rows["CentroCostoDet3"].ToString().Length != 0 && rows["CentroCostoDet3"].ToString() != null &&
                                rows["ActCargoSiguienteDet2"].ToString().Length != 0 && rows["ActCargoSiguienteDet2"].ToString() != null &&

                                rows["RolesPlantaDet4"].ToString().Length != 0 && rows["RolesPlantaDet4"].ToString() != null &&
                                rows["ActRolesPlantaDet4"].ToString().Length != 0 && rows["ActRolesPlantaDet4"].ToString() != null)
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

                                //Variables validacion DETALLE 1
                                string TablaDetalle1 = rows["TablaDetalle1"].ToString();
                                string CampoDetalle1 = rows["CampoDetalle1"].ToString();
                                string EditCampoDetalle1 = rows["EditCampoDetalle1"].ToString();

                                //Variables validacion DETALLE 2
                                string TablaDetalle2 = rows["TablaDetalle2"].ToString();
                                string CampoDetalle2 = rows["CampoDetalle2"].ToString();
                                string EditCampoDetalle2 = rows["EditCampoDetalle2"].ToString();

                                //Variables validacion DETALLE 3
                                string TablaDetalle3 = rows["TablaDetalle3"].ToString();
                                string CampoDetalle3 = rows["CampoDetalle3"].ToString();
                                string EditCampoDetalle3 = rows["EditCampoDetalle3"].ToString();

                                //Variables validacion DETALLE 4
                                string TablaDetalle4 = rows["TablaDetalle4"].ToString();
                                string CampoDetalle4 = rows["CampoDetalle4"].ToString();
                                string EditCampoDetalle4 = rows["EditCampoDetalle4"].ToString();

                                //Variables crud principal
                                string Rol = rows["Rol"].ToString();
                                string DescripcionRol = rows["DescripcionRol"].ToString();
                                string ActDescripcionRol = rows["ActDescripcionRol"].ToString();

                                //Variables Crud Det 1
                                string CargoDetalle1 = rows["CargoDetalle1"].ToString();
                                string ActCargoDetalle1 = rows["ActCargoDetalle1"].ToString();

                                //Variables Crud Det 2
                                string CargoSiguienteDet2 = rows["CargoSiguienteDet2"].ToString();
                                string ActCargoSiguienteDet2 = rows["ActCargoSiguienteDet2"].ToString();

                                //Variables Crud Det 3
                                string CentroCostoDet3 = rows["CentroCostoDet3"].ToString();
                                string ActCentroCostoDet3 = rows["ActCentroCostoDet3"].ToString();

                                //Variables Crud Det 4
                                string RolesPlantaDet4 = rows["RolesPlantaDet4"].ToString();
                                string ActRolesPlantaDet4 = rows["ActRolesPlantaDet4"].ToString();

                                //string Tabla = rows["Tabla"].ToString();
                                //string Campo = rows["Campo"].ToString();


                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { Rol, DescripcionRol, ActDescripcionRol };
                                //LISTA DETALLE 1
                                List<string> crudDet1 = new List<string>() { CargoDetalle1, ActCargoDetalle1 };
                                //LISTA DETALLE 2
                                List<string> crudDet2 = new List<string>() { CargoSiguienteDet2, ActCargoSiguienteDet2 };
                                //LISTA DETALLE 3
                                List<string> crudDet3 = new List<string>() { CentroCostoDet3, ActCentroCostoDet3 };
                                //LISTA DETALLE 4
                                List<string> crudDet4 = new List<string>() { RolesPlantaDet4, ActRolesPlantaDet4 };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "ED", motor);

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
                                        ////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        ClasePrueba.AgregarCrud(desktopSession, 0, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);

                                        //////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Rol, Rol, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        ClasePrueba.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        //////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Rol, DescripcionRol, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        //////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        ClasePrueba.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        Thread.Sleep(1000);
                                        //validacion error al editar
                                        bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                            //////ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                            ClasePrueba.AgregarCrud(desktopSession, 1, crudPrincipal);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        //////Validacion Editar
                                        Thread.Sleep(1500);
                                        //// //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Rol, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Rol, ActDescripcionRol, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1); 

                                         Thread.Sleep(1000);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////AGREGAR REGISTRO DETALLE 1
                                        Dictionary<string, Point> botonesNavegadorD1 = new Dictionary<string, Point>();
                                        botonesNavegadorD1 = PruebaCRUD.findname(desktopSession, 30, 2);
                                        var ElementListD1 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementListD1[2].Coordinates, botonesNavegadorD1["Insertar"].X, botonesNavegadorD1["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        ClasePrueba.AgregarCrud2(desktopSession, 0, crudDet1);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementListD1[2].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegadorD1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 1", true, file);

                                        Thread.Sleep(1500);

                                        //////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Rol, Rol, CampoDetalle1, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementListD1[2].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegadorD1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                        ClasePrueba.AgregarCrud2(desktopSession, 1, crudDet1);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementListD1[2].Coordinates, botonesNavegadorD1["Cancelar"].X, botonesNavegadorD1["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 1", true, file);

                                        Thread.Sleep(1000);

                                        //////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Rol, CargoDetalle1, EditCampoDetalle1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        //////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementListD1[2].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegadorD1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                        ClasePrueba.AgregarCrud2(desktopSession, 1, crudDet1);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementListD1[2].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegadorD1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        Thread.Sleep(1000);
                                        //validacion error al editar
                                        bool valEditOraDet1 = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOraDet1 == true)
                                        {
                                            ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //////ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementListD1[2].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegadorD1["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                            ClasePrueba.AgregarCrud2(desktopSession, 1, crudDet1);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                            desktopSession.Mouse.MouseMove(ElementListD1[2].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegadorD1["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 1", true, file);

                                        //////Validacion Editar
                                        Thread.Sleep(1500);
                                        //// //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle1, motor, user, CampoDetalle1, Rol, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Rol, EditCampoDetalle1, EditCampoDetalle1, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                         Thread.Sleep(1000);

                                        ////AGREGAR REGISTRO DETALLE 2
                                        Dictionary<string, Point> botonesNavegadorD2 = new Dictionary<string, Point>();
                                        botonesNavegadorD2 = PruebaCRUD.findname(desktopSession, 30, 1);
                                        var ElementListD2 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementListD2[1].Coordinates, botonesNavegadorD2["Insertar"].X, botonesNavegadorD2["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        ClasePrueba.AgregarCrud3(desktopSession, 0, crudDet2);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementListD2[1].Coordinates, botonesNavegadorD2["Aplicar"].X, botonesNavegadorD2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 2", true, file);

                                        Thread.Sleep(1500);

                                        //Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle2, user, motor, CampoDetalle2, Rol, Rol, CampoDetalle2, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementListD2[1].Coordinates, botonesNavegadorD2["Editar"].X, botonesNavegadorD2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);

                                        ClasePrueba.AgregarCrud3(desktopSession, 1, crudDet2);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementListD2[1].Coordinates, botonesNavegadorD2["Cancelar"].X, botonesNavegadorD2["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 2", true, file);

                                        Thread.Sleep(2000);

                                        //Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle2, user, motor, CampoDetalle2, Rol, CargoSiguienteDet2, EditCampoDetalle2, c, ErrorValidacion, "No se edito correctamente", 1);

                                        Thread.Sleep(2000);

                                        //////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementListD2[1].Coordinates, botonesNavegadorD2["Editar"].X, botonesNavegadorD2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);

                                        ClasePrueba.AgregarCrud3(desktopSession, 1, crudDet2);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementListD2[1].Coordinates, botonesNavegadorD2["Aplicar"].X, botonesNavegadorD2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        Thread.Sleep(1000);
                                        //validacion error al editar
                                        bool valEditOraDet2 = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOraDet2 == true)
                                        {
                                            ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //////ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementListD2[1].Coordinates, botonesNavegadorD2["Editar"].X, botonesNavegadorD2["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);

                                            ClasePrueba.AgregarCrud3(desktopSession, 1, crudDet2);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                            desktopSession.Mouse.MouseMove(ElementListD2[1].Coordinates, botonesNavegadorD2["Aplicar"].X, botonesNavegadorD2["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 2", true, file);

                                        //Validacion Editar
                                        Thread.Sleep(1500);
                                        //// //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle2, motor, user, CampoDetalle2, Rol, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle2, user, motor, CampoDetalle2, Rol, EditCampoDetalle2, EditCampoDetalle2, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(2000);

                                        ClasePrueba.Clickbutton(desktopSession, 1);

                                        Thread.Sleep(2000);

                                        ////AGREGAR REGISTRO DETALLE 3
                                        Dictionary<string, Point> botonesNavegadorD3 = new Dictionary<string, Point>();
                                        botonesNavegadorD3 = PruebaCRUD.findname(desktopSession, 30, 1);
                                        var ElementListD3 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementListD3[1].Coordinates, botonesNavegadorD3["Insertar"].X, botonesNavegadorD3["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        ClasePrueba.AgregarCrud4(desktopSession, 0, crudDet3);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementListD3[1].Coordinates, botonesNavegadorD3["Aplicar"].X, botonesNavegadorD3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 3", true, file);

                                        Thread.Sleep(1500);

                                        //Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle3, user, motor, CampoDetalle3, Rol, Rol, CampoDetalle3, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementListD3[1].Coordinates, botonesNavegadorD3["Editar"].X, botonesNavegadorD3["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 3", true, file);

                                        ClasePrueba.AgregarCrud4(desktopSession, 1, crudDet3);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementListD3[1].Coordinates, botonesNavegadorD3["Cancelar"].X, botonesNavegadorD3["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 3", true, file);

                                        Thread.Sleep(1000);

                                        //Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle3, user, motor, CampoDetalle3, Rol, CentroCostoDet3, EditCampoDetalle3, c, ErrorValidacion, "No se edito correctamente", 1);

                                        //////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementListD3[1].Coordinates, botonesNavegadorD3["Editar"].X, botonesNavegadorD3["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 3", true, file);

                                        ClasePrueba.AgregarCrud4(desktopSession, 1, crudDet3);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementListD3[1].Coordinates, botonesNavegadorD3["Aplicar"].X, botonesNavegadorD3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        Thread.Sleep(1000);
                                        //validacion error al editar
                                        bool valEditOraDet3 = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOraDet3 == true)
                                        {
                                            ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //////ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementListD3[1].Coordinates, botonesNavegadorD3["Editar"].X, botonesNavegadorD3["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar Detalle 3", true, file);

                                            ClasePrueba.AgregarCrud4(desktopSession, 1, crudDet3);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 3", true, file);

                                            desktopSession.Mouse.MouseMove(ElementListD3[1].Coordinates, botonesNavegadorD3["Aplicar"].X, botonesNavegadorD3["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        //Validacion Editar
                                        Thread.Sleep(1500);
                                        //// //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle3, motor, user, CampoDetalle3, Rol, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle3, user, motor, CampoDetalle3, Rol, EditCampoDetalle3, EditCampoDetalle3, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //Thread.Sleep(1000);

                                        ClasePrueba.Clickbutton(desktopSession, 2);

                                        ////AGREGAR REGISTRO DETALLE 4
                                        Dictionary<string, Point> botonesNavegadorD4 = new Dictionary<string, Point>();
                                        botonesNavegadorD4 = PruebaCRUD.findname(desktopSession, 30, 1);
                                        var ElementListD4 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementListD4[1].Coordinates, botonesNavegadorD4["Insertar"].X, botonesNavegadorD4["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        ClasePrueba.AgregarCrud5(desktopSession, 0, crudDet4);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 4", true, file);

                                        desktopSession.Mouse.MouseMove(ElementListD4[1].Coordinates, botonesNavegadorD4["Aplicar"].X, botonesNavegadorD4["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 4", true, file);

                                        Thread.Sleep(1500);

                                        //Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle4, user, motor, CampoDetalle4, Rol, Rol, CampoDetalle4, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementListD4[1].Coordinates, botonesNavegadorD4["Editar"].X, botonesNavegadorD4["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 4", true, file);

                                        ClasePrueba.AgregarCrud5(desktopSession, 1, crudDet4);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 4", true, file);

                                        desktopSession.Mouse.MouseMove(ElementListD4[1].Coordinates, botonesNavegadorD4["Cancelar"].X, botonesNavegadorD4["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 4", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle4, user, motor, CampoDetalle4, Rol, RolesPlantaDet4, EditCampoDetalle4, c, ErrorValidacion, "No se edito correctamente", 1);

                                        //////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementListD4[1].Coordinates, botonesNavegadorD4["Editar"].X, botonesNavegadorD4["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 4", true, file);

                                        ClasePrueba.AgregarCrud5(desktopSession, 1, crudDet4);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 4", true, file);

                                        desktopSession.Mouse.MouseMove(ElementListD4[1].Coordinates, botonesNavegadorD4["Aplicar"].X, botonesNavegadorD4["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        Thread.Sleep(1000);
                                        //validacion error al editar
                                        bool valEditOraDet4 = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOraDet4 == true)
                                        {
                                            ErrorValidacion.Add("El programa no deja editar el registro sin antes llamar a la lupa de auditoria");
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //////ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementListD4[1].Coordinates, botonesNavegadorD4["Editar"].X, botonesNavegadorD4["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar Detalle 4", true, file);

                                            ClasePrueba.AgregarCrud5(desktopSession, 1, crudDet4);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 4", true, file);

                                            desktopSession.Mouse.MouseMove(ElementListD4[1].Coordinates, botonesNavegadorD4["Aplicar"].X, botonesNavegadorD4["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        //////Validacion Editar
                                        Thread.Sleep(1500);
                                        //// //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle4, motor, user, CampoDetalle4, Rol, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle4, user, motor, CampoDetalle1, Rol, EditCampoDetalle4, EditCampoDetalle4, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //Thread.Sleep(1000);

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
                                        /*ring COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, tipo, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));*/

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR REGISTRO CRUD
                                        ClasePrueba.Clickbutton(desktopSession, 2);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementListD4[1], botonesNavegadorD4, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 4", true, file);
                                        //VALIDACIÓN ELIMINAR REGISTRO
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle4, motor, user, CampoDetalle4, Rol, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle4, user, motor, CampoDetalle4, Rol, "", CampoDetalle4, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);


                                        ClasePrueba.Clickbutton(desktopSession, 1);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementListD3[1], botonesNavegadorD3, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 3", true, file);
                                        //VALIDACIÓN ELIMINAR REGISTRO
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle3, motor, user, CampoDetalle3, Rol, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle3, user, motor, CampoDetalle3, Rol, "", CampoDetalle3, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);


                                        ClasePrueba.Clickbutton(desktopSession, 0);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementListD2[1], botonesNavegadorD2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 2", true, file);
                                        //VALIDACIÓN ELIMINAR REGISTRO
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle2, motor, user, CampoDetalle2, Rol, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle2, user, motor, CampoDetalle2, Rol, "", CampoDetalle2, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementListD1[2], botonesNavegadorD1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados Detalle 1", true, file);
                                        //VALIDACIÓN ELIMINAR REGISTRO
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaDetalle1, motor, user, CampoDetalle1, Rol, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaDetalle1, user, motor, CampoDetalle1, Rol, "", CampoDetalle1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Eliminados", true, file);
                                        //VALIDACIÓN ELIMINAR REGISTRO
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Rol, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Rol, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);


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
        public void KEd3PilcChecKlist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosED.ModulosED_6.KEd3PilcChecKlist")
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

                                //Variables Validaciones
                                rows["Lupa"].ToString().Length != 0 && rows["Lupa"].ToString() != null )

                            //rows["varqbe"].ToString().Length != 0 && rows["varqbe"].ToString() != null)


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
                                string Lupa = rows["Lupa"].ToString();
                                //string varqbe = rows["varqbe"].ToString();

                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { Lupa, /*varqbe*/ };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "ED", motor);

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

                                        ////VERSION
                                        //FuncionesGlobales.GetVersion(desktopSession, file);
                                        //Thread.Sleep(2500);
                                        ////NOTAS
                                        //newFunctions_4.openInnerNote(desktopSession, file);
                                        //Thread.Sleep(1500);

                                        newFunctions_4.ScreenshotNuevo("lupa ", true, file);
                                        CrudEd3pilc.LupaDinamica(desktopSession, crudPrincipal);
                                        Thread.Sleep(1500);


                                        ////QBE2
                                        CrudEd3pilc.qbe2(desktopSession, crudPrincipal);
                                        newFunctions_4.ScreenshotNuevo(" QBE", true, file);
                                        Thread.Sleep(500);

                                        CrudEd3pilc.Enter(desktopSession);
                                        Thread.Sleep(700);


                                        CrudEd3pilc.Enter(desktopSession);
                                        newFunctions_4.ScreenshotNuevo("Registros QBE", true, file);
                                        Thread.Sleep(1000);


                                        ////QBE
                                        //errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(2000);

                                        // Click Aceptar
                                        newFunctions_4.ScreenshotNuevo("Boton aceptar", true, file);
                                        CrudEd3pilc.ClickBtnAceptar(desktopSession, crudPrincipal);


                                        newFunctions_4.ScreenshotNuevo("Pantalla 1 ", true, file);
                                        CrudEd3pilc.Enter(desktopSession);


                                        newFunctions_4.ScreenshotNuevo("Pantalla 2 ", true, file);
                                        Thread.Sleep(300);
                                        CrudEd3pilc.Enter(desktopSession);


                                        CrudEd3pilc.Enter(desktopSession);
                                        Thread.Sleep(300);
                                        newFunctions_4.ScreenshotNuevo("Pantalla 3 ", true, file);

                                        // Click Errores
                                        CrudEd3pilc.ClickBtnErrores(desktopSession);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Boton Errores", true, file);
                                        Thread.Sleep(500);

                                        // PDF reporte errores
                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                                        newFunctions_4.ScreenshotNuevo("Programa Finalizado", true, file);
                                        Thread.Sleep(1000);


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
        public void KEd3PccoChecKlist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosED.ModulosED_6.KEd3PccoChecKlist")
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
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null )

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
                                string file = CrearDocumentoWordDinamico(codProgram, "ED", motor);

                                string campo = "1";

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

                                        //MODIFICAMOS Fecha Inicial para traer datos.
                                        CrudKed3pcco.fecha(desktopSession);
                                        newFunctions_4.ScreenshotNuevo("fecha", true, file);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        // Click Aceptar
                                        CrudKed3pcco.Click(desktopSession);
                                        newFunctions_4.ScreenshotNuevo("Boton Aceptar", true, file);

                                        //  Enter
                                        CrudKed3pcco.Enter(desktopSession);
                                        newFunctions_4.ScreenshotNuevo("pantalla1", true, file);

                                        //  Enter
                                        CrudKed3pcco.Enter(desktopSession);
                                        newFunctions_4.ScreenshotNuevo("pantalla2", true, file);

                                        // Click Errores
                                        CrudKed3pcco.ClickBtnErrores(desktopSession);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Errores", true, file);
                                        Thread.Sleep(500);

                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                                        newFunctions_4.ScreenshotNuevo("Programa Finalizado", true, file);
                                        Thread.Sleep(1000);


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
        public void KEd3PlccChecKlist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosED.ModulosED_6.KEd3PlccChecKlist")
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
                                //rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                //rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                //rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                //rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&

                                //Variables Validaciones
                                rows["Lupa"].ToString().Length != 0 && rows["Lupa"].ToString() != null )

                            {
                                //Variables obligatorias
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();

                                //Variables generales
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                //string banExcel = rows["banExcel"].ToString();
                                //string BanderaLupa = rows["BanderaLupa"].ToString();
                                //string BanderaPreli = rows["BanderaPreli"].ToString();
                                //string BanderaSum = rows["BanderaSum"].ToString();

                                //Variables validacion
                                string Lupa = rows["Lupa"].ToString();

                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { Lupa };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "ED", motor);

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
                                        Thread.Sleep(2000);


                                        //agregfar crud falta em comparacion con meta
                                        CrudKed3plcc.LupaDinamica(desktopSession, crudPrincipal);
                                        Thread.Sleep(1500);


                                        ////QBE2
                                        CrudKed3plcc.qbe2(desktopSession);
                                        newFunctions_4.ScreenshotNuevo(" QBE", true, file);
                                        Thread.Sleep(500);

                                        CrudKed3plcc.Enter(desktopSession);
                                        Thread.Sleep(700);


                                        CrudKed3plcc.Enter(desktopSession);
                                        newFunctions_4.ScreenshotNuevo("Registros QBE", true, file);
                                        Thread.Sleep(1000);

                                        ////QBE
                                        //errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        //// Click Boton Aceptar.
                                        CrudKed3plcc.Click(desktopSession);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Boton Aceptar", true, file);

                                        newFunctions_4.ScreenshotNuevo("pantalla1", true, file);
                                        CrudKed3plcc.Enter(desktopSession);
                                        Thread.Sleep(60000);

                                        newFunctions_4.ScreenshotNuevo("pantalla2", true, file);
                                        CrudKed3plcc.Enter(desktopSession);
                                        Thread.Sleep(60000);

                                        Thread.Sleep(60000);
                                        newFunctions_4.ScreenshotNuevo("pantalla3", true, file);
                                        CrudKed3plcc.Enter(desktopSession);
                                        Thread.Sleep(60000);
                                     
                                        //// Click Boton Errores.
                                        CrudKed3plcc.ClickBtnErrores(desktopSession);
                                        Thread.Sleep(60000);
                                        newFunctions_4.ScreenshotNuevo("Boton Errores", true, file);
                                        Thread.Sleep(60000);

                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                                        newFunctions_4.ScreenshotNuevo("Programa Finalizado", true, file);
                                        Thread.Sleep(1000);

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
        public void KEdImetaChecKlist()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosED.ModulosED_6.KEdImetaChecKlist")
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

                                //Variables Validaciones
                                rows["Lupa"].ToString().Length != 0 && rows["Lupa"].ToString() != null)

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
                                string Lupa = rows["Lupa"].ToString();

                                //LISTA CRUD PRINCIPAL
                                List<string> crudPrincipal = new List<string>() { Lupa };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "ED", motor);

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
                                        Thread.Sleep(2500);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1500);

                                        CrudKedlmeta.LupaDinamica(desktopSession, crudPrincipal);
                                        Thread.Sleep(1500);

                                        CrudKedlmeta.AgregarCrud(desktopSession);
                                        Thread.Sleep(1500);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(2000);


                                        newFunctions_4.ScreenshotNuevo("Boton Aceptar", true, file);
                                        CrudKedlmeta.Click(desktopSession);
                                        Thread.Sleep(1000);



                                        newFunctions_4.ScreenshotNuevo(" Ventana 1", true, file);
                                        CrudKedlmeta.Enter(desktopSession);
                                        Thread.Sleep(1000);

                                        newFunctions_4.ScreenshotNuevo(" Ventana 2", true, file);
                                        CrudKedlmeta.Enter(desktopSession);
                                        Thread.Sleep(1000);
                                        
                                      
                                        CrudKedlmeta.Click2(desktopSession);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Boton Errores", true, file);
                                        Thread.Sleep(1000);

                                        ////SUMATORIA
                                        //errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //newFunctions_1.lapiz(desktopSession);

                                        ////REPORTE DINAMICO
                                        //string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
                                        //errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        ////Rectificar Pie Pagina PDF DINAMICO
                                        //listaResultado = Textopdf(pdf, codProgram, user);
                                        //listaResultado.ForEach(e => ErrorValidacion.Add(e));

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

                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                                        newFunctions_4.ScreenshotNuevo("Programa Finalizado", true, file);

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

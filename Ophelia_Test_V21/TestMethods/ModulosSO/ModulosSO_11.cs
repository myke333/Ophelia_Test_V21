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
using OpenQA.Selenium.Interactions;

//add Daniel
using System.Windows.Forms;
using System.Drawing.Imaging;

namespace Ophelia_Test_V21.TestMethods.ModulosSO
{
    /// <summary>
    /// Descripción resumida de ModulosSO_11
    /// </summary>
    [TestClass]
    public class ModulosSO_11 : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();
        public ModulosSO_11()
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
        public void KSoPamalAbrirPrograma()
        {
            string file = CrearDocumentoWordDinamico("KSoPamal", "SO", "SQL");
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();

            AbrirPrograma a = new AbrirPrograma("KSoPamal", "UCalida1");
            desktopSession = a.Start2("SQL", "");

            string BanderaLupa = "1";
            string TipoQbe = "0";
            string QbeFilter = "22";
            string BanderaPreli = "1";
            Thread.Sleep(1000);

            string Identificacion = "1";
            string IdentificacionEditar = "2";

            List<string> crudVariables = new List<string>() { Identificacion, IdentificacionEditar };

            string FechaI = "29/12/2020";
            string Seguimiento = "Prueba Seguimiento";
            string Observacion = "Prueba Observacion";
            string Estado = "P";
            string ObservacionEditar = "Prueba Observacion Editar";

            List<string> crudVariablesdet1 = new List<string>() { FechaI, Seguimiento, Observacion, Estado, ObservacionEditar };


            string FechaPreli = "18/08/2020";

            ////QBE
            //errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
            //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            //Thread.Sleep(1000);
            ////REPORTE PRELIMINAR
            //string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
            //errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
            //CrudKsoviepi.PreliKSoViepi(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli, FechaPreli);

            //LUPA
            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

            //////////AGREGAR REGISTRO
            Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
            //botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
            var ElementList = PruebaCRUD.NavClass(desktopSession);

            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            //desktopSession.Mouse.Click(null);
            //CrudKsoviepi.CRUDKSoViepi(desktopSession, 1, crudVariables, file);
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(200);
            //newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

            //////Validacion Agregar
            ////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);

            ////DESCARTAR EDICION
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            //desktopSession.Mouse.Click(null);
            //CrudKsoviepi.CRUDKSoViepi(desktopSession, 2, crudVariables, file);
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(1000);
            //newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

            //////Validacion Editar Descartar
            //////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, "OBS_ERVA", c, ErrorValidacion, "No se edito correctamente", 1);

            ////ACEPTAR EDICION
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            //desktopSession.Mouse.Click(null);
            //CrudKsoviepi.CRUDKSoViepi(desktopSession, 2, crudVariables, file);
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(1000);
            //newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

            //////////AGREGAR DETALLE 1
            //botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
            //var ElementList1 = PruebaCRUD.NavClass(desktopSession);

            //desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            //desktopSession.Mouse.Click(null);
            //CrudKsoviepi.CRUDDetalle1KSoViepi(desktopSession, 1, crudVariablesdet1, file);
            //desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(200);
            //newFunctions_4.ScreenshotNuevo("Agregar Detalle 1", true, file);

            ////DESCARTAR EDICION
            //desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            //desktopSession.Mouse.Click(null);
            //CrudKsoviepi.CRUDDetalle1KSoViepi(desktopSession, 2, crudVariablesdet1, file);
            //desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(1000);
            //newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 1", true, file);

            ////ACEPTAR EDICION
            //desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            //desktopSession.Mouse.Click(null);
            //CrudKsoviepi.CRUDDetalle1KSoViepi(desktopSession, 2, crudVariablesdet1, file);
            //desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(1000);
            //newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 1", true, file);

            ////ELIMINAR DETALLE 1
            //desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(500);
            //PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[1], botonesNavegador, file);
            //Thread.Sleep(1000);

            ////ELIMINAR REGISTRO
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(500);
            //PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
            //Thread.Sleep(1000);
        }

        [TestMethod]
        public void KSoParamCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_11.KSoParamCheckList")
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
                                rows["CorreoEEMp"].ToString().Length != 0 && rows["CorreoEEMp"].ToString() != null
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

                                string CorreoEEMp = rows["CorreoEEMp"].ToString();
                                string TipoAreaI = rows["TipoAreaI"].ToString();
                                string VulnePerso = rows["VulnePerso"].ToString();
                                string VulneRecu = rows["VulneRecu"].ToString();
                                string Vulnesis = rows["Vulnesis"].ToString();
                                string VulnesisEditar = rows["VulnesisEditar"].ToString();
                                List<string> crudVariables = new List<string>() { CorreoEEMp };

                                string Anio = rows["Anio"].ToString();
                                string anioEditar = rows["anioEditar"].ToString();
                                List<string> crudVariablesdet1 = new List<string>() { Anio, anioEditar };
                                if (motor == "SQL")
                                {
                                    string Intelectual = rows["Intelectual"].ToString();
                                    string Personal = rows["Personal"].ToString();
                                    string Valorativa = rows["Valorativa"].ToString();
                                    string Fisiologica = rows["Fisiologica"].ToString();

                                    string MinFirCand = rows["MinFirCand"].ToString();
                                    string PorcMinCand = rows["PorcMinCand"].ToString();
                                    string CodCarcar = rows["CodCarcar"].ToString();                                    

                                    crudVariables = new List<string>() { CorreoEEMp, Intelectual, Personal, Valorativa, Fisiologica,
                                    MinFirCand, PorcMinCand, CodCarcar, TipoAreaI, VulnePerso, VulneRecu, Vulnesis, VulnesisEditar };

                                }
                                if (motor == "ORA")
                                {
                                    string CompeOcu = rows["CompeOcu"].ToString();

                                    crudVariables = new List<string>() { CorreoEEMp, TipoAreaI, VulnePerso, VulneRecu, Vulnesis, CompeOcu, VulnesisEditar };

                                    string Anio1 = rows["Anio1"].ToString();
                                    string Anio2 = rows["Anio2"].ToString();
                                    crudVariablesdet1 = new List<string>() { Anio1, Anio2, Anio, anioEditar };
                                }

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                //Debugger.Launch();
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);

                                try
                                {
                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();
                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
                                    //Debugger.Launch();
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
                                        Thread.Sleep(1000);
                                        newFunctions_1.lapiz(desktopSession);
                                        Thread.Sleep(500);

                                        if (motor == "ORA")
                                        {
                                            //PASAR DETALLE 1
                                            CrudKsoparam.PasarPestaña(desktopSession, 11, 1);
                                        }

                                        ////IDENTIFICAR BOTONERAS
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        if (motor == "ORA")
                                        {
                                            botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                            ElementList = PruebaCRUD.NavClass(desktopSession);
                                            //ELIMINAR DETALLE 1
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador, file);
                                            Thread.Sleep(500);
                                            botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                            var ElementList2 = PruebaCRUD.NavClass(desktopSession);
                                            Thread.Sleep(500);

                                            desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(500);
                                            newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
                                            Thread.Sleep(500);
                                            WindowsDriver<WindowsElement> rootSession = null;
                                            rootSession = PruebaCRUD.RootSession();
                                            rootSession = PruebaCRUD.ReloadSession(rootSession, "TMessageForm");
                                            var OK = rootSession.FindElementsByName("OK");
                                            OK[0].Click();
                                            newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
                                            Thread.Sleep(2000);

                                            //PASAR DETALLE 1
                                            CrudKsoparam.PasarPestaña(desktopSession, 2, 2);
                                        }
                                        //ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        Thread.Sleep(500);

                                        //Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CorreoEEMp, "", "COR_EMPL", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, CorreoEEMp, "D", c, ErrorValidacion);

                                        //////////AGREGAR REGISTRO           
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoparam.CRUDKSoParam(desktopSession, 1, crudVariables, file, motor);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

                                        //Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CorreoEEMp, CorreoEEMp, "COR_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //PASAR A PESTAÑA 8
                                        CrudKsoparam.PasarPestaña(desktopSession, 6, 1);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoparam.CRUDKSoParam(desktopSession, 2, crudVariables, file, motor);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        //Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CorreoEEMp, Vulnesis, "VUL_SIST", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoparam.CRUDKSoParam(desktopSession, 2, crudVariables, file, motor);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, CorreoEEMp, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CorreoEEMp, VulnesisEditar, "VUL_SIST", c, ErrorValidacion, "El registro no se editó correctamente", 1);


                                        //PASAR DETALLE 1
                                        CrudKsoparam.PasarPestaña(desktopSession, 11, 1);

                                        ////////AGREGAR DETALLE 1
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoparam.CRUDDetalle1KSoParam(desktopSession, 1, crudVariablesdet1, file, motor);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 1", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoparam.CRUDDetalle1KSoParam(desktopSession, 2, crudVariablesdet1, file, motor);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 1", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoparam.CRUDDetalle1KSoParam(desktopSession, 2, crudVariablesdet1, file, motor);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 1", true, file);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

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
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, CorreoEEMp, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));
                                        Thread.Sleep(1000);
                                        newFunctions_1.lapiz(desktopSession);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //PASAR DETALLE 1
                                        CrudKsoparam.PasarPestaña(desktopSession, 9, 1);

                                        //ELIMINAR DETALLE 1
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador, file);
                                        Thread.Sleep(500);

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
                                    Thread.Sleep(5000);

                                    // Add Juan
                                    newFunctions_4.ScreenshotNuevo("Error en la prueba", true, file);

                                    // Add Daniel                                   
                                    TestContext.AddResultFile(newFunctions_4.ScreenshotInicio(codProgram));

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
        public void KSoPlaemCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_11.KSoPlaemCheckList")
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
                                //VARIABLES VALIDACION
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["TablaD1"].ToString().Length != 0 && rows["TablaD1"].ToString() != null &&
                                rows["CampoD1"].ToString().Length != 0 && rows["CampoD1"].ToString() != null &&
                                rows["TablaD2"].ToString().Length != 0 && rows["TablaD2"].ToString() != null &&
                                rows["CampoD2"].ToString().Length != 0 && rows["CampoD2"].ToString() != null &&
                                rows["TablaD3"].ToString().Length != 0 && rows["TablaD3"].ToString() != null &&
                                rows["CampoD3"].ToString().Length != 0 && rows["CampoD3"].ToString() != null &&
                                rows["TablaD4"].ToString().Length != 0 && rows["TablaD4"].ToString() != null &&
                                rows["CampoD4"].ToString().Length != 0 && rows["CampoD4"].ToString() != null &&
                                rows["TablaD5"].ToString().Length != 0 && rows["TablaD5"].ToString() != null &&
                                rows["CampoD5"].ToString().Length != 0 && rows["CampoD5"].ToString() != null &&
                                rows["TablaD6"].ToString().Length != 0 && rows["TablaD6"].ToString() != null &&
                                rows["CampoD6"].ToString().Length != 0 && rows["CampoD6"].ToString() != null &&
                                rows["TablaD7"].ToString().Length != 0 && rows["TablaD7"].ToString() != null &&
                                rows["CampoD7"].ToString().Length != 0 && rows["CampoD7"].ToString() != null &&
                                rows["TablaD8"].ToString().Length != 0 && rows["TablaD8"].ToString() != null &&
                                rows["CampoD8"].ToString().Length != 0 && rows["CampoD8"].ToString() != null &&
                                //CABECERA
                                rows["NombreP"].ToString().Length != 0 && rows["NombreP"].ToString() != null &&
                                rows["Vigencia"].ToString().Length != 0 && rows["Vigencia"].ToString() != null &&
                                rows["Version"].ToString().Length != 0 && rows["Version"].ToString() != null &&
                                rows["FechaE"].ToString().Length != 0 && rows["FechaE"].ToString() != null &&
                                rows["NombrePEditar"].ToString().Length != 0 && rows["NombrePEditar"].ToString() != null &&
                                //DETALLE 1
                                rows["PuntoC"].ToString().Length != 0 && rows["PuntoC"].ToString() != null &&
                                rows["PuntoCEditar"].ToString().Length != 0 && rows["PuntoCEditar"].ToString() != null &&
                                //DETALLE 2
                                rows["Area"].ToString().Length != 0 && rows["Area"].ToString() != null &&
                                rows["AreaEditar"].ToString().Length != 0 && rows["AreaEditar"].ToString() != null &&
                                //DETALLE 3
                                rows["Registro"].ToString().Length != 0 && rows["Registro"].ToString() != null &&
                                rows["Requerimiento"].ToString().Length != 0 && rows["Requerimiento"].ToString() != null &&
                                rows["Especificacion"].ToString().Length != 0 && rows["Especificacion"].ToString() != null &&
                                rows["TipoR"].ToString().Length != 0 && rows["TipoR"].ToString() != null &&
                                rows["Tercero"].ToString().Length != 0 && rows["Tercero"].ToString() != null &&
                                rows["Direccion"].ToString().Length != 0 && rows["Direccion"].ToString() != null &&
                                rows["Telefono"].ToString().Length != 0 && rows["Telefono"].ToString() != null &&
                                rows["Email"].ToString().Length != 0 && rows["Email"].ToString() != null &&
                                rows["EmpresaO"].ToString().Length != 0 && rows["EmpresaO"].ToString() != null &&
                                rows["DireccionEditar"].ToString().Length != 0 && rows["DireccionEditar"].ToString() != null &&
                                //DETALLE 4
                                rows["NombreD"].ToString().Length != 0 && rows["NombreD"].ToString() != null &&
                                rows["Grado"].ToString().Length != 0 && rows["Grado"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["NombreDEditar"].ToString().Length != 0 && rows["NombreDEditar"].ToString() != null &&
                                //DETALLE 5
                                rows["TAccion"].ToString().Length != 0 && rows["TAccion"].ToString() != null &&
                                rows["Responsable"].ToString().Length != 0 && rows["Responsable"].ToString() != null &&
                                rows["Estado"].ToString().Length != 0 && rows["Estado"].ToString() != null &&
                                rows["EstadoEditar"].ToString().Length != 0 && rows["EstadoEditar"].ToString() != null &&
                                //DETALLE 6
                                rows["FechaS"].ToString().Length != 0 && rows["FechaS"].ToString() != null &&
                                rows["EstadoS"].ToString().Length != 0 && rows["EstadoS"].ToString() != null &&
                                rows["EstadoSEditar"].ToString().Length != 0 && rows["EstadoSEditar"].ToString() != null &&
                                //DETALLE 7
                                rows["GradoE"].ToString().Length != 0 && rows["GradoE"].ToString() != null &&
                                rows["Nivel"].ToString().Length != 0 && rows["Nivel"].ToString() != null &&
                                rows["GradoEEditar"].ToString().Length != 0 && rows["GradoEEditar"].ToString() != null &&
                                //DETALLE 7
                                rows["CodigoC"].ToString().Length != 0 && rows["CodigoC"].ToString() != null &&
                                rows["CodigoCEditar"].ToString().Length != 0 && rows["CodigoCEditar"].ToString() != null

                                )
                            {
                                //VARIABLES GENERALES
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();
                                //VARIABLES VALIDACION
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string TablaD1 = rows["TablaD1"].ToString();
                                string CampoD1 = rows["CampoD1"].ToString();
                                string TablaD2 = rows["TablaD2"].ToString();
                                string CampoD2 = rows["CampoD2"].ToString();
                                string TablaD3 = rows["TablaD3"].ToString();
                                string CampoD3 = rows["CampoD3"].ToString();
                                string TablaD4 = rows["TablaD4"].ToString();
                                string CampoD4 = rows["CampoD4"].ToString();
                                string TablaD5 = rows["TablaD5"].ToString();
                                string CampoD5 = rows["CampoD5"].ToString();
                                string TablaD6 = rows["TablaD6"].ToString();
                                string CampoD6 = rows["CampoD6"].ToString();
                                string TablaD7 = rows["TablaD7"].ToString();
                                string CampoD7 = rows["CampoD7"].ToString();
                                string TablaD8 = rows["TablaD8"].ToString();
                                string CampoD8 = rows["CampoD8"].ToString();
                                //CABECERA
                                string NombreP = rows["NombreP"].ToString();
                                string Vigencia = rows["Vigencia"].ToString();
                                string Version = rows["Version"].ToString();
                                string FechaE = rows["FechaE"].ToString();
                                string NombrePEditar = rows["NombrePEditar"].ToString();
                                //DETALLE 1
                                string PuntoC = rows["PuntoC"].ToString();
                                string PuntoCEditar = rows["PuntoCEditar"].ToString();
                                //DETALLE 2
                                string Area = rows["Area"].ToString();
                                string AreaEditar = rows["AreaEditar"].ToString();
                                //DETALLE 3
                                string Registro = rows["Registro"].ToString();
                                string Requerimiento = rows["Requerimiento"].ToString();
                                string Especificacion = rows["Especificacion"].ToString();
                                string TipoR = rows["TipoR"].ToString();
                                string Tercero = rows["Tercero"].ToString();
                                string Direccion = rows["Direccion"].ToString();
                                string Telefono = rows["Telefono"].ToString();
                                string Email = rows["Email"].ToString();
                                string EmpresaO = rows["EmpresaO"].ToString();
                                string DireccionEditar = rows["DireccionEditar"].ToString();
                                //DETALLE 4
                                string NombreD = rows["NombreD"].ToString();
                                string Grado = rows["Grado"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string NombreDEditar = rows["NombreDEditar"].ToString();
                                //DETALLE 5
                                string TAccion = rows["TAccion"].ToString();
                                string Responsable = rows["Responsable"].ToString();
                                string Estado = rows["Estado"].ToString();
                                string EstadoEditar = rows["EstadoEditar"].ToString();
                                //DETALLE 6
                                string FechaS = rows["FechaS"].ToString();
                                string EstadoS = rows["EstadoS"].ToString();
                                string EstadoSEditar = rows["EstadoSEditar"].ToString();
                                //DETALLE 7
                                string GradoE = rows["GradoE"].ToString();
                                string Nivel = rows["Nivel"].ToString();
                                string GradoEEditar = rows["GradoEEditar"].ToString();
                                //DETALLE 8
                                string CodigoC = rows["CodigoC"].ToString();
                                string CodigoCEditar = rows["CodigoCEditar"].ToString();

                                List<string> crudVariables = new List<string>() { NombreP, Vigencia, Version, FechaE, NombrePEditar };

                                List<string> crudVariablesD1 = new List<string>() { PuntoC, PuntoCEditar };

                                List<string> crudVariablesD2 = new List<string>() { Area, AreaEditar };

                                List<string> crudVariablesD3 = new List<string>() { Registro, Requerimiento, Especificacion, TipoR, Tercero, Direccion, Telefono, Email, EmpresaO, DireccionEditar };

                                List<string> crudVariablesD4 = new List<string>() { NombreD, Grado, Observacion, NombreDEditar };

                                List<string> crudVariablesD5 = new List<string>() { TAccion, Responsable, Estado, EstadoEditar };

                                List<string> crudVariablesD6 = new List<string>() { FechaS, EstadoS, EstadoSEditar };

                                List<string> crudVariablesD7 = new List<string>() { GradoE, Nivel, GradoEEditar };

                                List<string> crudVariablesD8 = new List<string>() { CodigoC, CodigoCEditar };


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                //Debugger.Launch();
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);

                                try
                                {
                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();
                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
                                    //Debugger.Launch();
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

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 8
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD8, user, motor, CampoD8, CodigoC, CodigoC, "COD_CARG", c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 7
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD7, user, motor, CampoD7, Nivel, Nivel, "COD_NIVE", c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 6
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD6, user, motor, CampoD6, Responsable, Responsable, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTRO PEGADOS DETALLE 5
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD5, user, motor, CampoD5, Responsable, Responsable, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTRO PEGADOS DETALLE 4
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD4, user, motor, CampoD4, Observacion, Observacion, "OBS_ERVA", c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 3
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD3, user, motor, CampoD3, Tercero, Tercero, "COD_TERC", c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD2, user, motor, CampoD2, AreaEditar, AreaEditar, "COD_AREA", c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD1, user, motor, CampoD1, PuntoCEditar, PuntoCEditar, "PUN_CARD", c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //ELIMINAR REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Version, Version, "VER_SION", c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        ////////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKSoPlaem(desktopSession, 1, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Version, Version, "VER_SION", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKSoPlaem(desktopSession, 2, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Version, NombreP, "NOM_PLAN", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKSoPlaem(desktopSession, 2, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(3000);
                                        bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsoplaem.CRUDKSoPlaem(desktopSession, 2, crudVariables, file);
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Version, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Version, NombrePEditar, "NOM_PLAN", c, ErrorValidacion, "El registro no se editó correctamente", 1);
                                        Thread.Sleep(1000);
                                        //////AGREGAR DETALLE 1
                                        Dictionary<string, Point> botonesNavegadorD1 = new Dictionary<string, Point>();
                                        botonesNavegadorD1 = PruebaCRUD.findname(desktopSession, 27, 1);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorD1["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle1SoPlaem(desktopSession, 1, crudVariablesD1, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Agregar detalle", true, file);

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD1, user, motor, CampoD1, PuntoC, PuntoC, "PUN_CARD", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle1SoPlaem(desktopSession, 2, crudVariablesD1, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorD1["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD1, user, motor, CampoD1, PuntoC, PuntoC, "PUN_CARD", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle1SoPlaem(desktopSession, 2, crudVariablesD1, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsoplaem.CRUDKDetalle1SoPlaem(desktopSession, 2, crudVariablesD1, file);
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD1, motor, user, CampoD1, Version, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD1, user, motor, CampoD1, PuntoCEditar, PuntoCEditar, "PUN_CARD", c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        ////////AGREGAR DETALLE 2
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorD1["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle2SoPlaem(desktopSession, 1, crudVariablesD2, file);
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Agregar detalle", true, file);

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD2, user, motor, CampoD2, Area, Area, "COD_AREA", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle2SoPlaem(desktopSession, 2, crudVariablesD2, file);
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorD1["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD2, user, motor, CampoD2, Area, Area, "COD_AREA", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle2SoPlaem(desktopSession, 2, crudVariablesD2, file);
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsoplaem.CRUDKDetalle2SoPlaem(desktopSession, 2, crudVariablesD2, file);
                                            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD2, motor, user, CampoD2, Version, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD2, user, motor, CampoD2, AreaEditar, AreaEditar, "COD_AREA", c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //Pasar Detalle 3
                                        CrudKsoplaem.PasarPestaña(desktopSession, 2, 1, 0);
                                        Thread.Sleep(500);

                                        ////////AGREGAR DETALLE 3
                                        var ElementList2 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorD1["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        CrudKsoplaem.CRUDKDetalle3SoPlaem(desktopSession, 1, crudVariablesD3, file);
                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Agregar detalle", true, file);

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD3, user, motor, CampoD3, Tercero, Tercero, "COD_TERC", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle3SoPlaem(desktopSession, 2, crudVariablesD3, file);
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorD1["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD3, user, motor, CampoD3, Tercero, Direccion, "DIR_TERC", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle3SoPlaem(desktopSession, 2, crudVariablesD3, file);
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsoplaem.CRUDKDetalle3SoPlaem(desktopSession, 2, crudVariablesD3, file);
                                            desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD3, motor, user, CampoD3, Version, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD3, user, motor, CampoD3, Tercero, DireccionEditar, "DIR_TERC", c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //Pasar Detalle 4
                                        CrudKsoplaem.PasarPestaña(desktopSession, 3, 1, 0);
                                        Thread.Sleep(500);

                                        ////////AGREGAR DETALLE 4
                                        var ElementList3 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorD1["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle4SoPlaem(desktopSession, 1, crudVariablesD4, file);
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Agregar detalle", true, file);

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD4, user, motor, CampoD4, Observacion, Observacion, "OBS_ERVA", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle4SoPlaem(desktopSession, 2, crudVariablesD4, file);
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorD1["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD4, user, motor, CampoD4, Observacion, NombreD, "NOM_DOCU", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle4SoPlaem(desktopSession, 2, crudVariablesD4, file);
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsoplaem.CRUDKDetalle4SoPlaem(desktopSession, 2, crudVariablesD4, file);
                                            desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD4, motor, user, CampoD4, Version, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD4, user, motor, CampoD4, Observacion, NombreDEditar, "NOM_DOCU", c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //Pasar Detalle 5
                                        CrudKsoplaem.PasarPestaña(desktopSession, 4, 1, 0);
                                        Thread.Sleep(500);

                                        ////////AGREGAR DETALLE 5            
                                        Dictionary<string, Point> botonesNavegadorD5 = new Dictionary<string, Point>();
                                        botonesNavegadorD5 = PruebaCRUD.findname(desktopSession, 24, 2);
                                        var ElementList5 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList5[2].Coordinates, botonesNavegadorD5["Insertar"].X, botonesNavegadorD5["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle5SoPlaem(desktopSession, 1, crudVariablesD5, file, motor);
                                        desktopSession.Mouse.MouseMove(ElementList5[2].Coordinates, botonesNavegadorD5["Aplicar"].X, botonesNavegadorD5["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Agregar detalle", true, file);

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD5, user, motor, CampoD5, Responsable, Responsable, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList5[2].Coordinates, botonesNavegadorD5["Editar"].X, botonesNavegadorD5["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle5SoPlaem(desktopSession, 2, crudVariablesD5, file, motor);
                                        desktopSession.Mouse.MouseMove(ElementList5[2].Coordinates, botonesNavegadorD5["Cancelar"].X, botonesNavegadorD5["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD5, user, motor, CampoD5, Responsable, Estado, "EST_TADO", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList5[2].Coordinates, botonesNavegadorD5["Editar"].X, botonesNavegadorD5["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle5SoPlaem(desktopSession, 2, crudVariablesD5, file, motor);
                                        desktopSession.Mouse.MouseMove(ElementList5[2].Coordinates, botonesNavegadorD5["Aplicar"].X, botonesNavegadorD5["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList5[2].Coordinates, botonesNavegadorD5["Editar"].X, botonesNavegadorD5["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsoplaem.CRUDKDetalle5SoPlaem(desktopSession, 2, crudVariablesD5, file, motor);
                                            desktopSession.Mouse.MouseMove(ElementList5[2].Coordinates, botonesNavegadorD5["Aplicar"].X, botonesNavegadorD5["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD5, motor, user, CampoD5, Version, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD5, user, motor, CampoD5, Responsable, EstadoEditar, "EST_TADO", c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        ////////AGREGAR DETALLE 6
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorD5["Insertar"].X, botonesNavegadorD5["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle6SoPlaem(desktopSession, 1, crudVariablesD6, file);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorD5["Aplicar"].X, botonesNavegadorD5["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Agregar detalle", true, file);

                                        ////Validacion Agregar
                                       // ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD5, user, motor, CampoD6, Responsable, Responsable, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorD5["Editar"].X, botonesNavegadorD5["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle6SoPlaem(desktopSession, 2, crudVariablesD6, file);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorD5["Cancelar"].X, botonesNavegadorD5["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD5, user, motor, CampoD6, Responsable, EstadoS, "EST_TADO", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorD5["Editar"].X, botonesNavegadorD5["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle6SoPlaem(desktopSession, 2, crudVariablesD6, file);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorD5["Aplicar"].X, botonesNavegadorD5["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorD5["Editar"].X, botonesNavegadorD5["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsoplaem.CRUDKDetalle6SoPlaem(desktopSession, 2, crudVariablesD6, file);
                                            desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorD5["Aplicar"].X, botonesNavegadorD5["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD5, motor, user, CampoD6, Version, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD5, user, motor, CampoD6, Responsable, EstadoSEditar, "EST_TADO", c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //Pasar Detalle 7
                                        CrudKsoplaem.PasarPestaña(desktopSession, 5, 1, 0);
                                        Thread.Sleep(500);

                                        ////////AGREGAR DETALLE 7
                                        Dictionary<string, Point> botonesNavegadorD7 = new Dictionary<string, Point>();
                                        botonesNavegadorD7 = PruebaCRUD.findname(desktopSession, 23, 2);
                                        var ElementList7 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorD7["Insertar"].X, botonesNavegadorD7["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle7SoPlaem(desktopSession, 1, crudVariablesD7, file);
                                        desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorD7["Aplicar"].X, botonesNavegadorD7["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Agregar detalle", true, file);

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD7, user, motor, CampoD7, Nivel, Nivel, "COD_NIVE", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorD7["Editar"].X, botonesNavegadorD7["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle7SoPlaem(desktopSession, 2, crudVariablesD7, file);
                                        desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorD7["Cancelar"].X, botonesNavegadorD7["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD7, user, motor, CampoD7, Nivel, GradoE, "GRA_EMER", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorD7["Editar"].X, botonesNavegadorD7["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle7SoPlaem(desktopSession, 2, crudVariablesD7, file);
                                        desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorD7["Aplicar"].X, botonesNavegadorD7["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorD7["Editar"].X, botonesNavegadorD7["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsoplaem.CRUDKDetalle7SoPlaem(desktopSession, 2, crudVariablesD7, file);
                                            desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorD7["Aplicar"].X, botonesNavegadorD7["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD7, motor, user, CampoD7, Nivel, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD7, user, motor, CampoD7, Nivel, GradoEEditar, "GRA_EMER", c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        ////////AGREGAR DETALLE 8
                                        desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegadorD7["Insertar"].X, botonesNavegadorD7["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle8SoPlaem(desktopSession, 1, crudVariablesD8, file);
                                        desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegadorD7["Aplicar"].X, botonesNavegadorD7["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Agregar detalle", true, file);

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD8, user, motor, CampoD8, CodigoC, CodigoC, "COD_CARG", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegadorD7["Editar"].X, botonesNavegadorD7["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle8SoPlaem(desktopSession, 2, crudVariablesD8, file);
                                        desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegadorD7["Cancelar"].X, botonesNavegadorD7["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD8, user, motor, CampoD8, CodigoC, CodigoC, "COD_CARG", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegadorD7["Editar"].X, botonesNavegadorD7["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoplaem.CRUDKDetalle8SoPlaem(desktopSession, 2, crudVariablesD8, file);
                                        desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegadorD7["Aplicar"].X, botonesNavegadorD7["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegadorD7["Editar"].X, botonesNavegadorD7["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKsoplaem.CRUDKDetalle8SoPlaem(desktopSession, 2, crudVariablesD8, file);
                                            desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegadorD7["Aplicar"].X, botonesNavegadorD7["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(1000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD8, motor, user, CampoD8, CodigoCEditar, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD8, user, motor, CampoD8, CodigoCEditar, CodigoCEditar, "COD_CARG", c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file,"1");
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

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
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, Version, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));
                                        Thread.Sleep(1000);
                                        newFunctions_1.lapiz(desktopSession);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR DETALLE 8
                                        CrudKsoplaem.EliminarRegistro(desktopSession, ElementList7[1], botonesNavegadorD7, file);
                                        Thread.Sleep(1000);

                                        //Validacion Eliminar Registro
                                       // ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD8, user, motor, CampoD8, CodigoCEditar, "", "COD_CARG", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD8, motor, user, CampoD8, CodigoCEditar, "D", c, ErrorValidacion);

                                        //ELIMINAR DETALLE 7
                                        CrudKsoplaem.EliminarRegistro(desktopSession, ElementList7[2], botonesNavegadorD7, file);
                                        Thread.Sleep(1000);

                                        //Validacion Eliminar Registro
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD7, user, motor, CampoD7, Nivel, "", "COD_NIVE", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD7, motor, user, CampoD7, Nivel, "D", c, ErrorValidacion);

                                        //Pasar Detalle 5
                                        CrudKsoplaem.PasarPestaña(desktopSession, 4, 1, 0);
                                        Thread.Sleep(500);

                                        //ELIMINAR DETALLE 6
                                        CrudKsoplaem.EliminarRegistro(desktopSession, ElementList5[1], botonesNavegadorD5, file);
                                        Thread.Sleep(1000);

                                        //Validacion Eliminar Registro
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD5, user, motor, CampoD6, Responsable, "", "COD_EMPL", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD5, motor, user, CampoD6, Responsable, "D", c, ErrorValidacion);

                                        //ELIMINAR DETALLE 5
                                        CrudKsoplaem.EliminarRegistro(desktopSession, ElementList5[2], botonesNavegadorD5, file);
                                        Thread.Sleep(1000);

                                        //Validacion Eliminar Registro
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD5, user, motor, CampoD5, Responsable, "", "COD_EMPL", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD5, motor, user, CampoD5, Responsable, "D", c, ErrorValidacion);

                                        //Pasar Detalle 4
                                        CrudKsoplaem.PasarPestaña(desktopSession, 3, 1, 0);
                                        Thread.Sleep(500);

                                        //ELIMINAR DETALLE 4
                                        CrudKsoplaem.EliminarRegistro(desktopSession, ElementList3[1], botonesNavegadorD1, file);
                                        Thread.Sleep(1000);

                                        //Validacion Eliminar Registro
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD4, user, motor, CampoD4, Observacion, "", "OBS_ERVA", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD4, motor, user, CampoD4, Observacion, "D", c, ErrorValidacion);

                                        //Pasar Detalle 3
                                        CrudKsoplaem.PasarPestaña(desktopSession, 2, 1, 0);
                                        Thread.Sleep(500);

                                        //ELIMINAR DETALLE 3
                                        CrudKsoplaem.EliminarRegistro(desktopSession, ElementList2[1], botonesNavegadorD1, file);
                                        Thread.Sleep(1000);

                                        //Validacion Eliminar Registro
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD3, user, motor, CampoD3, Tercero, "", "COD_TERC", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD3, motor, user, CampoD3, Tercero, "D", c, ErrorValidacion);

                                        //Pasar detalle 2
                                        var TPageControl = desktopSession.FindElementsByClassName("TPageControl");
                                        desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 30, 13);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);

                                        //ELIMINAR DETALLE 2
                                        CrudKsoplaem.EliminarRegistro(desktopSession, ElementList[2], botonesNavegadorD1, file);
                                        Thread.Sleep(1000);

                                        //Validacion Eliminar Registro
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD2, user, motor, CampoD2, AreaEditar, "", "COD_AREA", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD2, motor, user, CampoD2, AreaEditar, "D", c, ErrorValidacion);

                                        //ELIMINAR DETALLE 1
                                        CrudKsoplaem.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorD1, file);
                                        Thread.Sleep(1000);

                                        //Validacion Eliminar Registro
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD1, user, motor, CampoD1, PuntoCEditar, "", "PUN_CARD", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD1, motor, user, CampoD1, PuntoCEditar, "D", c, ErrorValidacion);

                                        //ELIMINAR REGISTRO
                                        CrudKsoplaem.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        Thread.Sleep(1000);

                                        //Validacion Eliminar Registro
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Version, "", "VER_SION", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Version, "D", c, ErrorValidacion);

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
        public void KSoPlaemAbrirPrograma()
        {
            string codProgram = "KSoPlaem";
            string motor = "SQL";

            string user = "";
            if (motor == "SQL")
            {
                user = "UCalida1";
            }
            if (motor == "ORA")
            {
                user = "ODESAR";
            }

            string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            AbrirPrograma a = new AbrirPrograma(codProgram, user);
            desktopSession = a.Start2(motor, "");

            string BanderaLupa = "1";
            string TipoQbe = "0";
            string QbeFilter = "51708624";//52160519
            string BanderaPreli = "1";
            Thread.Sleep(1000);


            string Tabla = "SO_PLAEM";
            string Campo = "VER_SION";

            string NombreP = "PRUEBAS";//
            string Vigencia = "22/02/2021";//
            string Version = "808";//
            string FechaE = "22/02/2021";//
            string NombrePEditar = "PRUEBAS EDITAR";//

            List<string> crudVariables = new List<string>() { NombreP, Vigencia, Version, FechaE, NombrePEditar };

            string TablaD1 = "SO_DLIND";
            string CampoD1 = "PUN_CARD";

            string PuntoC = "N";//
            string PuntoCEditar = "S";//

            List<string> crudVariablesD1 = new List<string>() { PuntoC, PuntoCEditar };

            string TablaD2 = "SO_DARHO";
            string CampoD2 = "COD_AREA";

            string Area = "1";//
            string AreaEditar = "0";//

            List<string> crudVariablesD2 = new List<string>() { Area, AreaEditar };

            string TablaD3 = "SO_DREDI";
            string CampoD3 = "COD_TERC";

            string Registro = "0125";//SQL
            string Requerimiento = "0";//SQL
            string Especificacion = "5";//SQL
            //string Registro = "B";//B
            //string Requerimiento = "9000";//900
            //string Especificacion = "9010";//9010
            string TipoR = "I";
            string Tercero = "1";//SQL
            //string Tercero = "111222333";//
            string Direccion = "PRUEBAS";//
            string Telefono = "123";//
            string Email = "PRUEBAS";//
            string EmpresaO = "1";//421
            //string EmpresaO = "421";//421
            string DireccionEditar = "PRUEBAS EDITAR";//

            List<string> crudVariablesD3 = new List<string>() { Registro, Requerimiento, Especificacion, TipoR, Tercero, Direccion, Telefono, Email, EmpresaO, DireccionEditar };

            string TablaD4 = "SO_DDPON";
            string CampoD4 = "OBS_ERVA";

            string NombreD = "PRUEBAS";//
            string Grado = "1";//
            string Observacion = "808";//
            string NombreDEditar = "PRUEBAS EDITAR";//

            List<string> crudVariablesD4 = new List<string>() { NombreD, Grado, Observacion, NombreDEditar };

            string TablaD5 = "SO_DPEAC";
            string CampoD5 = "COD_EMPL";

            string TAccion = "C";//
            string Responsable = "1000";//SQL
            //string Responsable = "21174134";//21174134
            string Estado = "P";//
            string EstadoEditar = "S";//

            List<string> crudVariablesD5 = new List<string>() { TAccion, Responsable, Estado, EstadoEditar };

            string TablaD6 = "SO_DPESE";
            string CampoD6 = "COD_EMPL";

            string FechaS = "22/02/2021";//
            string EstadoS = "S";//
            string EstadoSEditar = "P";//

            List<string> crudVariablesD6 = new List<string>() { FechaS, EstadoS, EstadoSEditar };

            string TablaD7 = "SO_DGRAD";
            string CampoD7 = "COD_NIVE";

            string GradoE = "1";//
            string Nivel = "0";//
            string GradoEEditar = "2";//

            List<string> crudVariablesD7 = new List<string>() { GradoE, Nivel, GradoEEditar };

            string TablaD8 = "SO_DGRCA";
            string CampoD8 = "COD_CARG";

            string CodigoC = "6";//SQL
            string CodigoCEditar = "14";//SQL
            //string CodigoC = "5151";//5151
            //string CodigoCEditar = "1001";//1001

            List<string> crudVariablesD8 = new List<string>() { CodigoC, CodigoCEditar };

            //LUPA
            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

            ////////AGREGAR REGISTRO
            Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
            botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
            var ElementList = PruebaCRUD.NavClass(desktopSession);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKSoPlaem(desktopSession, 1, crudVariables, file);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

            ////Validacion Agregar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Version, Version, "VER_SION", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKSoPlaem(desktopSession, 2, crudVariables, file);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

            ////Validacion Editar Descartar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Version, NombreP, "NOM_PLAN", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKSoPlaem(desktopSession, 2, crudVariables, file);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
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
                CrudKsoplaem.CRUDKSoPlaem(desktopSession, 2, crudVariables, file);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
            }
            newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Version, "E", c, ErrorValidacion);

            //Validacion Editar Aceptar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Version, NombrePEditar, "NOM_PLAN", c, ErrorValidacion, "El registro no se editó correctamente", 1);

            //////AGREGAR DETALLE 1
            Dictionary<string, Point> botonesNavegadorD1 = new Dictionary<string, Point>();
            botonesNavegadorD1 = PruebaCRUD.findname(desktopSession, 27, 1);
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorD1["Insertar"].X, botonesNavegador["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle1SoPlaem(desktopSession, 1, crudVariablesD1, file);
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Agregar detalle", true, file);

            ////Validacion Agregar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD1, user, motor, CampoD1, PuntoC, PuntoC, "PUN_CARD", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle1SoPlaem(desktopSession, 2, crudVariablesD1, file);
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorD1["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

            ////Validacion Editar Descartar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD1, user, motor, CampoD1, PuntoC, PuntoC, "PUN_CARD", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle1SoPlaem(desktopSession, 2, crudVariablesD1, file);
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            valEditOra = PruebaCRUD.ValEditOra(desktopSession);
            if (valEditOra == true)
            {
                //LUPA
                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                //ACEPTAR EDICION
                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
                desktopSession.Mouse.Click(null);
                CrudKsoplaem.CRUDKDetalle1SoPlaem(desktopSession, 2, crudVariablesD1, file);
                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
            }
            newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD1, motor, user, CampoD1, Version, "E", c, ErrorValidacion);

            //Validacion Editar Aceptar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD1, user, motor, CampoD1, PuntoCEditar, PuntoCEditar, "PUN_CARD", c, ErrorValidacion, "El registro no se editó correctamente", 1);

            ////////AGREGAR DETALLE 2
            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorD1["Insertar"].X, botonesNavegador["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle2SoPlaem(desktopSession, 1, crudVariablesD2, file);
            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Agregar detalle", true, file);

            ////Validacion Agregar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD2, user, motor, CampoD2, Area, Area, "COD_AREA", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle2SoPlaem(desktopSession, 2, crudVariablesD2, file);
            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorD1["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

            ////Validacion Editar Descartar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD2, user, motor, CampoD2, Area, Area, "COD_AREA", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle2SoPlaem(desktopSession, 2, crudVariablesD2, file);
            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD2, motor, user, CampoD2, Version, "E", c, ErrorValidacion);

            //Validacion Editar Aceptar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD2, user, motor, CampoD2, AreaEditar, AreaEditar, "COD_AREA", c, ErrorValidacion, "El registro no se editó correctamente", 1);

            //Pasar Detalle 3
            CrudKsoplaem.PasarPestaña(desktopSession, 2, 1, 0);
            Thread.Sleep(500);

            ////////AGREGAR DETALLE 3
            var ElementList2 = PruebaCRUD.NavClass(desktopSession);
            desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorD1["Insertar"].X, botonesNavegador["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle3SoPlaem(desktopSession, 1, crudVariablesD3, file);
            desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Agregar detalle", true, file);

            ////Validacion Agregar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD3, user, motor, CampoD3, Tercero, Tercero, "COD_TERC", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle3SoPlaem(desktopSession, 2, crudVariablesD3, file);
            desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorD1["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

            ////Validacion Editar Descartar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD3, user, motor, CampoD3, Tercero, Direccion, "DIR_TERC", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle3SoPlaem(desktopSession, 2, crudVariablesD3, file);
            desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000); 
            valEditOra = PruebaCRUD.ValEditOra(desktopSession);
            if (valEditOra == true)
            {
                //LUPA
                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                //ACEPTAR EDICION
                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
                desktopSession.Mouse.Click(null);
                CrudKsoplaem.CRUDKDetalle3SoPlaem(desktopSession, 2, crudVariablesD3, file);
                desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
            }
            newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD3, motor, user, CampoD3, Version, "E", c, ErrorValidacion);

            //Validacion Editar Aceptar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD3, user, motor, CampoD3, Tercero, DireccionEditar, "DIR_TERC", c, ErrorValidacion, "El registro no se editó correctamente", 1);

            //Pasar Detalle 4
            CrudKsoplaem.PasarPestaña(desktopSession, 3, 1, 0);
            Thread.Sleep(500);

            ////////AGREGAR DETALLE 4
            var ElementList3 = PruebaCRUD.NavClass(desktopSession);
            desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorD1["Insertar"].X, botonesNavegador["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle4SoPlaem(desktopSession, 1, crudVariablesD4, file);
            desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Agregar detalle", true, file);

            ////Validacion Agregar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD4, user, motor, CampoD4, Observacion, Observacion, "OBS_ERVA", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle4SoPlaem(desktopSession, 2, crudVariablesD4, file);
            desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorD1["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

            ////Validacion Editar Descartar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD4, user, motor, CampoD4, Observacion, NombreD, "NOM_DOCU", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorD1["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle4SoPlaem(desktopSession, 2, crudVariablesD4, file);
            desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorD1["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD4, motor, user, CampoD4, Version, "E", c, ErrorValidacion);

            //Validacion Editar Aceptar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD4, user, motor, CampoD4, Observacion, NombreDEditar, "NOM_DOCU", c, ErrorValidacion, "El registro no se editó correctamente", 1);

            //Pasar Detalle 5
            CrudKsoplaem.PasarPestaña(desktopSession, 4, 1, 0);
            Thread.Sleep(500);

            ////////AGREGAR DETALLE 5            
            Dictionary<string, Point> botonesNavegadorD5 = new Dictionary<string, Point>();
            botonesNavegadorD5 = PruebaCRUD.findname(desktopSession, 24, 2);
            var ElementList5 = PruebaCRUD.NavClass(desktopSession);
            desktopSession.Mouse.MouseMove(ElementList5[2].Coordinates, botonesNavegadorD5["Insertar"].X, botonesNavegadorD5["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle5SoPlaem(desktopSession, 1, crudVariablesD5, file, motor);
            desktopSession.Mouse.MouseMove(ElementList5[2].Coordinates, botonesNavegadorD5["Aplicar"].X, botonesNavegadorD5["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Agregar detalle", true, file);

            ////Validacion Agregar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD5, user, motor, CampoD5, Responsable, Responsable, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList5[2].Coordinates, botonesNavegadorD5["Editar"].X, botonesNavegadorD5["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle5SoPlaem(desktopSession, 2, crudVariablesD5, file, motor);
            desktopSession.Mouse.MouseMove(ElementList5[2].Coordinates, botonesNavegadorD5["Cancelar"].X, botonesNavegadorD5["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

            ////Validacion Editar Descartar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD5, user, motor, CampoD5, Responsable, Estado, "EST_TADO", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList5[2].Coordinates, botonesNavegadorD5["Editar"].X, botonesNavegadorD5["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle5SoPlaem(desktopSession, 2, crudVariablesD5, file, motor);
            desktopSession.Mouse.MouseMove(ElementList5[2].Coordinates, botonesNavegadorD5["Aplicar"].X, botonesNavegadorD5["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            valEditOra = PruebaCRUD.ValEditOra(desktopSession);
            if (valEditOra == true)
            {
                //LUPA
                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                //ACEPTAR EDICION
                desktopSession.Mouse.MouseMove(ElementList5[2].Coordinates, botonesNavegadorD5["Editar"].X, botonesNavegadorD5["Editar"].Y);
                desktopSession.Mouse.Click(null);
                CrudKsoplaem.CRUDKDetalle5SoPlaem(desktopSession, 2, crudVariablesD5, file, motor);
                desktopSession.Mouse.MouseMove(ElementList5[2].Coordinates, botonesNavegadorD5["Aplicar"].X, botonesNavegadorD5["Aplicar"].Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
            }
            newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD5, motor, user, CampoD5, Version, "E", c, ErrorValidacion);

            //Validacion Editar Aceptar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD5, user, motor, CampoD5, Responsable, EstadoEditar, "EST_TADO", c, ErrorValidacion, "El registro no se editó correctamente", 1);

            ////////AGREGAR DETALLE 6
            desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorD5["Insertar"].X, botonesNavegadorD5["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle6SoPlaem(desktopSession, 1, crudVariablesD6, file);
            desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorD5["Aplicar"].X, botonesNavegadorD5["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Agregar detalle", true, file);

            ////Validacion Agregar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD6, user, motor, CampoD6, Responsable, Responsable, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorD5["Editar"].X, botonesNavegadorD5["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle6SoPlaem(desktopSession, 2, crudVariablesD6, file);
            desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorD5["Cancelar"].X, botonesNavegadorD5["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

            ////Validacion Editar Descartar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD6, user, motor, CampoD6, Responsable, EstadoS, "EST_TADO", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorD5["Editar"].X, botonesNavegadorD5["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle6SoPlaem(desktopSession, 2, crudVariablesD6, file);
            desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorD5["Aplicar"].X, botonesNavegadorD5["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            valEditOra = PruebaCRUD.ValEditOra(desktopSession);
            if (valEditOra == true)
            {
                //LUPA
                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                //ACEPTAR EDICION
                desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorD5["Editar"].X, botonesNavegadorD5["Editar"].Y);
                desktopSession.Mouse.Click(null);
                CrudKsoplaem.CRUDKDetalle6SoPlaem(desktopSession, 2, crudVariablesD6, file);
                desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorD5["Aplicar"].X, botonesNavegadorD5["Aplicar"].Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
            }
            newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD6, motor, user, CampoD6, Version, "E", c, ErrorValidacion);

            //Validacion Editar Aceptar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD6, user, motor, CampoD6, Responsable, EstadoSEditar, "EST_TADO", c, ErrorValidacion, "El registro no se editó correctamente", 1);

            //Pasar Detalle 7
            CrudKsoplaem.PasarPestaña(desktopSession, 5, 1, 0);
            Thread.Sleep(500);

            ////////AGREGAR DETALLE 7
            Dictionary<string, Point> botonesNavegadorD7 = new Dictionary<string, Point>();
            botonesNavegadorD7 = PruebaCRUD.findname(desktopSession, 23, 2);
            var ElementList7 = PruebaCRUD.NavClass(desktopSession);
            desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorD7["Insertar"].X, botonesNavegadorD7["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle7SoPlaem(desktopSession, 1, crudVariablesD7, file);
            desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorD7["Aplicar"].X, botonesNavegadorD7["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Agregar detalle", true, file);

            ////Validacion Agregar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD7, user, motor, CampoD7, Nivel, Nivel, "COD_NIVE", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorD7["Editar"].X, botonesNavegadorD7["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle7SoPlaem(desktopSession, 2, crudVariablesD7, file);
            desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorD7["Cancelar"].X, botonesNavegadorD7["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

            ////Validacion Editar Descartar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD7, user, motor, CampoD7, Nivel, GradoE, "GRA_EMER", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorD7["Editar"].X, botonesNavegadorD7["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle7SoPlaem(desktopSession, 2, crudVariablesD7, file);
            desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorD7["Aplicar"].X, botonesNavegadorD7["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD7, motor, user, CampoD7, Nivel, "E", c, ErrorValidacion);

            //Validacion Editar Aceptar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD7, user, motor, CampoD7, Nivel, GradoEEditar, "GRA_EMER", c, ErrorValidacion, "El registro no se editó correctamente", 1);

            ////////AGREGAR DETALLE 8
            desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegadorD7["Insertar"].X, botonesNavegadorD7["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle8SoPlaem(desktopSession, 1, crudVariablesD8, file);
            desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegadorD7["Aplicar"].X, botonesNavegadorD7["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Agregar detalle", true, file);

            ////Validacion Agregar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD8, user, motor, CampoD8, CodigoC, CodigoC, "COD_CARG", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegadorD7["Editar"].X, botonesNavegadorD7["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle8SoPlaem(desktopSession, 2, crudVariablesD8, file);
            desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegadorD7["Cancelar"].X, botonesNavegadorD7["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

            ////Validacion Editar Descartar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD8, user, motor, CampoD8, CodigoC, CodigoC, "COD_CARG", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegadorD7["Editar"].X, botonesNavegadorD7["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKsoplaem.CRUDKDetalle8SoPlaem(desktopSession, 2, crudVariablesD8, file);
            desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegadorD7["Aplicar"].X, botonesNavegadorD7["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            valEditOra = PruebaCRUD.ValEditOra(desktopSession);
            if (valEditOra == true)
            {
                //LUPA
                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                //ACEPTAR EDICION
                desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegadorD7["Editar"].X, botonesNavegadorD7["Editar"].Y);
                desktopSession.Mouse.Click(null);
                CrudKsoplaem.CRUDKDetalle8SoPlaem(desktopSession, 2, crudVariablesD8, file);
                desktopSession.Mouse.MouseMove(ElementList7[1].Coordinates, botonesNavegadorD7["Aplicar"].X, botonesNavegadorD7["Aplicar"].Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
            }
            newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD8, motor, user, CampoD8, CodigoCEditar, "E", c, ErrorValidacion);

            //Validacion Editar Aceptar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD8, user, motor, CampoD8, CodigoCEditar, CodigoCEditar, "COD_CARG", c, ErrorValidacion, "El registro no se editó correctamente", 1);

            //ELIMINAR DETALLE 8
            CrudKsoplaem.EliminarRegistro(desktopSession, ElementList7[1], botonesNavegadorD7, file);
            Thread.Sleep(1000);

            //Validacion Eliminar Registro
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD8, user, motor, CampoD8, CodigoCEditar, "", "COD_CARG", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

            //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD8, motor, user, CampoD8, CodigoCEditar, "D", c, ErrorValidacion);

            //ELIMINAR DETALLE 7
            CrudKsoplaem.EliminarRegistro(desktopSession, ElementList7[2], botonesNavegadorD7, file);
            Thread.Sleep(1000);

            //Validacion Eliminar Registro
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD7, user, motor, CampoD7, Nivel, "", "COD_NIVE", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

            //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD7, motor, user, CampoD7, Nivel, "D", c, ErrorValidacion);

            //Pasar Detalle 5
            CrudKsoplaem.PasarPestaña(desktopSession, 4, 1, 0);
            Thread.Sleep(500);

            //ELIMINAR DETALLE 6
            CrudKsoplaem.EliminarRegistro(desktopSession, ElementList5[1], botonesNavegadorD5, file);
            Thread.Sleep(1000);

            //Validacion Eliminar Registro
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD6, user, motor, CampoD6, Responsable, "", "COD_EMPL", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

            //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD5, motor, user, CampoD6, Responsable, "D", c, ErrorValidacion);

            //ELIMINAR DETALLE 5
            CrudKsoplaem.EliminarRegistro(desktopSession, ElementList5[2], botonesNavegadorD5, file);
            Thread.Sleep(1000);

            //Validacion Eliminar Registro
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD5, user, motor, CampoD5, Responsable, "", "COD_EMPL", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

            //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD5, motor, user, CampoD5, Responsable, "D", c, ErrorValidacion);

            //Pasar Detalle 4
            CrudKsoplaem.PasarPestaña(desktopSession, 3, 1, 0);
            Thread.Sleep(500);

            //ELIMINAR DETALLE 4
            CrudKsoplaem.EliminarRegistro(desktopSession, ElementList3[1], botonesNavegadorD1, file);
            Thread.Sleep(1000);

            //Validacion Eliminar Registro
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD4, user, motor, CampoD4, Observacion, "", "OBS_ERVA", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

            //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD4, motor, user, CampoD4, Observacion, "D", c, ErrorValidacion);

            //Pasar Detalle 3
            CrudKsoplaem.PasarPestaña(desktopSession, 2, 1, 0);
            Thread.Sleep(500);

            //ELIMINAR DETALLE 3
            CrudKsoplaem.EliminarRegistro(desktopSession, ElementList2[1], botonesNavegadorD1, file);
            Thread.Sleep(1000);

            //Validacion Eliminar Registro
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD3, user, motor, CampoD3, Tercero, "", "COD_TERC", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

            //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD3, motor, user, CampoD3, Tercero, "D", c, ErrorValidacion);

            //Pasar detalle 2
            var TPageControl = desktopSession.FindElementsByClassName("TPageControl");
            desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 30, 13);
            desktopSession.Mouse.Click(null);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);

            //ELIMINAR DETALLE 2
            CrudKsoplaem.EliminarRegistro(desktopSession, ElementList[2], botonesNavegadorD1, file);
            Thread.Sleep(1000);

            //Validacion Eliminar Registro
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD2, user, motor, CampoD2, AreaEditar, "", "COD_AREA", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

            //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD2, motor, user, CampoD2, AreaEditar, "D", c, ErrorValidacion);

            //ELIMINAR DETALLE 1
            CrudKsoplaem.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorD1, file);
            Thread.Sleep(1000);

            //Validacion Eliminar Registro
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(TablaD1, user, motor, CampoD1, PuntoCEditar, "", "PUN_CARD", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

            //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(TablaD1, motor, user, CampoD1, PuntoCEditar, "D", c, ErrorValidacion);

            //ELIMINAR REGISTRO
            CrudKsoplaem.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
            Thread.Sleep(1000);

            //Validacion Eliminar Registro
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Version, "", "VER_SION", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

            //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Version, "D", c, ErrorValidacion);
        }
    }
}

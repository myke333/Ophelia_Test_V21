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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_6;
using OpenQA.Selenium.Interactions;

namespace Ophelia_Test_V21.TestMethods.ModulosNM
{
    /// <summary>
    /// Descripción resumida de ModulosNM_58
    /// </summary>
    [TestClass]
    public class ModulosNM_58: FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();
        public ModulosNM_58()
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
        public void KNmNcconCheckList()
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
            //TableOrder = "ktest1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_58.KNmNcconCheckList")
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
                                //rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                //VARIABLES VALIDACION
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["Tabla1"].ToString().Length != 0 && rows["Tabla1"].ToString() != null &&
                                rows["Campo1"].ToString().Length != 0 && rows["Campo1"].ToString() != null &&
                                rows["Tabla2"].ToString().Length != 0 && rows["Tabla2"].ToString() != null &&
                                rows["Campo2"].ToString().Length != 0 && rows["Campo2"].ToString() != null &&
                                rows["Tabla3"].ToString().Length != 0 && rows["Tabla3"].ToString() != null &&
                                rows["Campo3"].ToString().Length != 0 && rows["Campo3"].ToString() != null &&
                                //VARIABLES CRUD
                                rows["Codigo"].ToString().Length != 0 && rows["Codigo"].ToString() != null &&
                                rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
                                rows["NombreEditar"].ToString().Length != 0 && rows["NombreEditar"].ToString() != null &&
                                //VARIABLES DETALLE 1
                                rows["Secuencial"].ToString().Length != 0 && rows["Secuencial"].ToString() != null &&
                                rows["SecuencialEditar"].ToString().Length != 0 && rows["SecuencialEditar"].ToString() != null &&
                                //VARIABLES DETALLE 2
                                rows["Secuencial2"].ToString().Length != 0 && rows["Secuencial2"].ToString() != null &&
                                rows["Secuencial2Editar"].ToString().Length != 0 && rows["Secuencial2Editar"].ToString() != null &&
                                //VARIABLES DETALLE 3
                                rows["CodUsuario"].ToString().Length != 0 && rows["CodUsuario"].ToString() != null &&
                                rows["CodUsuarioEditar"].ToString().Length != 0 && rows["CodUsuarioEditar"].ToString() != null
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
                                string Tabla1 = rows["Tabla1"].ToString();
                                string Campo1 = rows["Campo1"].ToString();
                                string Tabla2 = rows["Tabla2"].ToString();
                                string Campo2 = rows["Campo2"].ToString();
                                string Tabla3 = rows["Tabla3"].ToString();
                                string Campo3 = rows["Campo3"].ToString();
                                //VARIABLES CRUD 
                                string Codigo = rows["Codigo"].ToString();
                                string Nombre = rows["Nombre"].ToString();
                                string NombreEditar = rows["NombreEditar"].ToString();
                                //VARIABLES DETALLE 1
                                string Secuencial = rows["Secuencial"].ToString();
                                string SecuencialEditar = rows["SecuencialEditar"].ToString();
                                //VARIABLES DETALLE 2
                                string Secuencial2 = rows["Secuencial2"].ToString();
                                string Secuencial2Editar = rows["Secuencial2Editar"].ToString();
                                //VARIABLES DETALLE 3
                                string CodUsuario = rows["CodUsuario"].ToString();
                                string CodUsuarioEditar = rows["CodUsuarioEditar"].ToString();

                                List<string> crudVariables = new List<string>() { Codigo, Nombre, NombreEditar };

                                List<string> crudVariables1 = new List<string>() { Secuencial, SecuencialEditar };

                                List<string> crudVariables2 = new List<string>() { Secuencial2, Secuencial2Editar };

                                List<string> crudVariables3 = new List<string>() { CodUsuario, CodUsuarioEditar };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
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
                                        
                                        ////////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmnccon.CRUDKNmNccon(desktopSession, 1, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, "COD_CONT", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmnccon.CRUDKNmNccon(desktopSession, 2, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, "NOM_CONT", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmnccon.CRUDKNmNccon(desktopSession, 2, crudVariables, file);
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
                                            
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                            CrudKnmnccon.CRUDKNmNccon(desktopSession, 2, crudVariables, file);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);
                                            
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, NombreEditar, "NOM_CONT", c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(3000);
                                        //Debugger.Launch();
                                        ////////AGREGAR DETALLE
                                        Dictionary<string, Point> botonesNavegador1 = new Dictionary<string, Point>();
                                        botonesNavegador1 = PruebaCRUD.findname(desktopSession, 27, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Insertar"].X, botonesNavegador1["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmnccon.CRUDKDetalle1KNmNccon(desktopSession, 1, crudVariables1, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Codigo, Codigo, "COD_CONT", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmnccon.CRUDKDetalle1KNmNccon(desktopSession, 2, crudVariables1, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Cancelar"].X, botonesNavegador1["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Codigo, Secuencial, "SEC_CAMP", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmnccon.CRUDKDetalle1KNmNccon(desktopSession, 2, crudVariables1, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla1, motor, user, Campo1, Codigo, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Codigo, SecuencialEditar, "SEC_CAMP", c, ErrorValidacion, "El registro no se editó correctamente", 1);
                                        Thread.Sleep(3000);
                                        //PASAR DETALLE 2
                                        CrudKnmnccon.PasarPestaña(desktopSession, 1, 0);
                                        Thread.Sleep(1000);

                                        ////////AGREGAR DETALLE 2            
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Insertar"].X, botonesNavegador1["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmnccon.CRUDKDetalle1KNmNccon(desktopSession, 1, crudVariables2, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, Campo2, Codigo, Codigo, "COD_CONT", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmnccon.CRUDKDetalle1KNmNccon(desktopSession, 2, crudVariables2, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Cancelar"].X, botonesNavegador1["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, Campo2, Codigo, Secuencial, "SEC_CAMP", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmnccon.CRUDKDetalle1KNmNccon(desktopSession, 2, crudVariables2, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla2, motor, user, Campo1, Codigo, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, Campo2, Codigo, SecuencialEditar, "SEC_CAMP", c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //PASAR DETALLE 3
                                        CrudKnmnccon.PasarPestaña(desktopSession, 2, 0);
                                        Thread.Sleep(1000);

                                        ////////AGREGAR DETALLE 3          
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Insertar"].X, botonesNavegador1["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmnccon.CRUDKDetalle3KNmNccon(desktopSession, 1, crudVariables3, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla3, user, motor, Campo3, Codigo, Codigo, "COD_CONT", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmnccon.CRUDKDetalle3KNmNccon(desktopSession, 2, crudVariables3, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Cancelar"].X, botonesNavegador1["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla3, user, motor, Campo3, Codigo, CodUsuario, "COD_USUA", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmnccon.CRUDKDetalle3KNmNccon(desktopSession, 2, crudVariables3, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla3, motor, user, Campo3, Codigo, "E", c, ErrorValidacion);

                                        //Valiacion Editar Aceptar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla3, user, motor, Campo3, Codigo, CodUsuarioEditar, "COD_USUA", c, ErrorValidacion, "El registro no se editó correctamente", 1);

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
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, Codigo, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));
                                        newFunctions_1.lapiz(desktopSession);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(500);

                                        //ELIMINAR DETALLE 3
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador1, file);
                                        Thread.Sleep(500);

                                        //Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla3, user, motor, Campo3, Codigo, "", "COD_CONT", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla3, motor, user, Campo3, Codigo, "D", c, ErrorValidacion);

                                        //PASAR DETALLE 2
                                        CrudKnmnccon.PasarPestaña(desktopSession, 1, 0);

                                        //ELIMINAR DETALLE 2
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador1, file);
                                        Thread.Sleep(500);

                                        //Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla2, user, motor, Campo2, Codigo, "", "COD_CONT", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla2, motor, user, Campo2, Codigo, "D", c, ErrorValidacion);

                                        //PASAR DETALLE 2
                                        CrudKnmnccon.PasarPestaña(desktopSession, 0, 0);

                                        //ELIMINAR DETALLE
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador1, file);
                                        Thread.Sleep(500);

                                        //Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Codigo, "", "COD_CONT", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla1, motor, user, Campo1, Codigo, "D", c, ErrorValidacion);

                                        //ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        Thread.Sleep(500);

                                        //Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, "", "COD_CONT", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "D", c, ErrorValidacion);

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
        public void KNmCincaAbrirPrograma()
        {
            string codProgram = "KNmCinca";
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

            string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            AbrirPrograma a = new AbrirPrograma(codProgram, user);
            desktopSession = a.Start2(motor, "");

            string BanderaLupa = "1";
            string TipoQbe = "0";
            string QbeFilter = "1";//52160519
            string BanderaPreli = "1";
            Thread.Sleep(1000);

            ////LUPA
            //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
            //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

            string Tabla = "NM_NCCON";
            string Campo = "COD_CONT";

            string Codigo = "808";
            string Nombre = "PRUEBAS";
            string NombreEditar = "PRUEBAS EDITAR";

            List<string> crudVariables = new List<string>() { Codigo, Nombre, NombreEditar };

            string Tabla1 = "NM_NCCDI";
            string Campo1 = "COD_CONT";

            string Secuencial = "2112";
            string SecuencialEditar = "8080";

            List<string> crudVariables1 = new List<string>() { Secuencial, SecuencialEditar };

            ////////AGREGAR REGISTRO
            Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
            botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
            var ElementList = PruebaCRUD.NavClass(desktopSession);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKnmnccon.CRUDKNmNccon(desktopSession, 1, crudVariables, file);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

            ////Validacion Agregar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, "COD_CONT", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKnmnccon.CRUDKNmNccon(desktopSession, 2, crudVariables, file);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

            ////Validacion Editar Descartar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, "NOM_CONT", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKnmnccon.CRUDKNmNccon(desktopSession, 2, crudVariables, file);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "E", c, ErrorValidacion);

            //Validacion Editar Aceptar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, NombreEditar, "NOM_CONT", c, ErrorValidacion, "El registro no se editó correctamente", 1);

            ////////AGREGAR DETALLE
            Dictionary<string, Point> botonesNavegador1 = new Dictionary<string, Point>();
            botonesNavegador1 = PruebaCRUD.findname(desktopSession, 27, 1);
            ElementList = PruebaCRUD.NavClass(desktopSession);
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Insertar"].X, botonesNavegador1["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKnmnccon.CRUDKDetalle1KNmNccon(desktopSession, 1, crudVariables1, file);
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

            ////Validacion Agregar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Codigo, Codigo, "COD_CONT", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKnmnccon.CRUDKDetalle1KNmNccon(desktopSession, 2, crudVariables1, file);
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Cancelar"].X, botonesNavegador1["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

            ////Validacion Editar Descartar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Codigo, Secuencial, "SEC_CAMP", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKnmnccon.CRUDKDetalle1KNmNccon(desktopSession, 2, crudVariables1, file);
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla1, motor, user, Campo1, Codigo, "E", c, ErrorValidacion);

            //Validacion Editar Aceptar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Codigo, SecuencialEditar, "SEC_CAMP", c, ErrorValidacion, "El registro no se editó correctamente", 1);
        }

        [TestMethod]
        public void KNmCinctAbrirPrograma()
        {
            string codProgram = "KNmCinct";
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

            string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            AbrirPrograma a = new AbrirPrograma(codProgram, user);
            desktopSession = a.Start2(motor, "");

            string BanderaLupa = "1";
            string TipoQbe = "0";
            string QbeFilter = "1";//52160519
            string BanderaPreli = "1";
            Thread.Sleep(1000);

            ////LUPA
            //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
            //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

            string FechaP = "25/01/2010";
            string Año = "2010";
            string Secuencial = "808";

            List<string> crudVariables = new List<string>() { FechaP, Año, Secuencial };

            CrudKnmcinct.CRUDKNmCinct(desktopSession, 1, crudVariables, file);

            //QBE
            errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

            CrudKnmcinct.CRUDKNmCinct(desktopSession, 2, crudVariables, file);
        }

        [TestMethod]
        public void KNmContrAbrirPrograma()
        {
            string codProgram = "KNmContr";
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

            string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            AbrirPrograma a = new AbrirPrograma(codProgram, user);
            desktopSession = a.Start2(motor, "");

            string BanderaLupa = "1";
            string TipoQbe = "0";
            string QbeFilter = "1";//52160519
            string BanderaPreli = "1";
            Thread.Sleep(1000);

            ////LUPA
            //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
            //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

            string Tabla = "NM_NCCON";
            string Campo = "COD_CONT";

            string Codigo = "808";
            string Nombre = "PRUEBAS";
            string NombreEditar = "PRUEBAS EDITAR";

            List<string> crudVariables = new List<string>() { Codigo, Nombre, NombreEditar };

            string Tabla1 = "NM_NCCDI";
            string Campo1 = "COD_CONT";

            string Secuencial = "2112";
            string SecuencialEditar = "8080";

            List<string> crudVariables1 = new List<string>() { Secuencial, SecuencialEditar };

            ////////AGREGAR REGISTRO
            Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
            botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
            var ElementList = PruebaCRUD.NavClass(desktopSession);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKnmnccon.CRUDKNmNccon(desktopSession, 1, crudVariables, file);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

            ////Validacion Agregar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, "COD_CONT", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKnmnccon.CRUDKNmNccon(desktopSession, 2, crudVariables, file);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

            ////Validacion Editar Descartar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, "NOM_CONT", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKnmnccon.CRUDKNmNccon(desktopSession, 2, crudVariables, file);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "E", c, ErrorValidacion);

            //Validacion Editar Aceptar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, NombreEditar, "NOM_CONT", c, ErrorValidacion, "El registro no se editó correctamente", 1);

            ////////AGREGAR DETALLE
            Dictionary<string, Point> botonesNavegador1 = new Dictionary<string, Point>();
            botonesNavegador1 = PruebaCRUD.findname(desktopSession, 27, 1);
            ElementList = PruebaCRUD.NavClass(desktopSession);
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Insertar"].X, botonesNavegador1["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKnmnccon.CRUDKDetalle1KNmNccon(desktopSession, 1, crudVariables1, file);
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

            ////Validacion Agregar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Codigo, Codigo, "COD_CONT", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKnmnccon.CRUDKDetalle1KNmNccon(desktopSession, 2, crudVariables1, file);
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Cancelar"].X, botonesNavegador1["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

            ////Validacion Editar Descartar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Codigo, Secuencial, "SEC_CAMP", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKnmnccon.CRUDKDetalle1KNmNccon(desktopSession, 2, crudVariables1, file);
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla1, motor, user, Campo1, Codigo, "E", c, ErrorValidacion);

            //Validacion Editar Aceptar
            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla1, user, motor, Campo1, Codigo, SecuencialEditar, "SEC_CAMP", c, ErrorValidacion, "El registro no se editó correctamente", 1);
        }


        [TestMethod]
        public void AbrirPrograma()
        {
            string codProgram = "KNmBepro";
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

            AbrirPrograma a = new AbrirPrograma(codProgram, user);
            desktopSession = a.Start2(motor, "");

            Thread.Sleep(100000);
        }
    }
}

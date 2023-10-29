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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3;

namespace Ophelia_Test_V21.TestMethods.ModulosNM
{
    /// <summary>
    /// Descripción resumida de ModulosNM_33
    /// </summary>
    [TestClass]
    public class ModulosNM_33 : FuncionesVitales
    {

        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();
        public ModulosNM_33()
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
        public void AbrirPrograma()
        {
            //TODO: Preguntar por error al agregar en oracle
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            string programa = "KNmAjuss";
            AbrirPrograma a = new
            AbrirPrograma(programa, "UCalida1");
            desktopSession = a.Start2("SQL", "");
            Thread.Sleep(2000);
            List<string> crudData = new List<string>() { };
            string file = CrearDocumentoWordDinamico("prueba", "Pruebas", "SQL");


            string TipoQbe = "0";
            string QbeFilter = "769099420";
            string BanderaSum = "0";
            string BanderaPreli = "0";
            string banExcel = "0";
            string codProgram = programa;
            string BanderaLupa = "1";
            string motor = "ORA";

            string fechaLiquidacion = "05/01/2021";
            string fechaPago = "05/01/2021";
            string prototipo = "5000";

            Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
            Dictionary<string, Point> botonesNavegadorDet1 = new Dictionary<string, Point>();
            Dictionary<string, Point> botonesNavegadorDet2 = new Dictionary<string, Point>();
            List<string> crudVariables = new List<string>() { prototipo, fechaLiquidacion, fechaPago };
            bool isErrorDisplayed = false;

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
            //Thread.Sleep(1000);

            //botonesNavegador = PruebaCRUD.findname(desktopSession, 27, 0);
            //var ElementList = PruebaCRUD.NavClass(desktopSession);

            ////AGREGAR REGISTRO
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            //desktopSession.Mouse.Click(null);
            CrudKnmajuss.KNmAjussCRUD(desktopSession, crudVariables, 0);
            ////QBE
            errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            Thread.Sleep(2000);
            CrudKnmajuss.BotonAceptar(desktopSession, file);
            Thread.Sleep(3000);
            //REPORTE PRELIMINAR
            //////Rectificar Pie Pagina PDF PRELIMINAR
            //listaResultado = Textopdf(pdf1, codProgram, user);
            //listaResultado.ForEach(e => ErrorValidacion.Add(e));
            //Thread.Sleep(6000);
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(2000);
            //newFunctions_4.ScreenshotNuevo("Registro Agregado", true, file);
            //////////////////////////Validacion Agregar
            //////////////////////////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, identificacion, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

            /////////DESCARTAR EDICION
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            //desktopSession.Mouse.Click(null);
            //CrudKnmtipsi.KNmTipsiCRUD(desktopSession, crudVariables, 1);
            //Thread.Sleep(1000);
            //newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(2000);
            //newFunctions_4.ScreenshotNuevo("Edición Descartada", true, file);
            ////////////////////////Validacion Editar Descartar
            //////////////////////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, tipo, editCampo, c, ErrorValidacion, "No se edito correctamente", 1);

            ////ACEPTAR EDICION
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            //desktopSession.Mouse.Click(null);
            //CrudKnmtipsi.KNmTipsiCRUD(desktopSession, crudVariables, 1);
            //Thread.Sleep(1000);
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(2000);
            //isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

            //if (isErrorDisplayed)
            //{
            //    ////ACEPTAR EDICION
            //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            //    desktopSession.Mouse.Click(null);
            //    CrudKnmtipsi.KNmTipsiCRUD(desktopSession, crudVariables, 1);
            //    Thread.Sleep(1000);
            //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            //    desktopSession.Mouse.Click(null);
            //    Thread.Sleep(2000);
            //}
            //newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
            //////////////////ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, identificacion, "E", c, ErrorValidacion);
            ////////////////////Validacion Editar Aceptar
            //////////////////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, editTipo, editCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);
            //////////////////Thread.Sleep(1000);
            ////////////////


            //botonesNavegadorDet1 = PruebaCRUD.findname(desktopSession, 25, 1);
            //ElementList = PruebaCRUD.NavClass(desktopSession);

            ////AGREGAR REGISTRO DETALLE
            //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Insertar"].X, botonesNavegadorDet1["Insertar"].Y);
            //desktopSession.Mouse.Click(null);
            //CrudKnmtipsi.KNmTipsiCRUDDet1(desktopSession, crudVariablesDet1, 0);
            //Thread.Sleep(3000);
            //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(2000);
            //newFunctions_4.ScreenshotNuevo("Registro Agregado Detalle", true, file);
            //////////////////////////Validacion Agregar
            //////////////////////////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, identificacion, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

            /////////DESCARTAR EDICION DETALLE
            //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
            //desktopSession.Mouse.Click(null);
            //CrudKnmtipsi.KNmTipsiCRUDDet1(desktopSession, crudVariablesDet1, 1);
            //Thread.Sleep(1000);
            //newFunctions_4.ScreenshotNuevo("Registro Editado Detalle", true, file);
            //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Cancelar"].X, botonesNavegadorDet1["Cancelar"].Y);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(2000);
            //newFunctions_4.ScreenshotNuevo("Edición Descartada Detalle", true, file);
            ////////////////////////Validacion Editar Descartar
            //////////////////////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, tipo, editCampo, c, ErrorValidacion, "No se edito correctamente", 1);

            ////ACEPTAR EDICION DETALLE
            //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
            //desktopSession.Mouse.Click(null);
            //CrudKnmtipsi.KNmTipsiCRUDDet1(desktopSession, crudVariablesDet1, 1);
            //Thread.Sleep(1000);
            //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(2000);
            //isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

            //if (isErrorDisplayed)
            //{
            //    ////ACEPTAR EDICION DETALLE
            //    desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
            //    desktopSession.Mouse.Click(null);
            //    CrudKnmtipsi.KNmTipsiCRUDDet1(desktopSession, crudVariablesDet1, 1);
            //    Thread.Sleep(1000);
            //    desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
            //    desktopSession.Mouse.Click(null);
            //    Thread.Sleep(2000);
            //}
            //newFunctions_4.ScreenshotNuevo("Registro Editado Detalle", true, file);
            //////////////////ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, identificacion, "E", c, ErrorValidacion);
            ////////////////////Validacion Editar Aceptar
            //////////////////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, editTipo, editCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);
            //////////////////Thread.Sleep(1000);
            ////////////////


            ////QBE
            //errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
            //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            //Thread.Sleep(1000);
            ////SUMATORIA
            //errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
            //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            //newFunctions_1.lapiz(desktopSession);

            ////REPORTE DINAMICO
            //string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
            //errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
            //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            //////Rectificar Pie Pagina PDF DINAMICO
            ////listaResultado = Textopdf(pdf, codProgram, user);
            ////listaResultado.ForEach(e => ErrorValidacion.Add(e));

            ////REPORTE PRELIMINAR
            //string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
            //errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
            //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            ////////Rectificar Pie Pagina PDF PRELIMINAR
            ////listaResultado = Textopdf(pdf1, codProgram, user);
            ////listaResultado.ForEach(e => ErrorValidacion.Add(e));

            ////EXPORTAR EXCEL
            //string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
            //errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
            //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

            //////RECTIFICACION DE CAMPOS DE EXCEL
            ////string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, identificacion, c);
            ////string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
            ////Celda = Celda_Excel(ruta, nom_empr);
            ////Celda.ForEach(u => ErrorValidacion.Add(u));
            ////Thread.Sleep(1000);
            ////LUPA
            //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
            //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }



            //////ELIMINAR REGISTRO
            //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Borrar"].X, botonesNavegadorDet1["Borrar"].Y);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(500);
            //newFunctions_4.ScreenshotNuevo("Borrar Registro Detalle", true, file);
            //PruebaCRUD.cerrarBorrar(desktopSession);
            //Thread.Sleep(500);
            //newFunctions_4.ScreenshotNuevo("Registro Borrado Detalle", true, file);
            //Thread.Sleep(1000);

            //isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

            //if (isErrorDisplayed)
            //{
            //    ////ELIMINAR REGISTRO
            //    desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Borrar"].X, botonesNavegadorDet1["Borrar"].Y);
            //    desktopSession.Mouse.Click(null);
            //    Thread.Sleep(500);
            //    newFunctions_4.ScreenshotNuevo("Borrar Registro Detalle", true, file);
            //    PruebaCRUD.cerrarBorrar(desktopSession);
            //    Thread.Sleep(500);
            //    newFunctions_4.ScreenshotNuevo("Registro Borrado Detalle", true, file);
            //    Thread.Sleep(1000);
            //}


            //////ELIMINAR REGISTRO
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(500);
            //newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
            //PruebaCRUD.cerrarBorrar(desktopSession);
            //Thread.Sleep(500);
            //newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
            //Thread.Sleep(1000);

            //isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

            //if (isErrorDisplayed)
            //{
            //    ////ELIMINAR REGISTRO
            //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
            //    desktopSession.Mouse.Click(null);
            //    Thread.Sleep(500);
            //    newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
            //    PruebaCRUD.cerrarBorrar(desktopSession);
            //    Thread.Sleep(500);
            //    newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
            //    Thread.Sleep(1000);
            //}


        }

        [TestMethod]
        public void KNmAcArbCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_33.KNmAcArbCheckList")
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
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&
                                //Datos Necesidades Formación
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                rows["OrigenArbol"].ToString().Length != 0 && rows["OrigenArbol"].ToString() != null &&
                                rows["OrigenDestino"].ToString().Length != 0 && rows["OrigenDestino"].ToString() != null &&
                                rows["OrigenNivel"].ToString().Length != 0 && rows["OrigenNivel"].ToString() != null &&
                                rows["NivelDestino"].ToString().Length != 0 && rows["NivelDestino"].ToString() != null
                                )
                            {

                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string editCampo = rows["EditCampo"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();

                                string origeArbol = rows["OrigenArbol"].ToString();
                                string origenDestino = rows["OrigenDestino"].ToString();
                                string origenNivel = rows["OrigenNivel"].ToString();
                                string nivelDestino = rows["NivelDestino"].ToString();

                                Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet1 = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet2 = new Dictionary<string, Point>();
                                List<string> crudVariables = new List<string>() { origeArbol, origenDestino, origenNivel, nivelDestino };
                                bool isErrorDisplayed = false;

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);
                                try
                                {
                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();

                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
                                    Thread.Sleep(5000);
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

                                        ////VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //DATOS
                                        CrudKnmacarb.KNmAcArbCRUD(desktopSession, crudVariables, 0);
                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);
                                        //BOTON ACEPTAR
                                        CrudKnmacarb.BotonAceptar(desktopSession, file);
                                        Thread.Sleep(5000);

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
        public void KNmAjussCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_33.KNmAjussCheckList")
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
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&
                                //Datos Necesidades Formación
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                rows["FechaLiquidacion"].ToString().Length != 0 && rows["FechaLiquidacion"].ToString() != null &&
                                rows["FechaPago"].ToString().Length != 0 && rows["FechaPago"].ToString() != null &&
                                rows["Prototipo"].ToString().Length != 0 && rows["Prototipo"].ToString() != null
                                )
                            {

                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string Tabla = rows["Tabla"].ToString();
                                string Campo = rows["Campo"].ToString();
                                string editCampo = rows["EditCampo"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();

                                string fechaLiquidacion = rows["FechaLiquidacion"].ToString();
                                string fechaPago = rows["FechaPago"].ToString();
                                string prototipo = rows["Prototipo"].ToString();

                                Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet1 = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet2 = new Dictionary<string, Point>();
                                List<string> crudVariables = new List<string>() { prototipo, fechaLiquidacion, fechaPago };
                                bool isErrorDisplayed = false;

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);
                                try
                                {
                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();

                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
                                    Thread.Sleep(5000);
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

                                        ////VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //DATOS
                                        CrudKnmajuss.KNmAjussCRUD(desktopSession, crudVariables, 0);
                                        ////QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(2000);
                                        //ACEPTAR
                                        CrudKnmajuss.BotonAceptar(desktopSession, file);
                                        Thread.Sleep(3000);

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

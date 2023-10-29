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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2;

namespace Ophelia_Test_V21.TestMethods.ModulosNM
{
    /// <summary>
    /// Descripción resumida de ModulosNM_28
    /// </summary>
    [TestClass]
    public class ModulosNM_28 : FuncionesVitales
    {

        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();

        public ModulosNM_28()
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
            string programa = "KSOSEYDE";
            AbrirPrograma a = new
            AbrirPrograma(programa, "UCalida1");
            desktopSession = a.Start2("SQL", "");
            Thread.Sleep(2000000);
            List<string> crudData = new List<string>() { };
            string file = CrearDocumentoWordDinamico("prueba", "Pruebas", "SQL");


            string TipoQbe = "0";
            string QbeFilter = "12";
            string BanderaSum = "0";
            string BanderaPreli = "0";
            string banExcel = "0";
            string codProgram = programa;
            string BanderaLupa = "1";
            string motor = "ORA";

            string tipoContabilizacion = "Nómina";
            string TipoNomina = "Normal";
            string fechaDesde = "01/09/2015";
            string fechaHasta = "30/09/2019";
            string fechaRegistro = "29/12/2010";
            string fechaDocumento = "31/12/2020";

            Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
            Dictionary<string, Point> botonesNavegadorDet1 = new Dictionary<string, Point>();
            Dictionary<string, Point> botonesNavegadorDet2 = new Dictionary<string, Point>();
            List<string> crudVariables = new List<string>() { tipoContabilizacion, TipoNomina, fechaDocumento, fechaRegistro, fechaHasta, fechaDesde };
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
            CrudKnmsapic.KNmSapicCRUD(desktopSession, crudVariables, 0);
            //QBE
            errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            Thread.Sleep(1000);
            CrudKnmsapic.BotonAceptar(desktopSession, file);
            Thread.Sleep(5000);
            //REPORTE PRELIMINAR
            string pdf1 = @"C:\Reportes\ReportePDF1_" + "Errores" + "_" + Hora();
            errors = CrudKnmsapic.BotonErrores(desktopSession, pdf1, file);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
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
        public void KNmPrenoCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_28.KNmPrenoCheckList")
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
                                rows["Identificacion"].ToString().Length != 0 && rows["Identificacion"].ToString() != null &&
                                rows["Concepto"].ToString().Length != 0 && rows["Concepto"].ToString() != null &&
                                rows["FechaInicio"].ToString().Length != 0 && rows["FechaInicio"].ToString() != null &&
                                rows["FechaPago"].ToString().Length != 0 && rows["FechaPago"].ToString() != null &&
                                rows["FechaAcumula"].ToString().Length != 0 && rows["FechaAcumula"].ToString() != null &&
                                rows["ValorTotal"].ToString().Length != 0 && rows["ValorTotal"].ToString() != null &&
                                rows["Valor"].ToString().Length != 0 && rows["Valor"].ToString() != null &&
                                rows["EditConcepto"].ToString().Length != 0 && rows["EditConcepto"].ToString() != null
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

                                string identificacion = rows["Identificacion"].ToString();
                                string concepto = rows["Concepto"].ToString();
                                string fechaInicio = rows["FechaInicio"].ToString();
                                string fechaPago = rows["FechaPago"].ToString();
                                string fechaAcumula = rows["FechaAcumula"].ToString();
                                string valorTotal = rows["ValorTotal"].ToString();
                                string valor = rows["Valor"].ToString();
                                string editConcepto = rows["EditConcepto"].ToString();

                                Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet1 = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet2 = new Dictionary<string, Point>();
                                List<string> crudVariables = new List<string>() { identificacion, concepto, fechaInicio, fechaPago, fechaAcumula, valorTotal, valor, editConcepto };
                                List<string> crudVariablesDet1 = new List<string>() { };
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

                                        //ELIMINAR REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, identificacion, Campo, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //AGREGAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmpreno.KNMPrenoCRUD(desktopSession, crudVariables, 0);
                                        Thread.Sleep(5000);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Registro Agregado", true, file);
                                        //Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, identificacion, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ///////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmpreno.KNMPrenoCRUD(desktopSession, crudVariables, 1);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada", true, file);
                                        //Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, concepto, editCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmpreno.KNMPrenoCRUD(desktopSession, crudVariables, 1);
                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

                                        if (isErrorDisplayed)
                                        {
                                            ////ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKnmpreno.KNMPrenoCRUD(desktopSession, crudVariables, 1);
                                            Thread.Sleep(1000);
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(2000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, identificacion, "E", c, ErrorValidacion);
                                        //Validacion Editar Aceptar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, editConcepto, editCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);
                                        Thread.Sleep(1000);



                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(60000);
                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_1.lapiz(desktopSession);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);
                                        //REPORTE DINAMICO
                                        string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
                                        errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF DINAMICO
                                        listaResultado = Textopdf(pdf, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        ////REPORTE PRELIMINAR
                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = CrudKnmpreno.ReportePreliminarIfnotOptions(desktopSession, pdf1, file, "Desprendible Standar", false);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////RECTIFICACION DE CAMPOS DE EXCEL
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, identificacion, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));
                                        Thread.Sleep(1000);
                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }


                                        ////ELIMINAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
                                        PruebaCRUD.cerrarBorrar(desktopSession);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
                                        Thread.Sleep(1000);

                                        //Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, "", Campo, c, ErrorValidacion, "los datos no se encontraron correctamente", 2);
                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, identificacion, "D", c, ErrorValidacion);

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
        public void KNmTipceCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_28.KNmTipceCheckList")
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
                                rows["Codigo"].ToString().Length != 0 && rows["Codigo"].ToString() != null &&
                                rows["Descripcion"].ToString().Length != 0 && rows["Descripcion"].ToString() != null &&
                                rows["Certificado"].ToString().Length != 0 && rows["Certificado"].ToString() != null &&
                                rows["EditDescripcion"].ToString().Length != 0 && rows["EditDescripcion"].ToString() != null &&
                                rows["Nit"].ToString().Length != 0 && rows["Nit"].ToString() != null &&
                                rows["Titulo"].ToString().Length != 0 && rows["Titulo"].ToString() != null &&
                                rows["IdentificacionDet"].ToString().Length != 0 && rows["IdentificacionDet"].ToString() != null &&
                                rows["CodigoCargo"].ToString().Length != 0 && rows["CodigoCargo"].ToString() != null &&
                                rows["Cargo"].ToString().Length != 0 && rows["Cargo"].ToString() != null &&
                                rows["Proyeccion"].ToString().Length != 0 && rows["Proyeccion"].ToString() != null &&
                                rows["Resolucion"].ToString().Length != 0 && rows["Resolucion"].ToString() != null &&
                                rows["EditResolucion"].ToString().Length != 0 && rows["EditResolucion"].ToString() != null
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

                                string codigo = rows["Codigo"].ToString();
                                string descripcion = rows["Descripcion"].ToString();
                                string certificado = rows["Certificado"].ToString();
                                string editDescripcion = rows["EditDescripcion"].ToString();
                                string identificacion = codigo;

                                string nit = rows["Nit"].ToString();
                                string titulo = rows["Titulo"].ToString();
                                string identificacionDet = rows["IdentificacionDet"].ToString();
                                string codigoCargo = rows["CodigoCargo"].ToString();
                                string cargo = rows["Cargo"].ToString();
                                string proyeccion = rows["Proyeccion"].ToString();
                                string resolucion = rows["Resolucion"].ToString();
                                string editResolucion = rows["EditResolucion"].ToString();


                                Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet1 = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet2 = new Dictionary<string, Point>();
                                List<string> crudVariables = new List<string>() { codigo, descripcion, certificado, editDescripcion };
                                List<string> crudVariablesDet1 = new List<string>() { nit, titulo, identificacionDet, codigoCargo, cargo, proyeccion, resolucion, editResolucion };
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

                                        //ELIMINAR REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, identificacion, Campo, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //AGREGAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmtipce.KNmTipceCRUD(desktopSession, crudVariables, 0);
                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Registro Agregado", true, file);
                                        //Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, identificacion, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ///////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmtipce.KNmTipceCRUD(desktopSession, crudVariables, 1);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada", true, file);
                                        //Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, descripcion, editCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmtipce.KNmTipceCRUD(desktopSession, crudVariables, 1);
                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

                                        if (isErrorDisplayed)
                                        {
                                            ////ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKnmtipce.KNmTipceCRUD(desktopSession, crudVariables, 1);
                                            Thread.Sleep(1000);
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(2000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, identificacion, "E", c, ErrorValidacion);
                                        //Validacion Editar Aceptar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, editDescripcion, editCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);
                                        Thread.Sleep(1000);

                                        if (motor == "SQL") {
                                            botonesNavegadorDet1 = PruebaCRUD.findname(desktopSession, 25, 1);
                                            ElementList = PruebaCRUD.NavClass(desktopSession);

                                            //AGREGAR REGISTRO DETALLE
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Insertar"].X, botonesNavegadorDet1["Insertar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKnmtipce.KNmTipceCRUDDet(desktopSession, crudVariablesDet1, 0);
                                            Thread.Sleep(3000);
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(2000);
                                            newFunctions_4.ScreenshotNuevo("Registro Agregado Detalle", true, file);

                                            ///////DESCARTAR EDICION DETALLE
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKnmtipce.KNmTipceCRUDDet(desktopSession, crudVariablesDet1, 1);
                                            Thread.Sleep(1000);
                                            newFunctions_4.ScreenshotNuevo("Registro Editado Detalle", true, file);
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Cancelar"].X, botonesNavegadorDet1["Cancelar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(2000);
                                            newFunctions_4.ScreenshotNuevo("Edición Descartada Detalle", true, file);

                                            //ACEPTAR EDICION DETALLE
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKnmtipce.KNmTipceCRUDDet(desktopSession, crudVariablesDet1, 1);
                                            Thread.Sleep(1000);
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(2000);
                                            isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

                                            if (isErrorDisplayed)
                                            {
                                                ////ACEPTAR EDICION DETALLE
                                                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                CrudKnmtipce.KNmTipceCRUDDet(desktopSession, crudVariablesDet1, 1);
                                                Thread.Sleep(1000);
                                                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                Thread.Sleep(2000);
                                            }
                                            newFunctions_4.ScreenshotNuevo("Registro Editado Detalle", true, file);


                                        }



                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);
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
                                        ////Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //RECTIFICACION DE CAMPOS DE EXCEL
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, identificacion, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));
                                        Thread.Sleep(1000);
                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        if (motor == "SQL") {
                                            ////ELIMINAR REGISTRO
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Borrar"].X, botonesNavegadorDet1["Borrar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(500);
                                            newFunctions_4.ScreenshotNuevo("Borrar Registro Detalle", true, file);
                                            PruebaCRUD.cerrarBorrar(desktopSession);
                                            Thread.Sleep(500);
                                            newFunctions_4.ScreenshotNuevo("Registro Borrado Detalle", true, file);
                                            Thread.Sleep(1000);

                                            isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

                                            if (isErrorDisplayed)
                                            {
                                                ////ELIMINAR REGISTRO
                                                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Borrar"].X, botonesNavegadorDet1["Borrar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                Thread.Sleep(500);
                                                newFunctions_4.ScreenshotNuevo("Borrar Registro Detalle", true, file);
                                                PruebaCRUD.cerrarBorrar(desktopSession);
                                                Thread.Sleep(500);
                                                newFunctions_4.ScreenshotNuevo("Registro Borrado Detalle", true, file);
                                                Thread.Sleep(1000);
                                            }
                                        }

                                        ////ELIMINAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
                                        PruebaCRUD.cerrarBorrar(desktopSession);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
                                        Thread.Sleep(1000);

                                        isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

                                        if (isErrorDisplayed)
                                        {
                                            ////ELIMINAR REGISTRO
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(500);
                                            newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
                                            PruebaCRUD.cerrarBorrar(desktopSession);
                                            Thread.Sleep(500);
                                            newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
                                            Thread.Sleep(1000);
                                        }

                                        //Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, "", Campo, c, ErrorValidacion, "los datos no se encontraron correctamente", 2);
                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, identificacion, "D", c, ErrorValidacion);

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
        public void KNmTipsiCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_28.KNmTipsiCheckList")
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
                                rows["Codigo"].ToString().Length != 0 && rows["Codigo"].ToString() != null &&
                                rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
                                rows["Clase"].ToString().Length != 0 && rows["Clase"].ToString() != null &&
                                rows["ParametrosAsociados"].ToString().Length != 0 && rows["ParametrosAsociados"].ToString() != null &&
                                rows["EditNombre"].ToString().Length != 0 && rows["EditNombre"].ToString() != null &&
                                rows["CodigoDet"].ToString().Length != 0 && rows["CodigoDet"].ToString() != null &&
                                rows["Descripcion"].ToString().Length != 0 && rows["Descripcion"].ToString() != null &&
                                rows["EditDescripcion"].ToString().Length != 0 && rows["EditDescripcion"].ToString() != null
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

                                string codigo = rows["Codigo"].ToString();
                                string nombre = rows["Nombre"].ToString();
                                string clase = rows["Clase"].ToString();
                                string parametrosAsociados = rows["ParametrosAsociados"].ToString();
                                string editNombre = rows["EditNombre"].ToString();

                                string codigoDet = rows["CodigoDet"].ToString();
                                string descripcion = rows["Descripcion"].ToString();
                                string editDescripcion = rows["EditDescripcion"].ToString();
                                string identificacion = codigo;

                                Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet1 = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet2 = new Dictionary<string, Point>();
                                List<string> crudVariables = new List<string>() { parametrosAsociados, codigo, nombre, clase, editNombre };
                                List<string> crudVariablesDet1 = new List<string>() { codigoDet, descripcion, editDescripcion };
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

                                        //ELIMINAR REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, identificacion, Campo, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 27, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //AGREGAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmtipsi.KNmTipsiCRUD(desktopSession, crudVariables, 0);
                                        Thread.Sleep(6000);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Registro Agregado", true, file);
                                        //Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, identificacion, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ///////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmtipsi.KNmTipsiCRUD(desktopSession, crudVariables, 1);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada", true, file);
                                        //Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, nombre, editCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmtipsi.KNmTipsiCRUD(desktopSession, crudVariables, 1);
                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

                                        if (isErrorDisplayed)
                                        {
                                            ////ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKnmtipsi.KNmTipsiCRUD(desktopSession, crudVariables, 1);
                                            Thread.Sleep(1000);
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(2000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
                                        Thread.Sleep(2000);
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, identificacion, "E", c, ErrorValidacion);
                                        //Validacion Editar Aceptar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, editNombre, editCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);
                                        Thread.Sleep(1000);

                                        botonesNavegadorDet1 = PruebaCRUD.findname(desktopSession, 25, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //AGREGAR REGISTRO DETALLE
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Insertar"].X, botonesNavegadorDet1["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmtipsi.KNmTipsiCRUDDet1(desktopSession, crudVariablesDet1, 0);
                                        Thread.Sleep(3000);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Registro Agregado Detalle", true, file);

                                        ///////DESCARTAR EDICION DETALLE
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmtipsi.KNmTipsiCRUDDet1(desktopSession, crudVariablesDet1, 1);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Registro Editado Detalle", true, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Cancelar"].X, botonesNavegadorDet1["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada Detalle", true, file);

                                        //ACEPTAR EDICION DETALLE
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKnmtipsi.KNmTipsiCRUDDet1(desktopSession, crudVariablesDet1, 1);
                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

                                        if (isErrorDisplayed)
                                        {
                                            ////ACEPTAR EDICION DETALLE
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKnmtipsi.KNmTipsiCRUDDet1(desktopSession, crudVariablesDet1, 1);
                                            Thread.Sleep(1000);
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(2000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Registro Editado Detalle", true, file);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);
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
                                        ////Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //RECTIFICACION DE CAMPOS DE EXCEL
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, identificacion, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));
                                        Thread.Sleep(1000);
                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }



                                        ////ELIMINAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Borrar"].X, botonesNavegadorDet1["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Borrar Registro Detalle", true, file);
                                        PruebaCRUD.cerrarBorrar(desktopSession);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Registro Borrado Detalle", true, file);
                                        Thread.Sleep(1000);

                                        isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

                                        if (isErrorDisplayed)
                                        {
                                            ////ELIMINAR REGISTRO
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Borrar"].X, botonesNavegadorDet1["Borrar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(500);
                                            newFunctions_4.ScreenshotNuevo("Borrar Registro Detalle", true, file);
                                            PruebaCRUD.cerrarBorrar(desktopSession);
                                            Thread.Sleep(500);
                                            newFunctions_4.ScreenshotNuevo("Registro Borrado Detalle", true, file);
                                            Thread.Sleep(1000);
                                        }


                                        ////ELIMINAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
                                        PruebaCRUD.cerrarBorrar(desktopSession);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
                                        Thread.Sleep(1000);

                                        isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

                                        if (isErrorDisplayed)
                                        {
                                            ////ELIMINAR REGISTRO
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(500);
                                            newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
                                            PruebaCRUD.cerrarBorrar(desktopSession);
                                            Thread.Sleep(500);
                                            newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
                                            Thread.Sleep(1000);
                                        }
                                        Thread.Sleep(2000);

                                        //Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, "", Campo, c, ErrorValidacion, "los datos no se encontraron correctamente", 2);
                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, identificacion, "D", c, ErrorValidacion);

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
        public void KNmSapicCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_28.KNmSapicCheckList")
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
                                rows["TipoContabilizacion"].ToString().Length != 0 && rows["TipoContabilizacion"].ToString() != null &&
                                rows["TipoNomina"].ToString().Length != 0 && rows["TipoNomina"].ToString() != null &&
                                rows["FechaDesde"].ToString().Length != 0 && rows["FechaDesde"].ToString() != null &&
                                rows["FechaHasta"].ToString().Length != 0 && rows["FechaHasta"].ToString() != null &&
                                rows["FechaRegistro"].ToString().Length != 0 && rows["FechaRegistro"].ToString() != null &&
                                rows["FechaDocumento"].ToString().Length != 0 && rows["FechaDocumento"].ToString() != null
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

                                string tipoContabilizacion = rows["TipoContabilizacion"].ToString();
                                string tipoNomina = rows["TipoNomina"].ToString();
                                string fechaDesde = rows["FechaDesde"].ToString();
                                string fechaHasta = rows["FechaHasta"].ToString();
                                string fechaRegistro = rows["FechaRegistro"].ToString();
                                string fechaDocumento = rows["FechaDocumento"].ToString();

                                Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet1 = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet2 = new Dictionary<string, Point>();
                                List<string> crudVariables = new List<string>() { tipoContabilizacion, tipoNomina, fechaDocumento, fechaRegistro, fechaHasta, fechaDesde };
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

                                        CrudKnmsapic.KNmSapicCRUD(desktopSession, crudVariables, 0);
                                        newFunctions_4.ScreenshotNuevo("Datos para filtrar registros", true, file);
                                        Thread.Sleep(1000);
                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        CrudKnmsapic.BotonAceptar(desktopSession, file);
                                        Thread.Sleep(5000);
                                        //REPORTE PRELIMINAR
                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Errores" + "_" + Hora();
                                        errors = CrudKnmsapic.BotonErrores(desktopSession, pdf1, file);
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

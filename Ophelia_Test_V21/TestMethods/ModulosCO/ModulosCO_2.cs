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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosCo;

namespace Ophelia_Test_V21.TestMethods.ModulosCO
{
    /// <summary>
    /// Descripción resumida de ModulosCO_2
    /// </summary>
    [TestClass]
    public class ModulosCO_2 : FuncionesVitales
    {



        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();

        public ModulosCO_2()
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
            string programa = "KFdPlcur";
            AbrirPrograma a = new AbrirPrograma(programa, "UCalida1");
            desktopSession = a.Start2("SQL", "");
            Thread.Sleep(200000000);
            List<string> crudData = new List<string>() { };
            string file = CrearDocumentoWordDinamico("prueba", "Pruebas", "SQL");


            string TipoQbe = "0";
            string QbeFilter = "2000";
            string BanderaSum = "1";
            string BanderaPreli = "0";
            string banExcel = "0";
            string codProgram = programa;
            string BanderaLupa = "1";
            string motor = "SQL";

            string modulo = "BP";
            string encuesta = "2000";
            string campoNombre = "PRUEBA";

            string escala = "12";
            string evento = "999";
            string programaLupa = "12";
            string editEvento = "8056";

            //string escala = "1";
            //string evento = "11";
            //string programaLupa = "1";
            //string editEvento = "10";

            string genero = "M";
            string numero = "6";
            string editNumero = "12";

            string idenDet2 = "78";
            string reaccion = "113";
            string editReaccion = "114";
            

            Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
            Dictionary<string, Point> botonesNavegadorDet1 = new Dictionary<string, Point>();
            Dictionary<string, Point> botonesNavegadorDet2 = new Dictionary<string, Point>();
            List<string> crudVariables = new List<string>() { escala, evento, programaLupa, modulo, campoNombre, encuesta, editEvento };
            List<string> crudVariablesDet1 = new List<string>() { genero, numero, editNumero };
            List<string> crudVariablesDet2 = new List<string>() { idenDet2, reaccion, editReaccion };
            bool isErrorDisplayed = false;


            //errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
            //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

            ////VALIDACION NOMBRE DEL PROGRAMA
            //errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
            //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

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
            CrudKcoparam.KCoParamCRUD(desktopSession, crudVariables, 0);
            Thread.Sleep(5000);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            //newFunctions_4.ScreenshotNuevo("Registro Agregado", true, file);
            ////Validacion Agregar
            ////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, identificacion, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

            ///////DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKcoparam.KCoParamCRUD(desktopSession, crudVariables, 1);
            newFunctions_4.changeWindow(desktopSession, "Población Objeto", 1);
            Thread.Sleep(1000);
            CrudKcoparam.CheckBoxes(desktopSession);
            newFunctions_4.changeWindow(desktopSession, "Encuesta", 0);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            //////newFunctions_4.ScreenshotNuevo("Edición Descartada", true, file);
            ////////Validacion Editar Descartar
            //////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, tipo, editCampo, c, ErrorValidacion, "No se edito correctamente", 1);

            /////ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKcoparam.KCoParamCRUD(desktopSession, crudVariables, 1);
            Thread.Sleep(1000);
            newFunctions_4.changeWindow(desktopSession, "Población Objeto", 1);
            CrudKcoparam.CheckBoxes(desktopSession);
            newFunctions_4.changeWindow(desktopSession, "Encuesta", 0);
            Thread.Sleep(1000);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

            if (isErrorDisplayed)
            {
                /////ACEPTAR EDICION
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                desktopSession.Mouse.Click(null);
                CrudKcoparam.KCoParamCRUD(desktopSession, crudVariables, 1);
                Thread.Sleep(1000);
                newFunctions_4.changeWindow(desktopSession, "Población Objeto", 1);
                CrudKcoparam.CheckBoxes(desktopSession);
                newFunctions_4.changeWindow(desktopSession, "Encuesta", 0);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
            }
            //////newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
            //////ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, identificacion, "E", c, ErrorValidacion);
            ////////Validacion Editar Aceptar
            //////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, editTipo, editCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);
            //////Thread.Sleep(1000);
            ////


            newFunctions_4.changeWindow(desktopSession, "Población Objeto", 1);
            botonesNavegadorDet1 = PruebaCRUD.findname(desktopSession, 25, 2);
            ElementList = PruebaCRUD.NavClass(desktopSession);
            Thread.Sleep(2000);

            //AGREGAR REGISTRO DETALLE
            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Insertar"].X, botonesNavegadorDet1["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKcoparam.KCoParamCRUDDet1(desktopSession, crudVariablesDet1, 0);
            Thread.Sleep(5000);
            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            newFunctions_4.ScreenshotNuevo("Registro Agregado Detalle", true, file);
            //////Validacion Agregar


            ///////DESCARTAR EDICION DETALLE
            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKcoparam.KCoParamCRUDDet1(desktopSession, crudVariablesDet1, 1);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Registro Editado Detalle", true, file);
            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Cancelar"].X, botonesNavegadorDet1["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            newFunctions_4.ScreenshotNuevo("Registro Descartado Detalle", true, file);

            /////ACEPTAR EDICION DETALLE
            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKcoparam.KCoParamCRUDDet1(desktopSession, crudVariablesDet1, 1);
            Thread.Sleep(1000);
            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

            if (isErrorDisplayed)
            {
                /////ACEPTAR EDICION DETALLE
                desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
                desktopSession.Mouse.Click(null);
                CrudKcoparam.KCoParamCRUDDet1(desktopSession, crudVariablesDet1, 1);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
            }
            newFunctions_4.ScreenshotNuevo("Registro Editado Detalle", true, file);


            botonesNavegadorDet2 = PruebaCRUD.findname(desktopSession, 25, 1);
            ElementList = PruebaCRUD.NavClass(desktopSession);

            //AGREGAR REGISTRO DETALLE 2
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Insertar"].X, botonesNavegadorDet2["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKcoparam.KCoParamCRUDDet2(desktopSession, crudVariablesDet2, 0);
            Thread.Sleep(5000);
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Aplicar"].X, botonesNavegadorDet2["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            //newFunctions_4.ScreenshotNuevo("Registro Agregado Detalle", true, file);
            //////Validacion Agregar
            ///
             ///////DESCARTAR EDICION DETALLE
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Editar"].X, botonesNavegadorDet2["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKcoparam.KCoParamCRUDDet2(desktopSession, crudVariablesDet2, 1);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Registro Editado Detalle", true, file);
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Cancelar"].X, botonesNavegadorDet2["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            newFunctions_4.ScreenshotNuevo("Registro Descartado Detalle", true, file);

            /////ACEPTAR EDICION DETALLE
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Editar"].X, botonesNavegadorDet2["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKcoparam.KCoParamCRUDDet2(desktopSession, crudVariablesDet2, 1);
            Thread.Sleep(1000);
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Aplicar"].X, botonesNavegadorDet2["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

            if (isErrorDisplayed)
            {
                /////ACEPTAR EDICION
                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Editar"].X, botonesNavegadorDet2["Editar"].Y);
                desktopSession.Mouse.Click(null);
                CrudKcoparam.KCoParamCRUDDet2(desktopSession, crudVariablesDet2, 1);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Aplicar"].X, botonesNavegadorDet2["Aplicar"].Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
            }
            newFunctions_4.ScreenshotNuevo("Registro Editado Detalle", true, file);


            //QBE
            errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

            //SUMATORIA
            errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            newFunctions_1.lapiz(desktopSession);

            Thread.Sleep(1000);
            //REPORTE DINAMICO
            string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
            errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            ////Rectificar Pie Pagina PDF DINAMICO
            //listaResultado = Textopdf(pdf, codProgram, user);
            //listaResultado.ForEach(e => ErrorValidacion.Add(e));

            //REPORTE PRELIMINAR
            string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar1" + "_" + Hora();
            newFunctions_4.ReportePreliminarOpciones(desktopSession, pdf1, "Reporte General", file, false);
            Thread.Sleep(3000);
            string pdf2 = @"C:\Reportes\ReportePDF1_" + "Preliminar2" + "_" + Hora();
            newFunctions_4.ReportePreliminarOpciones(desktopSession, pdf2, "Reporte Encuesta", file, false);
            Thread.Sleep(3000);
            string pdf3 = @"C:\Reportes\ReportePDF1_" + "Preliminar3" + "_" + Hora();
            newFunctions_4.ReportePreliminarOpciones(desktopSession, pdf3, "Reporte Población", file, false);
            Thread.Sleep(3000);
            string pdf4 = @"C:\Reportes\ReportePDF1_" + "Preliminar4" + "_" + Hora();
            newFunctions_4.ReportePreliminarIfnotOptions(desktopSession, pdf4, file, "Formulario de Encuesta", false);
            Thread.Sleep(3000);
            //string pdf5 = @"C:\Reportes\ReportePDF1_" + "Preliminar5" + "_" + Hora();
            //newFunctions_4.ReportePreliminarOpciones(desktopSession, pdf5, "Exportación Excel desde Formulario", file, false);
            //Thread.Sleep(3000);
            //////Rectificar Pie Pagina PDF PRELIMINAR
            //listaResultado = Textopdf(pdf1, codProgram, user);
            //listaResultado.ForEach(e => ErrorValidacion.Add(e));

            //EXPORTAR EXCEL
            string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
            errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

            //////RECTIFICACION DE CAMPOS DE EXCEL
            //string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, identificacion, c);
            //string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
            //Celda = Celda_Excel(ruta, nom_empr);
            //Celda.ForEach(u => ErrorValidacion.Add(u));
            //Thread.Sleep(1000);
            //LUPA
            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

            ////ELIMINAR REGISTRO DETALLE
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Borrar"].X, botonesNavegadorDet2["Borrar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
            Thread.Sleep(3000);
            WindowsDriver<WindowsElement> rootSession = newFunctions_4.RootSessionNew();
            Thread.Sleep(5000);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
            Thread.Sleep(1000);


            ////ELIMINAR REGISTRO DETALLE
            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Borrar"].X, botonesNavegadorDet1["Borrar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
            Thread.Sleep(3000);
            rootSession = newFunctions_4.RootSessionNew();
            Thread.Sleep(5000);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
            Thread.Sleep(1000);
            isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

            if (isErrorDisplayed)
            {
                desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Borrar"].X, botonesNavegadorDet1["Borrar"].Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
                Thread.Sleep(3000);
                rootSession = newFunctions_4.RootSessionNew();
                Thread.Sleep(5000);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
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


        }


        [TestMethod]
        public void KCoRcoenCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosCO.ModulosCO_2.KCoRcoenCheckList")
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
                                rows["EncuestaBase"].ToString().Length != 0 && rows["EncuestaBase"].ToString() != null &&
                                rows["EncuestaComparacion"].ToString().Length != 0 && rows["EncuestaComparacion"].ToString() != null &&
                                rows["BotonReporte"].ToString().Length != 0 && rows["BotonReporte"].ToString() != null &&
                                rows["TabNumber"].ToString().Length != 0 && rows["TabNumber"].ToString() != null
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
                                string encuestaBase = rows["EncuestaBase"].ToString();
                                string encuestaComparacion = rows["EncuestaComparacion"].ToString();
                                string botonReporte = rows["BotonReporte"].ToString();
                                string tabnumber = rows["TabNumber"].ToString();
                                List<string> lupasData = new List<string>() { encuestaBase, encuestaComparacion };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "CO", motor);
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

                                        //VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //DATOS
                                        CrudKcorcoen.KCoRcoenCRUD(desktopSession, Int32.Parse(tabnumber), botonReporte, lupasData);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Datos Necesarios", true, file);
                                        Thread.Sleep(1000);

                                        //REPORTE PRELIMINAR
                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        ////Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));
                                        //ClickWindow
                                        Thread.Sleep(1000);
                                        CrudKcorcoen.ClickWindow(desktopSession);
                                        Thread.Sleep(1000);

                                        //IMPRIMIR
                                        string pdf2 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        newFunctions_3.BotonImprimir(desktopSession, file);

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
        public void KCoParamCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosCO.ModulosCO_2.KCoParamCheckList")
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
                                //rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                rows["Modulo"].ToString().Length != 0 && rows["Modulo"].ToString() != null &&
                                rows["Encuesta"].ToString().Length != 0 && rows["Encuesta"].ToString() != null &&
                                rows["CampoNombre"].ToString().Length != 0 && rows["CampoNombre"].ToString() != null &&
                                rows["Escala"].ToString().Length != 0 && rows["Escala"].ToString() != null &&
                                rows["Evento"].ToString().Length != 0 && rows["Evento"].ToString() != null &&
                                rows["ProgramaLupa"].ToString().Length != 0 && rows["ProgramaLupa"].ToString() != null &&
                                rows["EditEvento"].ToString().Length != 0 && rows["EditEvento"].ToString() != null &&
                                rows["Genero"].ToString().Length != 0 && rows["Genero"].ToString() != null &&
                                rows["Numero"].ToString().Length != 0 && rows["Numero"].ToString() != null &&
                                rows["EditNumero"].ToString().Length != 0 && rows["EditNumero"].ToString() != null &&
                                rows["IdentDet2"].ToString().Length != 0 && rows["IdentDet2"].ToString() != null &&
                                rows["Reaccion"].ToString().Length != 0 && rows["Reaccion"].ToString() != null &&
                                rows["EditReaccion"].ToString().Length != 0 && rows["EditReaccion"].ToString() != null
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


                                string modulo = rows["Modulo"].ToString();
                                string encuesta = rows["Encuesta"].ToString();
                                string campoNombre = rows["CampoNombre"].ToString();

                                string escala = rows["Escala"].ToString();
                                string evento = rows["Evento"].ToString();
                                string programaLupa = rows["ProgramaLupa"].ToString();
                                string editEvento = rows["EditEvento"].ToString();

                                //string escala = "1";
                                //string evento = "11";
                                //string programaLupa = "1";
                                //string editEvento = "10";

                                string genero = rows["Genero"].ToString();
                                string numero = rows["Numero"].ToString();
                                string editNumero = rows["EditNumero"].ToString();

                                string idenDet2 = rows["IdentDet2"].ToString();
                                string reaccion = rows["Reaccion"].ToString();
                                string editReaccion = rows["EditReaccion"].ToString();
                                string identificacion = encuesta;

                                Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet1 = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet2 = new Dictionary<string, Point>();
                                List<string> crudVariables = new List<string>() { escala, evento, programaLupa, modulo, campoNombre, encuesta, editEvento };
                                List<string> crudVariablesDet1 = new List<string>() { genero, numero, editNumero };
                                List<string> crudVariablesDet2 = new List<string>() { idenDet2, reaccion, editReaccion };
                                bool isErrorDisplayed = false;


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "CO", motor);
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

                                        //VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, identificacion, Campo, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

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
                                        CrudKcoparam.KCoParamCRUD(desktopSession, crudVariables, 0);
                                        Thread.Sleep(5000);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Registro Agregado", true, file);
                                        //Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, identificacion, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ///////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKcoparam.KCoParamCRUD(desktopSession, crudVariables, 1);
                                        newFunctions_4.changeWindow(desktopSession, "Población Objeto", 1);
                                        Thread.Sleep(1000);
                                        CrudKcoparam.CheckBoxes(desktopSession);
                                        newFunctions_4.changeWindow(desktopSession, "Encuesta", 0);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada", true, file);
                                        //Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, evento, editCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        /////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKcoparam.KCoParamCRUD(desktopSession, crudVariables, 1);
                                        Thread.Sleep(1000);
                                        newFunctions_4.changeWindow(desktopSession, "Población Objeto", 1);
                                        CrudKcoparam.CheckBoxes(desktopSession);
                                        newFunctions_4.changeWindow(desktopSession, "Encuesta", 0);
                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

                                        if (isErrorDisplayed)
                                        {
                                            /////ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKcoparam.KCoParamCRUD(desktopSession, crudVariables, 1);
                                            Thread.Sleep(1000);
                                            newFunctions_4.changeWindow(desktopSession, "Población Objeto", 1);
                                            CrudKcoparam.CheckBoxes(desktopSession);
                                            newFunctions_4.changeWindow(desktopSession, "Encuesta", 0);
                                            Thread.Sleep(1000);
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(2000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, identificacion, "E", c, ErrorValidacion);
                                        //Validacion Editar Aceptar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, editEvento, editCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);
                                        Thread.Sleep(1000);



                                        newFunctions_4.changeWindow(desktopSession, "Población Objeto", 1);
                                        botonesNavegadorDet1 = PruebaCRUD.findname(desktopSession, 25, 2);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        Thread.Sleep(2000);

                                        //AGREGAR REGISTRO DETALLE
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Insertar"].X, botonesNavegadorDet1["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKcoparam.KCoParamCRUDDet1(desktopSession, crudVariablesDet1, 0);
                                        Thread.Sleep(5000);
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Registro Agregado Detalle", true, file);
                                        //////Validacion Agregar


                                        ///////DESCARTAR EDICION DETALLE
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKcoparam.KCoParamCRUDDet1(desktopSession, crudVariablesDet1, 1);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Registro Editado Detalle", true, file);
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Cancelar"].X, botonesNavegadorDet1["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Registro Descartado Detalle", true, file);

                                        /////ACEPTAR EDICION DETALLE
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKcoparam.KCoParamCRUDDet1(desktopSession, crudVariablesDet1, 1);
                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

                                        if (isErrorDisplayed)
                                        {
                                            /////ACEPTAR EDICION DETALLE
                                            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKcoparam.KCoParamCRUDDet1(desktopSession, crudVariablesDet1, 1);
                                            Thread.Sleep(1000);
                                            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(2000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Registro Editado Detalle", true, file);


                                        if (motor == "SQL") {
                                            botonesNavegadorDet2 = PruebaCRUD.findname(desktopSession, 25, 1);
                                            ElementList = PruebaCRUD.NavClass(desktopSession);

                                            //AGREGAR REGISTRO DETALLE 2
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Insertar"].X, botonesNavegadorDet2["Insertar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKcoparam.KCoParamCRUDDet2(desktopSession, crudVariablesDet2, 0);
                                            Thread.Sleep(5000);
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Aplicar"].X, botonesNavegadorDet2["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(2000);
                                            newFunctions_4.ScreenshotNuevo("Registro Agregado Detalle 2 ", true, file);
                                            //////Validacion Agregar
                                            ///
                                            ///////DESCARTAR EDICION DETALLE
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Editar"].X, botonesNavegadorDet2["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKcoparam.KCoParamCRUDDet2(desktopSession, crudVariablesDet2, 1);
                                            Thread.Sleep(1000);
                                            newFunctions_4.ScreenshotNuevo("Registro Editado Detalle2 ", true, file);
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Cancelar"].X, botonesNavegadorDet2["Cancelar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(2000);
                                            newFunctions_4.ScreenshotNuevo("Registro Descartado Detalle 2", true, file);

                                            /////ACEPTAR EDICION DETALLE
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Editar"].X, botonesNavegadorDet2["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKcoparam.KCoParamCRUDDet2(desktopSession, crudVariablesDet2, 1);
                                            Thread.Sleep(1000);
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Aplicar"].X, botonesNavegadorDet2["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(2000);
                                            isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

                                            if (isErrorDisplayed)
                                            {
                                                /////ACEPTAR EDICION
                                                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Editar"].X, botonesNavegadorDet2["Editar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                CrudKcoparam.KCoParamCRUDDet2(desktopSession, crudVariablesDet2, 1);
                                                Thread.Sleep(1000);
                                                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Aplicar"].X, botonesNavegadorDet2["Aplicar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                Thread.Sleep(2000);
                                            }
                                            newFunctions_4.ScreenshotNuevo("Registro Editado Detalle 2", true, file);
                                        }

                                      
                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_1.lapiz(desktopSession);

                                        Thread.Sleep(1000);
                                        //REPORTE DINAMICO
                                        string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
                                        errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF DINAMICO
                                        listaResultado = Textopdf(pdf, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //REPORTE PRELIMINAR
                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar1" + "_" + Hora();
                                        newFunctions_4.ReportePreliminarOpciones(desktopSession, pdf1, "Reporte General", file, false);
                                        Thread.Sleep(3000);
                                        string pdf2 = @"C:\Reportes\ReportePDF1_" + "Preliminar2" + "_" + Hora();
                                        newFunctions_4.ReportePreliminarOpciones(desktopSession, pdf2, "Reporte Encuesta", file, false);
                                        Thread.Sleep(3000);
                                        string pdf3 = @"C:\Reportes\ReportePDF1_" + "Preliminar3" + "_" + Hora();
                                        newFunctions_4.ReportePreliminarOpciones(desktopSession, pdf3, "Reporte Población", file, false);
                                        if (motor == "SQL") {
                                            Thread.Sleep(3000);
                                            string pdf4 = @"C:\Reportes\ReportePDF1_" + "Preliminar4" + "_" + Hora();
                                            newFunctions_4.ReportePreliminarIfnotOptions(desktopSession, pdf4, file, "Formulario de Encuesta", false);
                                            Thread.Sleep(3000);
                                        }
                                        //string pdf5 = @"C:\Reportes\ReportePDF1_" + "Preliminar5" + "_" + Hora();
                                        //newFunctions_4.ReportePreliminarOpciones(desktopSession, pdf5, "Exportación Excel desde Formulario", file, false);
                                        //Thread.Sleep(3000);
                                        ////Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));
                                        Thread.Sleep(1000);
                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////RECTIFICACION DE CAMPOS DE EXCEL
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, identificacion, c);
                                        //string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        //Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));
                                        Thread.Sleep(1000);
                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        WindowsDriver<WindowsElement> rootSession = null;
                                        if (motor == "SQL")
                                        {
                                            ////ELIMINAR REGISTRO DETALLE 2
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Borrar"].X, botonesNavegadorDet2["Borrar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(500);
                                            newFunctions_4.ScreenshotNuevo("Borrar Registro Detalle 2", true, file);
                                            Thread.Sleep(3000);
                                            rootSession = newFunctions_4.RootSessionNew();
                                            Thread.Sleep(5000);
                                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                            Thread.Sleep(500);
                                            newFunctions_4.ScreenshotNuevo("Registro Borrado Detalle 2", true, file);
                                            Thread.Sleep(1000);
                                        }

                                        ////ELIMINAR REGISTRO DETALLE 1
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Borrar"].X, botonesNavegadorDet1["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Borrar Registro Detalle 1", true, file);
                                        Thread.Sleep(3000);
                                        rootSession = newFunctions_4.RootSessionNew();
                                        Thread.Sleep(5000);
                                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Registro Borrado Detalle 1", true, file);
                                        Thread.Sleep(1000);
                                        isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

                                        if (isErrorDisplayed)
                                        {
                                            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDet1["Borrar"].X, botonesNavegadorDet1["Borrar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(500);
                                            newFunctions_4.ScreenshotNuevo("Borrar Registro Detalle 1", true, file);
                                            Thread.Sleep(3000);
                                            rootSession = newFunctions_4.RootSessionNew();
                                            Thread.Sleep(5000);
                                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                            Thread.Sleep(500);
                                            newFunctions_4.ScreenshotNuevo("Registro Borrado Detalle 1", true, file);
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

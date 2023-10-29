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

namespace Ophelia_Test_V21.TestMethods.ModulosSO
{
    /// <summary>
    /// Descripción resumida de ModulosSO_8
    /// </summary>
    [TestClass]
    public class ModulosSO_8 : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();
        public ModulosSO_8()
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
        public void KSoResolCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_8.KSoResolCheckList")
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
                                rows["Secuencial"].ToString().Length != 0 && rows["Secuencial"].ToString() != null &&
                                rows["Codigo"].ToString().Length != 0 && rows["Codigo"].ToString() != null &&
                                rows["Color"].ToString().Length != 0 && rows["Color"].ToString() != null &&
                                rows["EditColor"].ToString().Length != 0 && rows["EditColor"].ToString() != null &&
                                rows["PlanAccion"].ToString().Length != 0 && rows["PlanAccion"].ToString() != null &&
                                rows["TipoMedida"].ToString().Length != 0 && rows["TipoMedida"].ToString() != null &&
                                rows["Actividad"].ToString().Length != 0 && rows["Actividad"].ToString() != null &&

                                rows["Clase"].ToString().Length != 0 && rows["Clase"].ToString() != null &&
                                rows["Seccional"].ToString().Length != 0 && rows["Seccional"].ToString() != null &&
                                rows["Dependencia"].ToString().Length != 0 && rows["Dependencia"].ToString() != null &&
                                rows["CentrodeCostos"].ToString().Length != 0 && rows["CentrodeCostos"].ToString() != null &&
                                rows["EditCentrodeCostos"].ToString().Length != 0 && rows["EditCentrodeCostos"].ToString() != null &&



                                rows["EditTipoMedida"].ToString().Length != 0 && rows["EditTipoMedida"].ToString() != null
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

                                string secuencial = rows["Secuencial"].ToString();
                                string codigo = rows["Codigo"].ToString();
                                string color = rows["Color"].ToString();
                                string editColor = rows["EditColor"].ToString();

                                string Clase = rows["Clase"].ToString();
                                string  Seccional = rows["Seccional"].ToString();
                                string  Dependencia = rows["Dependencia"].ToString();
                                string  CentrodeCostos = rows["CentrodeCostos"].ToString();
                                string EditCentrodeCostos = rows["EditCentrodeCostos"].ToString();

                                string planAccion = rows["PlanAccion"].ToString();
                                string tipoMedida = rows["TipoMedida"].ToString();
                                string actividad = rows["Actividad"].ToString();
                                string editTipoMedida = rows["EditTipoMedida"].ToString();
                                string identificacion = secuencial;

                                Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet1 = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet2 = new Dictionary<string, Point>();
                                List<string> crudVariables = new List<string>() {Clase, Seccional, Dependencia, CentrodeCostos, EditCentrodeCostos };
                                List<string> crudVariablesDet1 = new List<string>() { planAccion, tipoMedida, actividad, editTipoMedida };
                                bool isErrorDisplayed = false;

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);
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
                                        CrudKsoresol.InsertarDatos(desktopSession, 0, crudVariables, file);
                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Registro Agregado", true, file);
                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, identificacion, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ///////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoresol.InsertarDatos(desktopSession, 1, crudVariables, file);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada", true, file);
                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, color, editCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoresol.InsertarDatos(desktopSession, 1, crudVariables, file);
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
                                            CrudKsoresol.KSoResolCRUD(desktopSession, crudVariables, 1);
                                            Thread.Sleep(1000);
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(2000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, identificacion, "E", c, ErrorValidacion);
                                        ////Validacion Editar Aceptar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, editColor, editCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);
                                        //Thread.Sleep(1000);


                                        //newFunctions_4.changeWindow(desktopSession, "Cronograma General de Actividades");
                                        //Thread.Sleep(1000);
                                        //botonesNavegadorDet1 = PruebaCRUD.findname(desktopSession, 25, 1);
                                        //ElementList = PruebaCRUD.NavClass(desktopSession);

                                        ////AGREGAR REGISTRO DETALLE
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Insertar"].X, botonesNavegadorDet1["Insertar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //CrudKsoresol.KSoResolCRUDDet1(desktopSession, crudVariablesDet1, 0);
                                        //Thread.Sleep(4000);
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //Thread.Sleep(2000);
                                        //newFunctions_4.ScreenshotNuevo("Registro Agregado Detalle", true, file);

                                        /////////DESCARTAR EDICION DETALLE
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //CrudKsoresol.KSoResolCRUDDet1(desktopSession, crudVariablesDet1, 1);
                                        //Thread.Sleep(1000);
                                        //newFunctions_4.ScreenshotNuevo("Registro Editado Detalle", true, file);
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Cancelar"].X, botonesNavegadorDet1["Cancelar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //Thread.Sleep(2000);
                                        //newFunctions_4.ScreenshotNuevo("Edición Descartada Detalle", true, file);

                                        ////ACEPTAR EDICION DETALLE
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //CrudKsoresol.KSoResolCRUDDet1(desktopSession, crudVariablesDet1, 1);
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
                                        //    CrudKsoresol.KSoResolCRUDDet1(desktopSession, crudVariablesDet1, 1);
                                        //    Thread.Sleep(1000);
                                        //    desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    Thread.Sleep(2000);
                                        //}
                                        //newFunctions_4.ScreenshotNuevo("Registro Editado Detalle", true, file);


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
        public void KSoSeydeCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_8.KSoSeydeCheckList")
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
                                rows["Seccional"].ToString().Length != 0 && rows["Seccional"].ToString() != null &&
                                rows["FechaElaboracion"].ToString().Length != 0 && rows["FechaElaboracion"].ToString() != null &&
                                rows["Sede"].ToString().Length != 0 && rows["Sede"].ToString() != null &&
                                rows["EditSede"].ToString().Length != 0 && rows["EditSede"].ToString() != null &&
                                rows["IdExtintor"].ToString().Length != 0 && rows["IdExtintor"].ToString() != null &&
                                rows["Piso"].ToString().Length != 0 && rows["Piso"].ToString() != null &&
                                rows["Ubicacion"].ToString().Length != 0 && rows["Ubicacion"].ToString() != null &&
                                rows["FechaRecarga"].ToString().Length != 0 && rows["FechaRecarga"].ToString() != null &&
                                rows["FechaVencimiento"].ToString().Length != 0 && rows["FechaVencimiento"].ToString() != null &&
                                rows["EditFechaVencimiento"].ToString().Length != 0 && rows["EditFechaVencimiento"].ToString() != null &&
                                rows["FechaInicio"].ToString().Length != 0 && rows["FechaInicio"].ToString() != null &&
                                rows["TipoAccion"].ToString().Length != 0 && rows["TipoAccion"].ToString() != null &&
                                rows["FechaCumplimiento"].ToString().Length != 0 && rows["FechaCumplimiento"].ToString() != null &&
                                rows["Responsable"].ToString().Length != 0 && rows["Responsable"].ToString() != null &&
                                rows["FechaSeguimiento"].ToString().Length != 0 && rows["FechaSeguimiento"].ToString() != null &&
                                rows["FechaCierre"].ToString().Length != 0 && rows["FechaCierre"].ToString() != null &&
                                rows["Estado"].ToString().Length != 0 && rows["Estado"].ToString() != null &&
                                rows["EditEstado"].ToString().Length != 0 && rows["EditEstado"].ToString() != null
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

                                string seccional = rows["Seccional"].ToString();
                                string fechaElaboracion = rows["FechaElaboracion"].ToString();
                                string sede = rows["Sede"].ToString();
                                string editSede = rows["EditSede"].ToString();
                                string identificacion = seccional;

                                string IdExtintor = rows["IdExtintor"].ToString();
                                string piso = rows["Piso"].ToString();
                                string ubicacion = rows["Ubicacion"].ToString();
                                string fechaRecarga = rows["FechaRecarga"].ToString();
                                string fechaVencimiento = rows["FechaVencimiento"].ToString();
                                string editFechaVencimiento = rows["EditFechaVencimiento"].ToString();

                                string fechaInicio = rows["FechaInicio"].ToString();
                                string tipoAccion = rows["TipoAccion"].ToString();
                                string fechaCumplimiento = rows["FechaCumplimiento"].ToString();
                                string responsable = rows["Responsable"].ToString();
                                string fechaSeguimiento = rows["FechaSeguimiento"].ToString();
                                string fechaCierre = rows["FechaCierre"].ToString();
                                string estado = rows["Estado"].ToString();
                                string editEstado = rows["EditEstado"].ToString();

                                Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet1 = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet2 = new Dictionary<string, Point>();
                                List<string> crudVariables = new List<string>() { sede, seccional, fechaElaboracion, editSede };
                                List<string> crudVariablesDet1 = new List<string>() { IdExtintor, piso, ubicacion, fechaRecarga, fechaVencimiento, editFechaVencimiento };
                                List<string> crudVariablesDet2 = new List<string>() { fechaInicio, tipoAccion, fechaCumplimiento, responsable, fechaSeguimiento, fechaCierre, estado, editEstado };
                                WindowsDriver<WindowsElement> rootSession = null;
                                bool isErrorDisplayed = false;

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);
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

                                        try
                                        {
                                            //AGREGAR REGISTRO
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                            CrudKsoseyde.KSoseydeCRUD(desktopSession, crudVariables, 0);
                                            Thread.Sleep(4000);
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(2000);
                                            newFunctions_4.ScreenshotNuevo("Registro Agregado", true, file);
                                           
                                        }
                                        catch
                                        {
                                            //AGREGAR REGISTRO
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                            CrudKsoseyde.KSoseydeCRUD(desktopSession, crudVariables, 0);
                                            Thread.Sleep(4000);
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(2000);
                                            newFunctions_4.ScreenshotNuevo("Registro Agregado", true, file);
                                            
                                        }
                                        //Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, identificacion, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);


                                        ///////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoseyde.KSoseydeCRUD(desktopSession, crudVariables, 1);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada", true, file);
                                        //Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, sede, editCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsoseyde.KSoseydeCRUD(desktopSession, crudVariables, 1);
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
                                            CrudKsoseyde.KSoseydeCRUD(desktopSession, crudVariables, 1);
                                            Thread.Sleep(1000);
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(2000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, identificacion, "E", c, ErrorValidacion);
                                        //Validacion Editar Aceptar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, editSede, editCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);
                                        Thread.Sleep(1000);

                                        ///LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        newFunctions_4.changeWindow(desktopSession, "Inspecciones de Extintores");
                                        Thread.Sleep(1000);
                                        botonesNavegadorDet1 = PruebaCRUD.findname(desktopSession, 25, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);

                                        ////AGREGAR REGISTRO DETALLE
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Insertar"].X, botonesNavegadorDet1["Insertar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //CrudKsoseyde.KSoseydeCRUDDet1(desktopSession, crudVariablesDet1, 0);
                                        //Thread.Sleep(4000);
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //Thread.Sleep(2000);
                                        //newFunctions_4.ScreenshotNuevo("Registro Agregado Detalle", true, file);

                                        /////////DESCARTAR EDICION DETALLE
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //CrudKsoseyde.KSoseydeCRUDDet1(desktopSession, crudVariablesDet1, 1);
                                        //Thread.Sleep(1000);
                                        //newFunctions_4.ScreenshotNuevo("Registro Editado Detalle", true, file);
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Cancelar"].X, botonesNavegadorDet1["Cancelar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //Thread.Sleep(2000);
                                        //newFunctions_4.ScreenshotNuevo("Edición Descartada Detalle", true, file);

                                        ////ACEPTAR EDICION DETALLE
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //CrudKsoseyde.KSoseydeCRUDDet1(desktopSession, crudVariablesDet1, 1);
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
                                        //    CrudKsoseyde.KSoseydeCRUDDet1(desktopSession, crudVariablesDet1, 1);
                                        //    Thread.Sleep(1000);
                                        //    desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    Thread.Sleep(2000);
                                        //}
                                        //newFunctions_4.ScreenshotNuevo("Registro Editado Detalle", true, file);

                                        ////DETALLE 2
                                        //newFunctions_4.changeWindow(desktopSession, "Plan de Acción");
                                        //Thread.Sleep(1000);
                                        //botonesNavegadorDet2 = PruebaCRUD.findname(desktopSession, 25, 1);
                                        //ElementList = PruebaCRUD.NavClass(desktopSession);

                                        ////AGREGAR REGISTRO DETALLE 2
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Insertar"].X, botonesNavegadorDet2["Insertar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //CrudKsoseyde.KSoseydeCRUDDet2(desktopSession, crudVariablesDet2, 0);
                                        //Thread.Sleep(4000);
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Aplicar"].X, botonesNavegadorDet2["Aplicar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //Thread.Sleep(2000);
                                        //newFunctions_4.ScreenshotNuevo("Registro Agregado Detalle 2", true, file);

                                        /////////DESCARTAR EDICION DETALLE 2
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Editar"].X, botonesNavegadorDet2["Editar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //CrudKsoseyde.KSoseydeCRUDDet1(desktopSession, crudVariablesDet2, 1);
                                        //Thread.Sleep(1000);
                                        //newFunctions_4.ScreenshotNuevo("Registro Editado Detalle 2", true, file);
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Cancelar"].X, botonesNavegadorDet2["Cancelar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //Thread.Sleep(2000);
                                        //newFunctions_4.ScreenshotNuevo("Edición Descartada Detalle 2", true, file);

                                        ////ACEPTAR EDICION DETALLE 2
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Editar"].X, botonesNavegadorDet2["Editar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //CrudKsoseyde.KSoseydeCRUDDet1(desktopSession, crudVariablesDet2, 1);
                                        //Thread.Sleep(1000);
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Aplicar"].X, botonesNavegadorDet2["Aplicar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //Thread.Sleep(2000);
                                        //isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

                                        //if (isErrorDisplayed)
                                        //{
                                        //    ////ACEPTAR EDICION DETALLE 2
                                        //    desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Editar"].X, botonesNavegadorDet2["Editar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    CrudKsoseyde.KSoseydeCRUDDet1(desktopSession, crudVariablesDet2, 1);
                                        //    Thread.Sleep(1000);
                                        //    desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Aplicar"].X, botonesNavegadorDet2["Aplicar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    Thread.Sleep(2000);
                                        //}
                                        //newFunctions_4.ScreenshotNuevo("Registro Editado Detalle 2", true, file);

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
                                        ////Rectificar Pie Pagina PDF DINAMICO
                                        listaResultado = Textopdf(pdf, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //REPORTE PRELIMINAR
                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = CrudKsoseyde.ReportePreliminarIfnotOptions(desktopSession, pdf1, file, "Reporte de Señalización", false);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //REPORTE PRELIMINAR
                                        string pdf2 = @"C:\Reportes\ReportePDF1_" + "Preliminar2" + "_" + Hora();
                                        errors = CrudKsoseyde.ReportePreliminarIfnotOptions(desktopSession, pdf2, file, "Reporte de Extintores", false);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf2, codProgram, user);
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

                                        //////ELIMINAR REGISTRO DETALLE 2
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Borrar"].X, botonesNavegadorDet2["Borrar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //Thread.Sleep(500);
                                        //newFunctions_4.ScreenshotNuevo("Borrar Registro Detalle 2", true, file);
                                        //PruebaCRUD.cerrarBorrar(desktopSession);
                                        //Thread.Sleep(500);
                                        //newFunctions_4.ScreenshotNuevo("Registro Borrado Detalle 2", true, file);
                                        //Thread.Sleep(1000);

                                        //isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

                                        //if (isErrorDisplayed)
                                        //{
                                        //    ////ELIMINAR REGISTRO DETALLE 2
                                        //    desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet2["Borrar"].X, botonesNavegadorDet2["Borrar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    Thread.Sleep(500);
                                        //    newFunctions_4.ScreenshotNuevo("Borrar Registro Detalle 2", true, file);
                                        //    PruebaCRUD.cerrarBorrar(desktopSession);
                                        //    Thread.Sleep(500);
                                        //    newFunctions_4.ScreenshotNuevo("Registro Borrado Detalle 2", true, file);
                                        //    Thread.Sleep(1000);
                                        //}

                                        //newFunctions_4.changeWindow(desktopSession, "Inspecciones de Extintores");
                                        //Thread.Sleep(1000);
                                        //botonesNavegadorDet1 = PruebaCRUD.findname(desktopSession, 25, 1);
                                        //ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //////ELIMINAR REGISTRO 1
                                        //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Borrar"].X, botonesNavegadorDet1["Borrar"].Y);
                                        //desktopSession.Mouse.Click(null);
                                        //Thread.Sleep(500);
                                        //newFunctions_4.ScreenshotNuevo("Borrar Registro Detalle 1", true, file);
                                        //rootSession = PruebaCRUD.RootSession();
                                        //Thread.Sleep(5000);
                                        //rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                        //Thread.Sleep(2000);
                                        //Thread.Sleep(500);
                                        //newFunctions_4.ScreenshotNuevo("Registro Borrado Detalle 1", true, file);
                                        //Thread.Sleep(1000);

                                        //isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

                                        //if (isErrorDisplayed)
                                        //{
                                        //    ////ELIMINAR REGISTRO 1
                                        //    desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Borrar"].X, botonesNavegadorDet1["Borrar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    Thread.Sleep(500);
                                        //    newFunctions_4.ScreenshotNuevo("Borrar Registro Detalle 1", true, file);
                                        //    rootSession = PruebaCRUD.RootSession();
                                        //    Thread.Sleep(5000);
                                        //    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                        //    Thread.Sleep(2000);
                                        //    Thread.Sleep(500);
                                        //    newFunctions_4.ScreenshotNuevo("Registro Borrado Detalle 1", true, file);
                                        //    Thread.Sleep(1000);
                                        //}


                                        ////ELIMINAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
                                        rootSession = PruebaCRUD.RootSession();
                                        Thread.Sleep(5000);
                                        PruebaCRUD.cerrarBorrar(rootSession);
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
                                            rootSession = PruebaCRUD.RootSession();
                                            Thread.Sleep(5000);
                                            PruebaCRUD.cerrarBorrar(rootSession);
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
        public void KSopamalCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_8.KSopamalCheckList")
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
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&
                                //Datos Necesidades Formación
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                //Datos Crud principal
                                rows["Ciclo"].ToString().Length != 0 && rows["Ciclo"].ToString() != null &&
                                rows["EdFechaI"].ToString().Length != 0 && rows["EdFechaI"].ToString() != null &&
                                //Datos Crud detalle 1
                                rows["Descripcion"].ToString().Length != 0 && rows["Descripcion"].ToString() != null &&
                                rows["Porcentaje"].ToString().Length != 0 && rows["Porcentaje"].ToString() != null &&
                                //Datos Crud detalle 2
                                rows["Descrip"].ToString().Length != 0 && rows["Descrip"].ToString() != null &&
                                rows["Porcen"].ToString().Length != 0 && rows["Porcen"].ToString() != null &&
                                //Datos Crud detalle 3
                                rows["Numeral"].ToString().Length != 0 && rows["Numeral"].ToString() != null &&
                                rows["Descripc"].ToString().Length != 0 && rows["Descripc"].ToString() != null &&
                                rows["Valor"].ToString().Length != 0 && rows["Valor"].ToString() != null
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

                                string Ciclo = rows["Ciclo"].ToString();
                                string EdFechaI = rows["EdFechaI"].ToString();

                                string Descripcion = rows["Descripcion"].ToString();
                                string Porcentaje = rows["Porcentaje"].ToString();

                                string Descrip = rows["Descrip"].ToString();
                                string Porcen = rows["Porcen"].ToString();

                                string Numeral = rows["Numeral"].ToString();
                                string Descripc = rows["Descripc"].ToString();
                                string Valor = rows["Valor"].ToString();

                                String campo = "1";

                                Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                
                                Dictionary<string, Point> botonesNavegadorDet2 = new Dictionary<string, Point>();
                                Dictionary<string, Point> botonesNavegadorDet3 = new Dictionary<string, Point>();
                                List<string> crudPrincipal = new List<string>() { Ciclo, EdFechaI };
                                List<string> crudDet1 = new List<string>() { Descripcion, Porcentaje};
                                List<string> crudDet2 = new List<string>() { Descrip, Porcen };
                                List<string> crudDet3 = new List<string>() { Numeral, Descripc, Valor };
                                bool isErrorDisplayed = false;

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);
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

                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 27, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //AGREGAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopamal.AgregarCrud(desktopSession, 0, crudPrincipal);
                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(3000);
                                        CrudKsopamal.AgregarCrud(desktopSession, 3, crudPrincipal);
                                        Thread.Sleep(1000);
                                        //WindowsDriver<WindowsElement> RootSession = null;
                                        //RootSession = PruebaCRUD.RootSession();
                                        //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                        //Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Registro Agregado", true, file);
                                        //Validacion Agregar

                                        ///////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopamal.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Edición Descartada", true, file);
                                        //Validacion Editar Descartar
                                        
                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopamal.AgregarCrud(desktopSession, 1, crudPrincipal);
                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        //isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

                                        //if (isErrorDisplayed)
                                        //{
                                        //    ////ACEPTAR EDICION
                                        //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    CrudKsoresol.KSoResolCRUD(desktopSession, crudVariables, 1);
                                        //    Thread.Sleep(1000);
                                        //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        //    desktopSession.Mouse.Click(null);
                                        //    Thread.Sleep(2000);
                                        //}
                                        newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, identificacion, "E", c, ErrorValidacion);
                                        ////Validacion Editar Aceptar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, editColor, editCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);
                                        //Thread.Sleep(1000);

                                        //////Validacion Editar
                                        Thread.Sleep(1000);
                                        //// //Debugger.Launch();
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, ActCargo, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //LUPA
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////AGREGAR REGISTRO DETALLE 1
                                        
                                        CrudKsopamal.AgregarCrud1(desktopSession, crudDet1);
                                        Thread.Sleep(4000);
                                        newFunctions_4.ScreenshotNuevo("Registro Agregado Detalle 1", true, file);

                                       

                                        //AGREGAR REGISTRO DETALLE 2

                                        CrudKsopamal.AgregarCrud2(desktopSession, crudDet2);
                                        Thread.Sleep(4000);
                                        newFunctions_4.ScreenshotNuevo("Registro Agregado Detalle 2", true, file);

                                        ////AGREGAR REGISTRO DETALLE 3
                                        
                                        Dictionary<string, Point> botonesNavegadorDet1 = new Dictionary<string, Point>();
                                        botonesNavegadorDet1 = PruebaCRUD.findname(desktopSession, 28, 1);
                                        var ElementList2 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDet1["Insertar"].X, botonesNavegadorDet1["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKsopamal.AgregarCrud3(desktopSession, crudDet3);
                                        Thread.Sleep(4000);
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Registro Agregado Detalle 1", true, file);



                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file, campo);
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

                                        ////ELIMINAR REGISTRO

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDet1["Borrar"].X, botonesNavegadorDet1["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Borrar Registro Detalle 3", true, file);
                                        PruebaCRUD.cerrarBorrar(desktopSession);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Registro Borrado Detalle 3", true, file);
                                        Thread.Sleep(1000);

                                        ////ELIMINAR REGISTRO

                                        CrudKsopamal.EliminarDetalle2(desktopSession, file);
                                        newFunctions_4.ScreenshotNuevo("Registro Borrado Detalle 2", true, file);

                                        ////ELIMINAR REGISTRO

                                        CrudKsopamal.EliminarDetalle1(desktopSession, file);
                                        newFunctions_4.ScreenshotNuevo("Registro Borrado Detalle 1", true, file);

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

                                        ////Validacion Eliminar Registro
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, "", Campo, c, ErrorValidacion, "los datos no se encontraron correctamente", 2);
                                        ////CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, identificacion, "D", c, ErrorValidacion);

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

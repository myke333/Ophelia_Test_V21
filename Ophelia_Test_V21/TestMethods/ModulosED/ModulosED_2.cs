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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosEd;

namespace Ophelia_Test_V21.TestMethods.ModulosED
{
    /// <summary>
    /// Descripción resumida de ModulosED_2
    /// </summary>
    [TestClass]
    public class ModulosED_2 : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();
        public ModulosED_2()
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
            string programa = "KNmIncap";
            AbrirPrograma a = new
            AbrirPrograma(programa, "UCalida1");
            desktopSession = a.Start2("SQL", "");
            Thread.Sleep(2000000);
            List<string> crudData = new List<string>() { };
            string file = CrearDocumentoWordDinamico("prueba", "Pruebas", "SQL");


            string TipoQbe = "0";
            string QbeFilter = "12221";
            string BanderaSum = "0";
            string BanderaPreli = "0";
            string banExcel = "0";
            string codProgram = programa;
            string BanderaLupa = "1";
            string motor = "ORA";

            string tipoEscala = "12221";
            string descripcion = "PRUEBASWD";
            string editDescripcion = "PRUEBAWD02";
            string identificacion = tipoEscala;

            string valorInicial = "200";
            string valorFinal = "300";
            string escala = "2";
            string descripcionDet = "PRUEBASWD";
            string editDescripcionDet = "PRUEBASWD02";


            Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
            Dictionary<string, Point> botonesNavegadorDet1 = new Dictionary<string, Point>();
            Dictionary<string, Point> botonesNavegadorDet2 = new Dictionary<string, Point>();
            List<string> crudVariables = new List<string>() { tipoEscala, descripcion, editDescripcion };
            List<string> crudVariablesDet1 = new List<string>() { valorInicial, valorFinal, escala, descripcionDet, editDescripcionDet };
            WindowsDriver<WindowsElement> rootSession = null;
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

            botonesNavegador = PruebaCRUD.findname(desktopSession, 27, 0);
            var ElementList = PruebaCRUD.NavClass(desktopSession);


            //AGREGAR REGISTRO
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            CrudKed3esca.KEd3escaCRUD(desktopSession, crudVariables, 0);
            Thread.Sleep(5000);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            newFunctions_4.ScreenshotNuevo("Registro Agregado", true, file);
            //////////////////////Validacion Agregar
            //////////////////////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, identificacion, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);



            ///////DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKed3esca.KEd3escaCRUD(desktopSession, crudVariables, 1);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            newFunctions_4.ScreenshotNuevo("Edición Descartada", true, file);
            //////////////////////Validacion Editar Descartar
            ////////////////////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, tipo, editCampo, c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKed3esca.KEd3escaCRUD(desktopSession, crudVariables, 1);
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
                CrudKed3esca.KEd3escaCRUD(desktopSession, crudVariables, 1);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
            }
            newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
            //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, identificacion, "E", c, ErrorValidacion);
            ////Validacion Editar Aceptar
            //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, editDescripcion, editCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);
            //Thread.Sleep(1000);


            botonesNavegadorDet1 = PruebaCRUD.findname(desktopSession, 27, 1);
            ElementList = PruebaCRUD.NavClass(desktopSession);


            //AGREGAR REGISTRO DETALLE
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Insertar"].X, botonesNavegadorDet1["Insertar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            CrudKed3esca.KEd3escaCRUDDet1(desktopSession, crudVariablesDet1, 0);
            Thread.Sleep(5000);
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            newFunctions_4.ScreenshotNuevo("Registro Agregado", true, file);
            //////////////////////Validacion Agregar
            //////////////////////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, identificacion, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);



            ///////DESCARTAR EDICION DETALLE
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKed3esca.KEd3escaCRUDDet1(desktopSession, crudVariablesDet1, 1);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Cancelar"].X, botonesNavegadorDet1["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            newFunctions_4.ScreenshotNuevo("Edición Descartada", true, file);
            //////////////////////Validacion Editar Descartar
            ////////////////////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, tipo, editCampo, c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION DETALLE
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Editar"].X, botonesNavegadorDet1["Editar"].Y);
            desktopSession.Mouse.Click(null);
            CrudKed3esca.KEd3escaCRUDDet1(desktopSession, crudVariablesDet1, 1);
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
                CrudKed3esca.KEd3escaCRUDDet1(desktopSession, crudVariablesDet1, 1);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Aplicar"].X, botonesNavegadorDet1["Aplicar"].Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
            }
            newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
            //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, identificacion, "E", c, ErrorValidacion);
            ////Validacion Editar Aceptar
            //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, identificacion, editDescripcion, editCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);
            //Thread.Sleep(1000);


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
            //listaResultado = Textopdf(pdf, codProgram, user);
            //listaResultado.ForEach(e => ErrorValidacion.Add(e));

            //REPORTE PRELIMINAR
            string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
            errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            //////Rectificar Pie Pagina PDF PRELIMINAR
            //listaResultado = Textopdf(pdf1, codProgram, user);
            //listaResultado.ForEach(e => ErrorValidacion.Add(e));

            //EXPORTAR EXCEL
            string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
            errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

            ////RECTIFICACION DE CAMPOS DE EXCEL
            //string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, identificacion, c);
            //string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
            //Celda = Celda_Excel(ruta, nom_empr);
            //Celda.ForEach(u => ErrorValidacion.Add(u));
            //Thread.Sleep(1000);
            //LUPA
            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }



            ////ELIMINAR REGISTRO DETALLE
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Borrar"].X, botonesNavegadorDet1["Borrar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
            PruebaCRUD.cerrarBorrar(desktopSession);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
            Thread.Sleep(1000);

            isErrorDisplayed = newFunctions_4.editErrorHandler(desktopSession, file);

            if (isErrorDisplayed)
            {
                ////ELIMINAR REGISTRO DETALLE
                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDet1["Borrar"].X, botonesNavegadorDet1["Borrar"].Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
                PruebaCRUD.cerrarBorrar(desktopSession);
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
                Thread.Sleep(1000);
            }


            ////ELIMINAR REGISTRO
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
            PruebaCRUD.cerrarBorrar(desktopSession);
            Thread.Sleep(1000);
            rootSession = PruebaCRUD.RootSession();
            Thread.Sleep(5000);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(1000);
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
                Thread.Sleep(1000);
                rootSession = PruebaCRUD.RootSession();
                Thread.Sleep(5000);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
                Thread.Sleep(1000);
            }


        }


        [TestMethod]
        public void TestMethod1()
        {
            //
            // TODO: Agregar aquí la lógica de las pruebas
            //
        }
    }
}

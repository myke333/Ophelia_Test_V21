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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSl;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;

namespace Ophelia_Test_V21.TestMethods.ModulosSL
{
    /// <summary>
    /// Descripción resumida de ModulosSL_2
    /// </summary>
    [TestClass]
    public class ModulosSL_2 : FuncionesVitales
    {

        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;

        public ModulosSL_2() { }

        [TestMethod]
        public void KSlKitinCheckList()
        {
            List<string> listaResultado;
            List<string> Celda = new List<string>();
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
            //TableOrder = "ktes1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSL.ModulosSL_2.KSlKitinCheckList")
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
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&

                                rows["Cargo"].ToString().Length != 0 && rows["Cargo"].ToString() != null &&
                                rows["CentroCosto"].ToString().Length != 0 && rows["CentroCosto"].ToString() != null &&
                                rows["EditarCargo"].ToString().Length != 0 && rows["EditarCargo"].ToString() != null &&
                                rows["CodigoDetalle1"].ToString().Length != 0 && rows["CodigoDetalle1"].ToString() != null &&
                                rows["CodigoDetalle2"].ToString().Length != 0 && rows["CodigoDetalle2"].ToString() != null &&
                                rows["CodigoDetalle3"].ToString().Length != 0 && rows["CodigoDetalle3"].ToString() != null &&
                                rows["CodigoDetalle4"].ToString().Length != 0 && rows["CodigoDetalle4"].ToString() != null &&

                                //Validacion
                                rows["tablaDetalle1"].ToString().Length != 0 && rows["tablaDetalle1"].ToString() != null &&
                                rows["campoDetalle1"].ToString().Length != 0 && rows["campoDetalle1"].ToString() != null && 
                                //rows["editCampoD1"].ToString().Length != 0 && rows["editCampoD1"].ToString() != null &&
                                rows["tablaDetalle2"].ToString().Length != 0 && rows["tablaDetalle2"].ToString() != null &&
                                rows["campoDetalle2"].ToString().Length != 0 && rows["campoDetalle2"].ToString() != null &&
                                //rows["editCampoD2"].ToString().Length != 0 && rows["editCampoD2"].ToString() != null &&
                                rows["tablaDetalle3"].ToString().Length != 0 && rows["tablaDetalle3"].ToString() != null &&
                                rows["campoDetalle3"].ToString().Length != 0 && rows["campoDetalle3"].ToString() != null &&
                                //rows["editCampoD3"].ToString().Length != 0 && rows["editCampoD3"].ToString() != null &&
                                rows["tablaDetalle4"].ToString().Length != 0 && rows["tablaDetalle4"].ToString() != null &&
                                rows["campoDetalle4"].ToString().Length != 0 && rows["campoDetalle4"].ToString() != null 
                                //rows["editCampoD4"].ToString().Length != 0 && rows["editCampoD4"].ToString() != null

                                )
                            {
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
                                string EditCampo = rows["EditCampo"].ToString();

                                string Cargo = rows["Cargo"].ToString();
                                string CentroCosto = rows["CentroCosto"].ToString();
                                string EditarCargo = rows["EditarCargo"].ToString();
                                string CodigoDetalle1 = rows["CodigoDetalle1"].ToString();
                                string CodigoDetalle2 = rows["CodigoDetalle2"].ToString();
                                string CodigoDetalle3 = rows["CodigoDetalle3"].ToString();
                                string CodigoDetalle4 = rows["CodigoDetalle4"].ToString();

                                string tablaDetalle1 = rows["tablaDetalle1"].ToString();
                                string campoDetalle1 = rows["campoDetalle1"].ToString();
                                //string editCampoD1 = rows["editCampoD1"].ToString();
                                string tablaDetalle2 = rows["tablaDetalle2"].ToString();
                                string campoDetalle2 = rows["campoDetalle2"].ToString();
                                //string editCampoD2 = rows["editCampoD2"].ToString();
                                string tablaDetalle3 = rows["tablaDetalle3"].ToString();
                                string campoDetalle3 = rows["campoDetalle3"].ToString();
                                //string editCampoD3 = rows["editCampoD3"].ToString();
                                string tablaDetalle4 = rows["tablaDetalle4"].ToString();
                                string campoDetalle4 = rows["campoDetalle4"].ToString();
                                //string editCampoD4 = rows["editCampoD4"].ToString();

                                List<string> crudVariables = new List<string>() { Cargo, CentroCosto, EditarCargo };
                                List<string> crudDetalle1 = new List<string>() { CodigoDetalle1};
                                List<string> crudDetalle2 = new List<string>() { CodigoDetalle2};
                                List<string> crudDetalle3 = new List<string>() { CodigoDetalle3};
                                List<string> crudDetalle4 = new List<string>() { CodigoDetalle4};

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "SL", motor);
                                try
                                {

                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();

                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
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
                                        ////VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, CodigoDetalle1, CodigoDetalle1, campoDetalle1, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, CodigoDetalle2, CodigoDetalle2, campoDetalle2, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 3
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, CodigoDetalle3, CodigoDetalle3, campoDetalle3, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 4
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle4, user, motor, campoDetalle4, CodigoDetalle4, CodigoDetalle4, campoDetalle4, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CentroCosto, CentroCosto, Campo, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        ////AGREGAR REGISTRO
                                        ///AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKslkitin.CRUD(desktopSession, 1, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);
                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CentroCosto, CentroCosto, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);
                                       
                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKslkitin.CRUD(desktopSession, 2, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);
                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CentroCosto, Cargo, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKslkitin.CRUD(desktopSession, 2, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Aplicar Datos Editados", true, file);

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

                                            CrudKslkitin.CRUD(desktopSession, 2, crudVariables, file);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);
                                        
                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, CentroCosto, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CentroCosto, EditarCargo, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);

                                        ////AGREGAR DETALLE 1
                                        // 1 Opción detalle -> Recursos para el cargo 
                                        CrudKslkitin.ClickButton(desktopSession, 1);
                                        // 1 Opción detalle interno -> Activos Fijos 
                                        CrudKslkitin.ClickButtonInterno(desktopSession, 1);

                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Insertar Detalle 1", true, file);

                                        CrudKslkitin.CrudDetalle1(desktopSession, 1, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Detalle 1 Insertado", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Detalle 1 Aplicado", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, CodigoDetalle1, CodigoDetalle1, campoDetalle1, c, ErrorValidacion, "No se agregó correctamente 1", 0);

                                        ////AGREGAR DETALLE 2
                                        // 2 -> Ropa y Calzado de la Labor 
                                        CrudKslkitin.ClickButtonInterno(desktopSession, 2);

                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Insertar Detalle 2", true, file);

                                        CrudKslkitin.CrudDetalle2(desktopSession, 1, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Detalle 2 Insertado", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Detalle 2 Aplicado", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, CodigoDetalle2, CodigoDetalle2, campoDetalle2, c, ErrorValidacion, "No se agregó correctamente 2", 0);

                                        ////AGREGAR DETALLE 3
                                        //  3 -> Identificacion
                                        CrudKslkitin.ClickButtonInterno(desktopSession, 3);

                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Insertar Detalle 3", true, file);

                                        CrudKslkitin.CrudDetalle3(desktopSession, 1, crudDetalle3, file);
                                        newFunctions_4.ScreenshotNuevo("Detalle 3 Insertado", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Detalle 3 Aplicado", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, CodigoDetalle3, CodigoDetalle3, campoDetalle3, c, ErrorValidacion, "No se agregó correctamente 3", 0);

                                        ////AGREGAR DETALLE 4
                                        //  2 -> Ropa y Calzado de la Labor 
                                        CrudKslkitin.ClickButton(desktopSession, 2);

                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Insertar Detalle 4", true, file);

                                        CrudKslkitin.CrudDetalle4(desktopSession, 1, crudDetalle4, file);
                                        newFunctions_4.ScreenshotNuevo("Detalle 4 Insertado", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Detalle 4 Aplicado", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle4, user, motor, campoDetalle4, CodigoDetalle4, CodigoDetalle4, campoDetalle4, c, ErrorValidacion, "No se agregó correctamente 3", 0);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_1.lapiz(desktopSession);

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
                                        //CrudKnmctpre.Preliminar(desktopSession, file, pdf1, BanderaPreli);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //RECTIFICACION DE CAMPOS DE EXCEL  
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, CentroCosto, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        Thread.Sleep(1000);

                                        //ELIMINAR DETALLE 4
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador, file);

                                        ////Validacion Eliminar Detalle 4
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle4, motor, user, campoDetalle4, CodigoDetalle4, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle4, user, motor, campoDetalle4, CodigoDetalle4, "", campoDetalle4, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //ELIMINAR DETALLE 3
                                        CrudKslkitin.ClickButton(desktopSession, 1);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador, file);

                                        ////Validacion Eliminar Detalle 3
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle3, motor, user, campoDetalle3, CodigoDetalle3, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, CodigoDetalle3, "", campoDetalle3, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //ELIMINAR DETALLE 2
                                        CrudKslkitin.ClickButtonInterno(desktopSession, 2);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador, file);


                                        ////Validacion Eliminar Detalle 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, CodigoDetalle2, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, CodigoDetalle2, "", campoDetalle2, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //ELIMINAR DETALLE 1
                                        CrudKslkitin.ClickButtonInterno(desktopSession, 1);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador, file);

                                        ////Validacion Eliminar Detalle 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, CodigoDetalle1, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, CodigoDetalle1, "", campoDetalle1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);

                                        ////Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, CentroCosto, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CentroCosto, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);
                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }
                                    //CONTADOR DE ERRORES
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
        public void KSlArxccCheckList()
        {
            List<string> listaResultado;
            List<string> Celda = new List<string>();
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
            //TableOrder = "ktes1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSL.ModulosSL_2.KSlArxccCheckList")
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
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&

                                rows["CodigoCargo"].ToString().Length != 0 && rows["CodigoCargo"].ToString() != null &&
                                rows["CodigoCC"].ToString().Length != 0 && rows["CodigoCC"].ToString() != null &&
                                rows["CodigoJefe"].ToString().Length != 0 && rows["CodigoJefe"].ToString() != null &&
                                rows["EditarCodCargo"].ToString().Length != 0 && rows["EditarCodCargo"].ToString() != null &&

                                 rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&
                                rows["CodigoDetalle1"].ToString().Length != 0 && rows["CodigoDetalle1"].ToString() != null &&
                                rows["CodigoDetalle2"].ToString().Length != 0 && rows["CodigoDetalle2"].ToString() != null &&
                                rows["CodigoDetalle3"].ToString().Length != 0 && rows["CodigoDetalle3"].ToString() != null &&
                                rows["tablaDetalle1"].ToString().Length != 0 && rows["tablaDetalle1"].ToString() != null &&
                                rows["campoDetalle1"].ToString().Length != 0 && rows["campoDetalle1"].ToString() != null &&
                                rows["tablaDetalle2"].ToString().Length != 0 && rows["tablaDetalle2"].ToString() != null &&
                                rows["campoDetalle2"].ToString().Length != 0 && rows["campoDetalle2"].ToString() != null &&
                                rows["tablaDetalle3"].ToString().Length != 0 && rows["tablaDetalle3"].ToString() != null &&
                                rows["campoDetalle3"].ToString().Length != 0 && rows["campoDetalle3"].ToString() != null 
                                )
                            {
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

                                string CodigoCargo = rows["CodigoCargo"].ToString();
                                string CodigoCC = rows["CodigoCC"].ToString();
                                string CodigoJefe = rows["CodigoJefe"].ToString();
                                string EditarCodCargo = rows["EditarCodCargo"].ToString();
                                
                                string EditCampo = rows["EditCampo"].ToString();
                                string CodigoDetalle1 = rows["CodigoDetalle1"].ToString();
                                string CodigoDetalle2 = rows["CodigoDetalle2"].ToString();
                                string CodigoDetalle3 = rows["CodigoDetalle3"].ToString();
                                string tablaDetalle1 = rows["tablaDetalle1"].ToString();
                                string campoDetalle1 = rows["campoDetalle1"].ToString();
                                string tablaDetalle2 = rows["tablaDetalle2"].ToString();
                                string campoDetalle2 = rows["campoDetalle2"].ToString();
                                string tablaDetalle3 = rows["tablaDetalle3"].ToString();
                                string campoDetalle3 = rows["campoDetalle3"].ToString();

                                List<string> crudVariables = new List<string>() { CodigoCargo, CodigoCC, CodigoJefe, EditarCodCargo };
                        
                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "SL", motor);
                                try
                                {

                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();

                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
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

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, CodigoDetalle1, CodigoDetalle1, CodigoDetalle1, CodigoDetalle1, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, CodigoDetalle2, CodigoDetalle2, CodigoDetalle2, CodigoDetalle2, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 3
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, CodigoDetalle3, CodigoDetalle3, CodigoDetalle3, CodigoDetalle3, c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CodigoCC, CodigoCC, Campo, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        ////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKslarxcc.CRUD(desktopSession, 1, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);
                                        Thread.Sleep(1500);

                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CodigoCC, CodigoCC, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);
                                        
                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKslarxcc.CRUD(desktopSession, 2, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);
                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CodigoCC, CodigoCargo, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKslarxcc.CRUD(desktopSession, 2, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Aplicar Datos Editados", true, file);

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

                                            CrudKslarxcc.CRUD(desktopSession, 2, crudVariables, file);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);
                                        
                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, CodigoCC, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CodigoCC, EditarCodCargo, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);

                                        // 1 - CLick en recursos para el cargo, 2 - Identificación 3 - Información de Permisos y Accesos 4 - Activos Fijos
                                        CrudKslarxcc.ClickButton(desktopSession, 1);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 1
                                        Dictionary<string, Point> botonesNavegadorDetalle = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle = PruebaCRUD.findname1(desktopSession, 27, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Insertar"].X, botonesNavegadorDetalle["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 1", true, file);

                                        //Opción 1 -> detalle 1,  2 -> detalle 2, 3 -> detalle 3
                                        CrudKslarxcc.CrudDetalle(desktopSession, 1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Aplicar Datos Insertados Detalle 1", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, CodigoDetalle1, CodigoDetalle1, campoDetalle1, c, ErrorValidacion, "No se agregó correctamente 1", 0);

                                        // 1 - CLick en recursos para el cargo, 2 - Identificación 3 - Información de Permisos y Accesos 4 - Activos Fijos
                                        CrudKslarxcc.ClickButton(desktopSession, 2);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 2
                                        botonesNavegadorDetalle = PruebaCRUD.findname1(desktopSession, 27, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Insertar"].X, botonesNavegadorDetalle["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 2", true, file);

                                        //Opción 1 -> detalle 1,  2 -> detalle 2, 3 -> detalle 3
                                        CrudKslarxcc.CrudDetalle(desktopSession, 2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Aplicar Datos Insertados Detalle 2", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, CodigoDetalle2, CodigoDetalle2, campoDetalle2, c, ErrorValidacion, "No se agregó correctamente 2", 0);

                                        // 1 - CLick en recursos para el cargo, 2 - Identificación 3 - Información de Permisos y Accesos 4 - Activos Fijos
                                        CrudKslarxcc.ClickButton(desktopSession, 3);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 3
                                        botonesNavegadorDetalle = PruebaCRUD.findname1(desktopSession, 27, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Insertar"].X, botonesNavegadorDetalle["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 3", true, file);

                                        //Opción 1 -> detalle 1,  2 -> detalle 2, 3 -> detalle 3
                                        CrudKslarxcc.CrudDetalle(desktopSession, 3, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Aplicar Datos Insertados Detalle 3", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, CodigoDetalle3, CodigoDetalle3, campoDetalle3, c, ErrorValidacion, "No se agregó correctamente 1", 0);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_1.lapiz(desktopSession);

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
                                        //CrudKnmctpre.Preliminar(desktopSession, file, pdf1, BanderaPreli);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //RECTIFICACION DE CAMPOS DE EXCEL  
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, CodigoCC, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        Thread.Sleep(1000);
                                        //ELIMINANDO DETALLE 3
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorDetalle, file);

                                        ////Validacion Eliminar Detalle 3
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle3, motor, user, campoDetalle3, CodigoDetalle3, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, CodigoDetalle3, "", campoDetalle3, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        // 1 - CLick en recursos para el cargo, 2 - Identificación 3 - Información de Permisos y Accesos 4 - Activos Fijos
                                        CrudKslarxcc.ClickButton(desktopSession, 1);

                                        //ELIMINANDO DETALLE 2
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorDetalle, file);

                                        ////Validacion Eliminar Detalle 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, CodigoDetalle2, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, CodigoDetalle2, "", campoDetalle2, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        // 1 - CLick en recursos para el cargo, 2 - Identificación 3 - Información de Permisos y Accesos 4 - Activos Fijos
                                        CrudKslarxcc.ClickButton(desktopSession, 4);

                                        //ELIMINANDO DETALLE 1
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorDetalle, file);

                                        ////Validacion Eliminar Detalle 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, CodigoDetalle1, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, CodigoDetalle1, "", campoDetalle1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);

                                        ////Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, CodigoCC, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CodigoCC, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);
                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }
                                    //CONTADOR DE ERRORES
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
        public void KSlParslCheckList()
        {
            List<string> listaResultado;
            List<string> Celda = new List<string>();
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
            //TableOrder = "ktes1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSL.ModulosSL_2.KSlParslCheckList")
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
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                //rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                 rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&
                                //VALIDACION
                                rows["tablaDetalle1"].ToString().Length != 0 && rows["tablaDetalle1"].ToString() != null &&
                                rows["campoDetalle1"].ToString().Length != 0 && rows["campoDetalle1"].ToString() != null &&
                                rows["editCampoD1"].ToString().Length != 0 && rows["editCampoD1"].ToString() != null &&
                                rows["tablaDetalle2"].ToString().Length != 0 && rows["tablaDetalle2"].ToString() != null &&
                                rows["campoDetalle2"].ToString().Length != 0 && rows["campoDetalle2"].ToString() != null &&
                                rows["editCampoD2"].ToString().Length != 0 && rows["editCampoD2"].ToString() != null &&

                                //VARIABLES SQL
                                rows["CambioInscripcion"].ToString().Length != 0 && rows["CambioInscripcion"].ToString() != null &&
                                rows["FiltrarVacanteCosto"].ToString().Length != 0 && rows["FiltrarVacanteCosto"].ToString() != null &&
                                rows["EditarFiltrarVacanteCosto"].ToString().Length != 0 && rows["EditarFiltrarVacanteCosto"].ToString() != null &&

                                //VARIABLES ORA
                                rows["CargoAusentismo"].ToString().Length != 0 && rows["CargoAusentismo"].ToString() != null &&
                                rows["RequisicionPersonal"].ToString().Length != 0 && rows["RequisicionPersonal"].ToString() != null &&
                                rows["ValidacionPerfiles"].ToString().Length != 0 && rows["ValidacionPerfiles"].ToString() != null &&
                                rows["ProcesosActivos"].ToString().Length != 0 && rows["ProcesosActivos"].ToString() != null &&
                                rows["RevisionHoja"].ToString().Length != 0 && rows["RevisionHoja"].ToString() != null &&
                                rows["EtapasSeleccion"].ToString().Length != 0 && rows["EtapasSeleccion"].ToString() != null &&
                                rows["KitIngreso"].ToString().Length != 0 && rows["KitIngreso"].ToString() != null &&
                                rows["CoberturaVacante"].ToString().Length != 0 && rows["CoberturaVacante"].ToString() != null &&
                                rows["ActivarSueldo"].ToString().Length != 0 && rows["ActivarSueldo"].ToString() != null &&

                                rows["ActivarTipoContrato"].ToString().Length != 0 && rows["ActivarTipoContrato"].ToString() != null &&
                                rows["Observaciones"].ToString().Length != 0 && rows["Observaciones"].ToString() != null &&
                                rows["ValidarCargo"].ToString().Length != 0 && rows["ValidarCargo"].ToString() != null &&

                                rows["ManejaAnexos"].ToString().Length != 0 && rows["ManejaAnexos"].ToString() != null &&
                                rows["ValidaDocumentos"].ToString().Length != 0 && rows["ValidaDocumentos"].ToString() != null &&
                                rows["SedesVacantes"].ToString().Length != 0 && rows["SedesVacantes"].ToString() != null &&
                                rows["ConvocatoriaProceso"].ToString().Length != 0 && rows["ConvocatoriaProceso"].ToString() != null &&
                                rows["VigenciaElegidos"].ToString().Length != 0 && rows["VigenciaElegidos"].ToString() != null &&
                                rows["CerrarPrueba"].ToString().Length != 0 && rows["CerrarPrueba"].ToString() != null &&

                                rows["CausalesRUIC"].ToString().Length != 0 && rows["CausalesRUIC"].ToString() != null &&
                                rows["ActivarConvocatoria"].ToString().Length != 0 && rows["ActivarConvocatoria"].ToString() != null &&
                                rows["FiltrarVacantes"].ToString().Length != 0 && rows["FiltrarVacantes"].ToString() != null &&
                                rows["EditarFiltrarVacantes"].ToString().Length != 0 && rows["EditarFiltrarVacantes"].ToString() != null &&

                                rows["VariableLupa"].ToString().Length != 0 && rows["VariableLupa"].ToString() != null &&

                                //DETALLE 1
                                rows["EstadoCargo"].ToString().Length != 0 && rows["EstadoCargo"].ToString() != null &&
                                rows["ApliHomologacion"].ToString().Length != 0 && rows["ApliHomologacion"].ToString() != null &&
                                rows["CodHomologacion"].ToString().Length != 0 && rows["CodHomologacion"].ToString() != null &&
                                rows["EditCodHomologacion"].ToString().Length != 0 && rows["EditCodHomologacion"].ToString() != null &&

                                //DETALLE2
                                rows["Validar"].ToString().Length != 0 && rows["Validar"].ToString() != null &&
                                rows["Factor"].ToString().Length != 0 && rows["Factor"].ToString() != null &&
                                rows["PesoFactor"].ToString().Length != 0 && rows["PesoFactor"].ToString() != null &&
                                rows["EditPesoFactor"].ToString().Length != 0 && rows["EditPesoFactor"].ToString() != null
                                )
                            {
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
                                string EditCampo = rows["EditCampo"].ToString();
                                //VALIDACION
                                string tablaDetalle1 = rows["tablaDetalle1"].ToString();
                                string campoDetalle1 = rows["campoDetalle1"].ToString();
                                string editCampoD1 = rows["editCampoD1"].ToString();
                                string tablaDetalle2 = rows["tablaDetalle2"].ToString();
                                string campoDetalle2 = rows["campoDetalle2"].ToString();
                                string editCampoD2 = rows["editCampoD2"].ToString();

                                //VARIABLES SQL
                                string CambioInscripcion = rows["CambioInscripcion"].ToString();
                                string FiltrarVacanteCosto = rows["FiltrarVacanteCosto"].ToString();
                                string EditarFiltrarVacanteCosto = rows["EditarFiltrarVacanteCosto"].ToString();

                                //VARIABLES ORA
                                string CargoAusentismo = rows["CargoAusentismo"].ToString();
                                string RequisicionPersonal = rows["RequisicionPersonal"].ToString();
                                string ValidacionPerfiles = rows["ValidacionPerfiles"].ToString();
                                string ProcesosActivos = rows["ProcesosActivos"].ToString();
                                string RevisionHoja = rows["RevisionHoja"].ToString();
                                string EtapasSeleccion = rows["EtapasSeleccion"].ToString();
                                string KitIngreso = rows["KitIngreso"].ToString();
                                string CoberturaVacante = rows["CoberturaVacante"].ToString();
                                string ActivarSueldo = rows["ActivarSueldo"].ToString();
                                string ActivarTipoContrato = rows["ActivarTipoContrato"].ToString();
                                string Observaciones = rows["Observaciones"].ToString();
                                string ValidarCargo = rows["ValidarCargo"].ToString();
                                string ManejaAnexos = rows["ManejaAnexos"].ToString();
                                string ValidaDocumentos = rows["ValidaDocumentos"].ToString();
                                string SedesVacantes = rows["SedesVacantes"].ToString();
                                string ConvocatoriaProceso = rows["ConvocatoriaProceso"].ToString();
                                string VigenciaElegidos = rows["VigenciaElegidos"].ToString();
                                string CerrarPrueba = rows["CerrarPrueba"].ToString();
                                string CausalesRUIC = rows["CausalesRUIC"].ToString();
                                string ActivarConvocatoria = rows["ActivarConvocatoria"].ToString();
                                string FiltrarVacantes = rows["FiltrarVacantes"].ToString();
                                string EditarFiltrarVacantes = rows["EditarFiltrarVacantes"].ToString();

                                string VariableLupa = rows["VariableLupa"].ToString();

                                string EstadoCargo = rows["EstadoCargo"].ToString();
                                string ApliHomologacion = rows["ApliHomologacion"].ToString();
                                string CodHomologacion = rows["CodHomologacion"].ToString();
                                string EditCodHomologacion = rows["EditCodHomologacion"].ToString();

                                string Validar = rows["Validar"].ToString();
                                string Factor = rows["Factor"].ToString();
                                string PesoFactor = rows["PesoFactor"].ToString();
                                string EditPesoFactor = rows["EditPesoFactor"].ToString();


                                List<string> crudVariables = new List<string>() { CambioInscripcion, FiltrarVacanteCosto, EditarFiltrarVacanteCosto};
                                List<string> crudVariablesORA = new List<string>() { CargoAusentismo, RequisicionPersonal, ValidacionPerfiles, ProcesosActivos, RevisionHoja, EtapasSeleccion, KitIngreso, CoberturaVacante, ActivarSueldo, ActivarTipoContrato, Observaciones, ValidarCargo, ManejaAnexos, ValidaDocumentos, SedesVacantes, ConvocatoriaProceso, VigenciaElegidos, CerrarPrueba, CausalesRUIC, ActivarConvocatoria, FiltrarVacantes, EditarFiltrarVacantes };
                                List<string> crudVariablesLupaORA = new List<string>() { VariableLupa };
                                List<string> crudDetalle1ORA = new List<string>() { EstadoCargo, ApliHomologacion, CodHomologacion, EditCodHomologacion };
                                List<string> crudDetalle2ORA = new List<string>() { Validar, Factor, PesoFactor, EditPesoFactor};

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "SL", motor);
                                try
                                {

                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();

                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
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

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_1.lapiz(desktopSession);

                                        ///AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        //validacion error al editar
                                        bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);

                                        if (motor == "SQL")
                                        {

                                            //ELIMINAR REGISTRO
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);

                                            ////Validacion Eliminar Registro
                                            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, CambioInscripcion, "D", c, ErrorValidacion);
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CambioInscripcion, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                            desktopSession.Mouse.Click(null);

                                            // 2 -> CLick en Parámetros Requisiciones 1 -> Parámetros Convocatorias
                                            CrudKslparsl.ClickButton(desktopSession, 1);
                                            
                                            //Thread.Sleep(5000);
                                            CrudKslparsl.CRUDSQL(desktopSession, 1, crudVariables);
                                            newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CambioInscripcion, CambioInscripcion, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);
                                            
                                            Thread.Sleep(1000);
                                            //DESCARTAR EDICION
                                            Thread.Sleep(1000);
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar", true, file);
                                            
                                            CrudKslparsl.CRUDSQL(desktopSession, 2, crudVariables);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);
                                            Thread.Sleep(1000);

                                            ////Validacion Editar Descartar
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CambioInscripcion, FiltrarVacanteCosto, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                            ////ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar", true, file);
                                            Thread.Sleep(1000);

                                            CrudKslparsl.CRUDSQL(desktopSession, 2, crudVariables);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Aplicar Datos Editados", true, file);
                                          
                                            ////Validacion Editar
                                            Thread.Sleep(1500);
                                            // //Debugger.Launch();
                                            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, CambioInscripcion, "E", c, ErrorValidacion);
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CambioInscripcion, EditarFiltrarVacanteCosto, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        }
                                        else
                                        {
                                            // 2 -> CLick en Parámetros Requisiciones 1 -> Parámetros Convocatorias
                                            CrudKslparsl.ClickButton(desktopSession, 1);

                                            Dictionary<string, Point> botonesNavegadorDetalle = new Dictionary<string, Point>();
                                            var botonesNavegadorDetalle1 = CrudKslparsl.findname(desktopSession, 26, 1);
                                            var ElementList1 = PruebaCRUD.NavClass(desktopSession);

                                            //ELIMINAR DETALLE 1
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[1], botonesNavegadorDetalle1, file);
                                            Thread.Sleep(500);

                                            ////Validacion Eliminar Detalle 1
                                            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, EstadoCargo, "D", c, ErrorValidacion);
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, EstadoCargo, "", campoDetalle1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                            // 2 -> CLick en Parámetros Requisiciones 1 -> Parámetros Convocatorias
                                            CrudKslparsl.ClickButton(desktopSession, 2);

                                            // 1 ->Generales 2 -> Parámetros de Requisición Web 3 -> Parámetros Reporte Validación de Perfil
                                            CrudKslparsl.ClickButtonInterno(desktopSession, 3);

                                            Dictionary<string, Point> botonesNavegadorDetalles = new Dictionary<string, Point>();
                                            var botonesNavegadorDetalle2 = PruebaCRUD.findname(desktopSession, 27, 1);
                                            var ElementList2 = PruebaCRUD.NavClass(desktopSession);
                                            
                                            //ELIMINAR DETALLE 2
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList2[1], botonesNavegadorDetalle2, file);
                                            Thread.Sleep(500);

                                            ////Validacion Eliminar Detalle 2
                                            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, Factor, "D", c, ErrorValidacion);
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, Factor, "", campoDetalle2, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                            // 1 ->Generales 2 -> Parámetros de Requisición Web 3 -> Parámetros Reporte Validación de Perfil
                                            CrudKslparsl.ClickButtonInterno(desktopSession, 1);
                                            
                                            //ELIMINAR REGISTRO
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);

                                            ////Validacion Eliminar Registro
                                            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, CambioInscripcion, "D", c, ErrorValidacion);
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CambioInscripcion, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                            desktopSession.Mouse.Click(null);

                                            //AGREGANDO AL CRUD crudVariablesORA
                                            CrudKslparsl.CRUDORA(desktopSession, 1, crudVariablesORA, crudVariablesLupaORA, file);
                                            newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CambioInscripcion, CambioInscripcion, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

                                            Thread.Sleep(1000);
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                            CrudKslparsl.CRUDORA(desktopSession, 2, crudVariablesORA, crudVariablesLupaORA, file);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);
                                            Thread.Sleep(1000);

                                            ////Validacion Editar Descartar
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CambioInscripcion, FiltrarVacantes, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar", true, file);
                                            Thread.Sleep(1000);

                                            CrudKslparsl.CRUDORA(desktopSession, 2, crudVariablesORA, crudVariablesLupaORA, file);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Aplicar Datos Editados", true, file);

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

                                                CrudKslparsl.CRUDORA(desktopSession, 2, crudVariablesORA, crudVariablesLupaORA, file);
                                                newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                                desktopSession.Mouse.Click(null);
                                                Thread.Sleep(3000);
                                            }
                                            newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);
                                            
                                            ////Validacion Editar
                                            Thread.Sleep(1500);
                                            // //Debugger.Launch();
                                            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, CambioInscripcion, "E", c, ErrorValidacion);
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CambioInscripcion, EditarFiltrarVacantes, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                            //AGREGAR CRUD DETALLE 1 PARAMETROS CONVOCATORIAS
                                            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Nuevo"].X, botonesNavegadorDetalle1["Nuevo"].Y);
                                            desktopSession.Mouse.Click(null);

                                            CrudKslparsl.CRUDDetalle1ORA(desktopSession, 1, crudDetalle1ORA);
                                            newFunctions_4.ScreenshotNuevo("Insertar Detalle 1", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Aplicar"].X, botonesNavegadorDetalle1["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Detalle 1 Insertado", true, file);

                                            Thread.Sleep(1500);
                                            ////Validacion Agregar
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, EstadoCargo, EstadoCargo, campoDetalle1, c, ErrorValidacion, "No se agregó correctamente 1", 0);

                                            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Editar"].X, botonesNavegadorDetalle1["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                            CrudKslparsl.CRUDDetalle1ORA(desktopSession, 2, crudDetalle1ORA);
                                            newFunctions_4.ScreenshotNuevo("Detalle 1 Editado", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Descartar"].X, botonesNavegadorDetalle1["Descartar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Cancelar Detalle 1 Editado", true, file);

                                            ////Validacion Editar Descartar
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, EstadoCargo, CodHomologacion, editCampoD1, c, ErrorValidacion, "No se edito correctamente", 1);
                                         
                                            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Editar"].X, botonesNavegadorDetalle1["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                            CrudKslparsl.CRUDDetalle1ORA(desktopSession, 2, crudDetalle1ORA);
                                            newFunctions_4.ScreenshotNuevo("Detalle 1 Editado", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Aplicar"].X, botonesNavegadorDetalle1["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Edición Detalle 1 Aplicado", true, file);

                                            ////Validacion Editar
                                            Thread.Sleep(3000);
                                            // //Debugger.Launch();
                                            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, EstadoCargo, "E", c, ErrorValidacion);
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, EstadoCargo, EditCodHomologacion, editCampoD1, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                            // 2 -> CLick en Parámetros Requisiciones 1 -> Parámetros Convocatorias
                                            CrudKslparsl.ClickButton(desktopSession, 2);

                                            // 1 ->Generales 2 -> Parámetros de Requisición Web 3 -> Parámetros Reporte Validación de Perfil
                                            CrudKslparsl.ClickButtonInterno(desktopSession, 3);

                                            //AGREGAR CRUD DETALLE 2 PARAMETROS REPORTE VALIDACION DE PERFIL
                                            desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Insertar"].X, botonesNavegadorDetalle2["Insertar"].Y);
                                            desktopSession.Mouse.Click(null);

                                            CrudKslparsl.CRUDDetalle2ORA(desktopSession, 1, crudDetalle2ORA);
                                            newFunctions_4.ScreenshotNuevo("Insertar Detalle 2", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Detalle 2 Insertado", true, file);
                                            Thread.Sleep(1000);

                                            Thread.Sleep(1500);
                                            ////Validacion Agregar
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, Factor, Factor, campoDetalle2, c, ErrorValidacion, "No se agregó correctamente 2", 0);

                                            desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);
                                            Thread.Sleep(1000);

                                            CrudKslparsl.CRUDDetalle2ORA(desktopSession, 2, crudDetalle2ORA);
                                            newFunctions_4.ScreenshotNuevo("Detalle 2 Editado", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Cancelar"].X, botonesNavegadorDetalle2["Cancelar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Cancelar Detalle 2 Editado", true, file);
                                            Thread.Sleep(1000);

                                            ////Validacion Editar Descartar
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, Factor, PesoFactor, editCampoD2, c, ErrorValidacion, "No se edito correctamente", 1);

                                            desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);
                                            Thread.Sleep(1000);

                                            CrudKslparsl.CRUDDetalle2ORA(desktopSession, 2, crudDetalle2ORA);
                                            newFunctions_4.ScreenshotNuevo("Detalle 2 Editado", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Edición Detalle 2 Aplicado", true, file);

                                            ////Validacion Editar
                                            Thread.Sleep(1500);
                                            // //Debugger.Launch();
                                            ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, Factor, "E", c, ErrorValidacion);
                                            ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, Factor, EditPesoFactor, editCampoD2, c, ErrorValidacion, "El registro no se editó correctamente", 1);
                                        }

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_1.lapiz(desktopSession);

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
                                        //CrudKnmctpre.Preliminar(desktopSession, file, pdf1, BanderaPreli);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //RECTIFICACION DE CAMPOS DE EXCEL  
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, CambioInscripcion, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Thread.Sleep(3000);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }
                                    //CONTADOR DE ERRORES
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
        public void KSlGrselCheckList()
        {
            List<string> listaResultado;
            List<string> Celda = new List<string>();
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
            //TableOrder = "ktes1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSL.ModulosSL_2.KSlGrselCheckList")
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
                                rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&

                                rows["GrupoSelecNum"].ToString().Length != 0 && rows["GrupoSelecNum"].ToString() != null &&
                                rows["GrupoSelecPal"].ToString().Length != 0 && rows["GrupoSelecPal"].ToString() != null &&
                                rows["Oportunidad"].ToString().Length != 0 && rows["Oportunidad"].ToString() != null &&
                                rows["Justificacion"].ToString().Length != 0 && rows["Justificacion"].ToString() != null &&
                                rows["EditOportunidad"].ToString().Length != 0 && rows["EditOportunidad"].ToString() != null
                                )
                            {
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
                                string EditCampo = rows["EditCampo"].ToString();

                                string GrupoSelecNum = rows["GrupoSelecNum"].ToString();
                                string GrupoSelecPal = rows["GrupoSelecPal"].ToString();
                                string Oportunidad = rows["Oportunidad"].ToString();
                                string Justificacion = rows["Justificacion"].ToString();
                                string EditOportunidad = rows["EditOportunidad"].ToString();


                                List<string> crudVariables = new List<string>() { GrupoSelecNum, GrupoSelecPal, Oportunidad, Justificacion, EditOportunidad };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "SL", motor);
                                try
                                {

                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();

                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
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
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, GrupoSelecNum, GrupoSelecNum, Campo, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        ///AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKslgrsel.CRUD(desktopSession, 1, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);
                                        Thread.Sleep(1500);
                                        
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, GrupoSelecNum, GrupoSelecNum, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);
                                        
                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKslgrsel.CRUD(desktopSession, 2, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);
                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, GrupoSelecNum, Oportunidad, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKslgrsel.CRUD(desktopSession, 2, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Aplicar Datos Editados", true, file);

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

                                            CrudKslgrsel.CRUD(desktopSession, 2, crudVariables, file);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);
                                       
                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, GrupoSelecNum, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, GrupoSelecNum, EditOportunidad, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_1.lapiz(desktopSession);

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
                                        //CrudKnmctpre.Preliminar(desktopSession, file, pdf1, BanderaPreli);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //RECTIFICACION DE CAMPOS DE EXCEL  
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, GrupoSelecNum, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        Thread.Sleep(1000);
                                        //ELIMINAR DETALLE
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);

                                        ////Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, GrupoSelecNum, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, GrupoSelecNum, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);
                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }
                                    //CONTADOR DE ERRORES
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
 
    }
}

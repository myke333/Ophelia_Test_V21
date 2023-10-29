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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosFd;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;

namespace Ophelia_Test_V21.TestMethods.ModulosFD
{
    /// <summary>
    /// Descripción resumida de ModulosFD_7
    /// </summary>
    [TestClass]
    public class ModulosFD_7 : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;

        public ModulosFD_7() { }


        [TestMethod]
        public void KFdPlcurCheckList()
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
            TableOrder = "ktes1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosFD.ModulosFD_7.KFdPlcurCheckList")
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
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                //CRUD
                                rows["Parametro"].ToString().Length != 0 && rows["Parametro"].ToString() != null &&
                                rows["Curso"].ToString().Length != 0 && rows["Curso"].ToString() != null &&
                                rows["Programa"].ToString().Length != 0 && rows["Programa"].ToString() != null &&
                                rows["Plan"].ToString().Length != 0 && rows["Plan"].ToString() != null &&
                                rows["Entidad"].ToString().Length != 0 && rows["Entidad"].ToString() != null &&
                                rows["EntidadAuspiciadora"].ToString().Length != 0 && rows["EntidadAuspiciadora"].ToString() != null &&
                                rows["TipoCapacitacion"].ToString().Length != 0 && rows["TipoCapacitacion"].ToString() != null &&
                                rows["FechaInicio"].ToString().Length != 0 && rows["FechaInicio"].ToString() != null &&
                                rows["FechaFinal"].ToString().Length != 0 && rows["FechaFinal"].ToString() != null &&
                                rows["Modalidad"].ToString().Length != 0 && rows["Modalidad"].ToString() != null &&
                                rows["Lugar"].ToString().Length != 0 && rows["Lugar"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["EditarObservacion"].ToString().Length != 0 && rows["EditarObservacion"].ToString() != null &&
                                //DETALLE 1
                                rows["SecuencialDefinicion"].ToString().Length != 0 && rows["SecuencialDefinicion"].ToString() != null &&
                                rows["TipEvaluacion"].ToString().Length != 0 && rows["TipEvaluacion"].ToString() != null &&
                                rows["Origen"].ToString().Length != 0 && rows["Origen"].ToString() != null &&
                                rows["Clasificacion"].ToString().Length != 0 && rows["Clasificacion"].ToString() != null &&
                                rows["Calificacion"].ToString().Length != 0 && rows["Calificacion"].ToString() != null &&
                                rows["FechaIni"].ToString().Length != 0 && rows["FechaIni"].ToString() != null &&
                                rows["FechaFin"].ToString().Length != 0 && rows["FechaFin"].ToString() != null &&
                                rows["HoraInicio"].ToString().Length != 0 && rows["HoraInicio"].ToString() != null &&
                                rows["HoraFinal"].ToString().Length != 0 && rows["HoraFinal"].ToString() != null && 
                                rows["TiempoMax"].ToString().Length != 0 && rows["TiempoMax"].ToString() != null &&
                                rows["EditarTiempoMax"].ToString().Length != 0 && rows["EditarTiempoMax"].ToString() != null &&
                                //DETALLE 2
                                rows["Item"].ToString().Length != 0 && rows["Item"].ToString() != null &&
                                //DETALLE 4
                                rows["Sesion"].ToString().Length != 0 && rows["Sesion"].ToString() != null &&
                                rows["Responsable"].ToString().Length != 0 && rows["Responsable"].ToString() != null &&
                                rows["Lupa"].ToString().Length != 0 && rows["Lupa"].ToString() != null &&
                                rows["LugarD"].ToString().Length != 0 && rows["LugarD"].ToString() != null &&
                                rows["EditarSesion"].ToString().Length != 0 && rows["EditarSesion"].ToString() != null &&
                                //DETALLE 5
                                rows["Proveedor"].ToString().Length != 0 && rows["Proveedor"].ToString() != null &&
                                rows["CodLogistica"].ToString().Length != 0 && rows["CodLogistica"].ToString() != null &&
                                rows["AfectaPresupuesto"].ToString().Length != 0 && rows["AfectaPresupuesto"].ToString() != null &&
                                rows["AfectaTotal"].ToString().Length != 0 && rows["AfectaTotal"].ToString() != null &&
                                rows["EditAfectaTotal"].ToString().Length != 0 && rows["EditAfectaTotal"].ToString() != null &&
                                //DETALLE 6
                                rows["FechaSolicitud"].ToString().Length != 0 && rows["FechaSolicitud"].ToString() != null &&
                                rows["EditarFechaSol"].ToString().Length != 0 && rows["EditarFechaSol"].ToString() != null &&
                                //DETALLE 7
                                rows["NombreDocumento"].ToString().Length != 0 && rows["NombreDocumento"].ToString() != null &&
                                rows["PorcentajeDoc"].ToString().Length != 0 && rows["PorcentajeDoc"].ToString() != null &&
                                //DETALLE 8
                                rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
                                rows["ResponsableDD"].ToString().Length != 0 && rows["ResponsableDD"].ToString() != null &&
                                rows["FechaCumplimiento"].ToString().Length != 0 && rows["FechaCumplimiento"].ToString() != null &&
                                rows["EditFechaCumplimiento"].ToString().Length != 0 && rows["EditFechaCumplimiento"].ToString() != null 
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

                                string Parametro = rows["Parametro"].ToString();
                                string Curso = rows["Curso"].ToString();
                                string Programa = rows["Programa"].ToString();
                                string Plan = rows["Plan"].ToString();
                                string Entidad = rows["Entidad"].ToString();
                                string EntidadAuspiciadora = rows["EntidadAuspiciadora"].ToString();
                                string TipoCapacitacion = rows["TipoCapacitacion"].ToString();
                                string FechaInicio = rows["FechaInicio"].ToString();
                                string FechaFinal = rows["FechaFinal"].ToString();
                                string Modalidad = rows["Modalidad"].ToString();
                                string Lugar = rows["Lugar"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string EditarObservacion = rows["EditarObservacion"].ToString();


                                string SecuencialDefinicion = rows["SecuencialDefinicion"].ToString();
                                string TipEvaluacion = rows["TipEvaluacion"].ToString();
                                string Origen = rows["Origen"].ToString();
                                string Clasificacion = rows["Clasificacion"].ToString();
                                string Calificacion = rows["Calificacion"].ToString();
                                string FechaIni = rows["FechaIni"].ToString();
                                string FechaFin = rows["FechaFin"].ToString();
                                string HoraInicio = rows["HoraInicio"].ToString();
                                string HoraFinal = rows["HoraFinal"].ToString(); 
                                string TiempoMax = rows["TiempoMax"].ToString();
                                string EditarTiempoMax = rows["EditarTiempoMax"].ToString();

                                string Item = rows["Item"].ToString();

                                string Sesion = rows["Sesion"].ToString();
                                string Responsable = rows["Responsable"].ToString();
                                string Lupa = rows["Lupa"].ToString(); 
                                string LugarD = rows["LugarD"].ToString();
                                string EditarSesion = rows["EditarSesion"].ToString();

                                string Proveedor = rows["Proveedor"].ToString();
                                string CodLogistica = rows["CodLogistica"].ToString();
                                string AfectaPresupuesto = rows["AfectaPresupuesto"].ToString(); 
                                string AfectaTotal = rows["AfectaTotal"].ToString();
                                string EditAfectaTotal = rows["EditAfectaTotal"].ToString();

                                string FechaSolicitud = rows["FechaSolicitud"].ToString();
                                string EditarFechaSol = rows["EditarFechaSol"].ToString();

                                string NombreDocumento = rows["NombreDocumento"].ToString();
                                string PorcentajeDoc = rows["PorcentajeDoc"].ToString();

                                string Fecha = rows["Fecha"].ToString();
                                string ResponsableDD = rows["ResponsableDD"].ToString();
                                string FechaCumplimiento = rows["FechaCumplimiento"].ToString();
                                string EditFechaCumplimiento = rows["EditFechaCumplimiento"].ToString();


                                List<string> crudVariables = new List<string>() { Parametro, Curso, Programa, Plan, Entidad, EntidadAuspiciadora, TipoCapacitacion, FechaInicio, FechaFinal, Modalidad, Lugar, Observacion, EditarObservacion };
                                List<string> crudDetalle1 = new List<string>() { SecuencialDefinicion, TipEvaluacion, Origen, Clasificacion, Calificacion, FechaIni, FechaFin, TiempoMax, EditarTiempoMax };
                                List<string> crudDetalle2 = new List<string>() { Item };
                                List<string> crudDetalle4 = new List<string>() { Sesion, Responsable, Lupa, LugarD, EditarSesion };
                                List<string> crudDetalle5 = new List<string>() { Proveedor, CodLogistica, AfectaPresupuesto, AfectaTotal, EditAfectaTotal };
                                List<string> crudDetalle6 = new List<string>() { FechaSolicitud, EditarFechaSol };
                                List<string> crudDetalle7 = new List<string>() { NombreDocumento, PorcentajeDoc };
                                List<string> crudDetalle8 = new List<string>() { Fecha, ResponsableDD, FechaCumplimiento, EditFechaCumplimiento };


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "FD", motor);
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

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        ////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 27, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKfdplcur.CRUD(desktopSession, 1, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(2000);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Lugar, Lugar, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKfdplcur.CRUD(desktopSession, 2, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Lugar, Observacion, "OBS_ERVA", c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKfdplcur.CRUD(desktopSession, 2, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

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

                                            CrudKfdplcur.CRUD(desktopSession, 2, crudVariables, file);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Lugar, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Lugar, EditarObservacion, "OBS_ERVA", c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);

                                        //3->Evaluación curso
                                        CrudKfdplcur.ClickButtonExterno(desktopSession, 3);
                                        //Thread.Sleep(20000000);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 1
                                        Dictionary<string, Point> botonesNavegadorDetalle = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle = PruebaCRUD.findname(desktopSession, 27, 1);
                                        var ElementList1 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Insertar"].X, botonesNavegadorDetalle["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKfdplcur.CrudDetalle1(desktopSession, 1, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                        Thread.Sleep(1000);

                                        CrudKfdplcur.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Cancelar"].X, botonesNavegadorDetalle["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 1", true, file);

                                        Thread.Sleep(1000);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                        Thread.Sleep(1000);

                                        CrudKfdplcur.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 1", true, file);

                                        if (motor == "SQL")
                                        {
                                            Thread.Sleep(1000);
                                            ////AGREGAR DETALLE 2
                                            Dictionary<string, Point> botonesNavegadorDetalle2 = new Dictionary<string, Point>();
                                            botonesNavegadorDetalle2 = PruebaCRUD.findname(desktopSession, 27, 2);
                                            var ElementList2 = PruebaCRUD.NavClass(desktopSession);
                                            desktopSession.Mouse.MouseMove(ElementList2[2].Coordinates, botonesNavegadorDetalle2["Insertar"].X, botonesNavegadorDetalle2["Insertar"].Y);
                                            desktopSession.Mouse.Click(null);

                                            CrudKfdplcur.CrudDetalle2(desktopSession, 1, crudDetalle1, file);
                                            newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 2", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList2[2].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 2", true, file);
                                        }

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //4->Informacion contractual
                                        CrudKfdplcur.ClickButtonExterno(desktopSession, 4);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 3
                                        Dictionary<string, Point> botonesNavegadorDetalle3 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle3 = PruebaCRUD.findname(desktopSession, 27, 1);
                                        var ElementList3 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorDetalle3["Insertar"].X, botonesNavegadorDetalle3["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKfdplcur.CrudDetalle3(desktopSession, 1, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorDetalle3["Aplicar"].X, botonesNavegadorDetalle3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 3", true, file);

                                        // 5->Sesiones o Talleres
                                        CrudKfdplcur.ClickButtonExterno(desktopSession, 5);

                                        //1 -> Definición Sesión
                                        CrudKfdplcur.ClickButtonInterno(desktopSession, 1);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 4
                                        Dictionary<string, Point> botonesNavegadorDetalle4 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle4 = PruebaCRUD.findname(desktopSession, 28, 1);
                                        var ElementList4 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Insertar"].X, botonesNavegadorDetalle4["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKfdplcur.CrudDetalle4(desktopSession, 1, crudDetalle4, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 4", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Aplicar"].X, botonesNavegadorDetalle4["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 4", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Editar"].X, botonesNavegadorDetalle4["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 4", true, file);

                                        Thread.Sleep(1000);

                                        CrudKfdplcur.CrudDetalle4(desktopSession, 2, crudDetalle4, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 4", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Cancelar"].X, botonesNavegadorDetalle4["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 4", true, file);

                                        Thread.Sleep(1000);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Editar"].X, botonesNavegadorDetalle4["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 4", true, file);

                                        Thread.Sleep(1000);

                                        CrudKfdplcur.CrudDetalle4(desktopSession, 2, crudDetalle4, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 4", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Aplicar"].X, botonesNavegadorDetalle4["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 5", true, file);

                                        //2->Logística
                                        CrudKfdplcur.ClickButtonInterno(desktopSession, 2);

                                         //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 5
                                        Dictionary<string, Point> botonesNavegadorDetalle5 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle5 = PruebaCRUD.findname(desktopSession, 30, 1);
                                        var ElementList5 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorDetalle5["Insertar"].X, botonesNavegadorDetalle5["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKfdplcur.CrudDetalle5(desktopSession, 1, crudDetalle5, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 5", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorDetalle5["Aplicar"].X, botonesNavegadorDetalle5["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 5", true, file);

                                        Thread.Sleep(2000);

                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorDetalle5["Editar"].X, botonesNavegadorDetalle5["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 5", true, file);

                                        Thread.Sleep(1000);

                                        CrudKfdplcur.CrudDetalle5(desktopSession, 2, crudDetalle5, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 5", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorDetalle5["Cancelar"].X, botonesNavegadorDetalle5["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 5", true, file);

                                        Thread.Sleep(1000);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorDetalle5["Editar"].X, botonesNavegadorDetalle5["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 5", true, file);

                                        Thread.Sleep(1000);

                                        CrudKfdplcur.CrudDetalle5(desktopSession, 2, crudDetalle5, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 5", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorDetalle5["Aplicar"].X, botonesNavegadorDetalle5["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 5", true, file);

                                        CrudKfdplcur.Ventanana(desktopSession);

                                        //3 -> Inscritos
                                        CrudKfdplcur.ClickButtonInterno(desktopSession, 3);
                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 6
                                        Dictionary<string, Point> botonesNavegadorDetalle6 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle6 = PruebaCRUD.findname(desktopSession, 30, 1);
                                        var ElementList6 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegadorDetalle6["Insertar"].X, botonesNavegadorDetalle6["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKfdplcur.CrudDetalle6(desktopSession, 1, crudDetalle6, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 6", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegadorDetalle6["Aplicar"].X, botonesNavegadorDetalle6["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 6", true, file);

                                        Thread.Sleep(2000);

                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegadorDetalle6["Editar"].X, botonesNavegadorDetalle6["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 6", true, file);
   
                                        Thread.Sleep(1000);

                                        CrudKfdplcur.CrudDetalle6(desktopSession, 2, crudDetalle6, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 6", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegadorDetalle6["Cancelar"].X, botonesNavegadorDetalle6["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 6", true, file);

                                        Thread.Sleep(1000);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegadorDetalle6["Editar"].X, botonesNavegadorDetalle6["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 6", true, file);

                                        Thread.Sleep(1000);

                                        CrudKfdplcur.CrudDetalle6(desktopSession, 2, crudDetalle6, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 6", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegadorDetalle6["Aplicar"].X, botonesNavegadorDetalle6["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 6", true, file);

                                        //4 -> Documentos
                                        CrudKfdplcur.ClickButtonInterno(desktopSession, 4);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 7
                                        Dictionary<string, Point> botonesNavegadorDetalle7 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle7 = CrudKfdplcur.findname1(desktopSession, 30, 2);
                                        var ElementList7 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorDetalle7["Insertar"].X, botonesNavegadorDetalle7["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKfdplcur.CrudDetalle7(desktopSession, 1, crudDetalle7, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 7", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorDetalle7["Aplicar"].X, botonesNavegadorDetalle7["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 7", true, file);

                                        // 6 -> Cronograma
                                        CrudKfdplcur.ClickButtonExterno(desktopSession, 6);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 8
                                        Dictionary<string, Point> botonesNavegadorDetalle8 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle8 = PruebaCRUD.findname(desktopSession, 27, 1);
                                        var ElementList8 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegadorDetalle8["Insertar"].X, botonesNavegadorDetalle8["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKfdplcur.CrudDetalle8(desktopSession, 1, crudDetalle8, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 8", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegadorDetalle8["Aplicar"].X, botonesNavegadorDetalle8["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 8", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegadorDetalle8["Editar"].X, botonesNavegadorDetalle8["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 8", true, file);

                                        Thread.Sleep(1000);

                                        CrudKfdplcur.CrudDetalle8(desktopSession, 2, crudDetalle8, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 8", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegadorDetalle8["Cancelar"].X, botonesNavegadorDetalle8["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 8", true, file);

                                        Thread.Sleep(1000);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegadorDetalle8["Editar"].X, botonesNavegadorDetalle8["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 8", true, file);
      
                                        Thread.Sleep(1000);

                                        CrudKfdplcur.CrudDetalle8(desktopSession, 2, crudDetalle8, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 8", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegadorDetalle8["Aplicar"].X, botonesNavegadorDetalle8["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 8", true, file);

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
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //RECTIFICACION DE CAMPOS DE EXCEL  
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, Lugar, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        // 6 -> Cronograma
                                        CrudKfdplcur.ClickButtonExterno(desktopSession, 6);

                                        //ELIMINAR DETALLE 8
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList8[1], botonesNavegadorDetalle8, file);

                                        // 5->Sesiones o Talleres
                                        CrudKfdplcur.ClickButtonExterno(desktopSession, 5);
                                        //4 -> Documentos
                                        CrudKfdplcur.ClickButtonInterno(desktopSession, 4);

                                        //ELIMINAR DETALLE 7
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList7[2], botonesNavegadorDetalle7, file);

                                        //3 -> Inscritos
                                        CrudKfdplcur.ClickButtonInterno(desktopSession, 3);

                                        //ELIMINAR DETALLE 6
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList6[1], botonesNavegadorDetalle6, file);

                                        //2->Logística
                                        CrudKfdplcur.ClickButtonInterno(desktopSession, 2);

                                        //ELIMINAR DETALLE 5
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList5[1], botonesNavegadorDetalle5, file);

                                        //1 -> Definición Sesión
                                        CrudKfdplcur.ClickButtonInterno(desktopSession, 1);

                                        //ELIMINAR DETALLE 4
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList4[1], botonesNavegadorDetalle4, file);

                                        //4->Informacion contractual
                                        CrudKfdplcur.ClickButtonExterno(desktopSession, 4);

                                        //ELIMINAR DETALLE 3
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList3[1], botonesNavegadorDetalle3, file);

                                        //3->Evaluación curso
                                        CrudKfdplcur.ClickButtonExterno(desktopSession, 3);

                                        //ELIMINAR DETALLE 2
                                        //PruebaCRUD.EliminarRegistro(desktopSession, ElementList2[2], botonesNavegadorDetalle2, file);
                                        //ELIMINAR DETALLE 1
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[1], botonesNavegadorDetalle, file);
                                        CrudKfdplcur.Ventanana(desktopSession);

                                        //1 -> Planeación Curso
                                        CrudKfdplcur.ClickButtonExterno(desktopSession, 1);

                                        //ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        Thread.Sleep(500);

                                        ////Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Lugar, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Lugar, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);
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
        public void AbrirPrograma1()
        {
            //string motor = "SQL";
            //AbrirPrograma a = new AbrirPrograma("KFdPlcur", "UCalida1");
            string motor = "ORA";
            AbrirPrograma a = new AbrirPrograma("KFdPlcur", "ODESAR");
            desktopSession = a.Start2(motor, "");

            Thread.Sleep(20000000);
            //CrudKgnwebse.CRUD(desktopSession);
            


            string file = CrearDocumentoWordDinamico("KEd3Teva", "PruebasCrud", motor);
            //newFunctions_3.LupaAud(desktopSession, "1", file);
            //string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
            //CrudKfdplcur.ReportePreliminar(desktopSession, "0", file, pdf1,1);

            List<string> crudVariables = new List<string> { "12", "12", "7", "9", "10", "10", "10", "05/03/2017", "28/08/2017", "I", "PRUEBA LUGAR", "PRUEBA OBSERVACION", "EDITAR PRUEBA OBSERVACION" };
            //List<string> crudVariables = new List<string> { "14312", "14312", "3", "2014", "2", "2", "2", "05/03/2017", "28/08/2017", "I", "PRUEBA LUGAR", "PRUEBA OBSERVACION", "EDITAR PRUEBA OBSERVACION" };


            //CrudKgnrpsut.CRUD(desktopSession,crudVariables,file);

            //COnfiguro el QBE
            //FuncionesGlobales.QBEQry(desktopSession, "0", "7", file);

            //Thread.Sleep(1000);
            //string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
            //string ruta = @"C:\Reportes\ExportarExcel_" + "GN" + "_" + Hora();
            //FuncionesGlobales.ReportePreliminar(desktopSession, "1", file, pdf1);

            //newFunctions_4.BotonAceptar(desktopSession, pdf1, file);

            ////AGREGAR REGISTRO
            Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
            botonesNavegador = PruebaCRUD.findname(desktopSession, 27, 0);
            var ElementList = PruebaCRUD.NavClass(desktopSession);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            desktopSession.Mouse.Click(null);

            CrudKfdplcur.CRUD(desktopSession, 1, crudVariables, file);
            ////CrudKgnmseri.CRUD(desktopSession, 1, crudVariables, file);

            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);

            CrudKfdplcur.CRUD(desktopSession, 2, crudVariables, file);
            //////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);


            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);

            ////ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);

            CrudKfdplcur.CRUD(desktopSession, 2, crudVariables, file);
            //////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);

            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);

            

            //bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
            //if (valEditOra == true)
            //{
            //    //LUPA
            //    newFunctions_3.LupaAud(desktopSession, "1", file);

            //    //ACEPTAR EDICION
            //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            //    desktopSession.Mouse.Click(null);

            //    //CrudKgntmarc.CRUD(desktopSession, 2, crudVariables, file);
            //    CrudKslarxcc.CRUD(desktopSession, 2, crudVariables, file);

            //    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            //    desktopSession.Mouse.Click(null);

            //}

            // 1 - CLick en recursos para el campo, 2 - CLick en IdentificaciÃƒÆ’Ã‚Â³n
            //CrudKslarxcc.ClickButton(desktopSession, 1);
            //-------------------------------------------------------------------------------------
            //3->EvaluaciÃ³n curso
            CrudKfdplcur.ClickButtonExterno(desktopSession, 3);
            //Thread.Sleep(20000000);

            Thread.Sleep(1000);
            List<string> crudDetalle1 = new List<string> { "1", "I", "A", "S", "40", "05/06/2017", "05/07/2017", "5", "8" };
            ////AGREGAR DETALLE 1
            Dictionary<string, Point> botonesNavegadorDetalle = new Dictionary<string, Point>();
            botonesNavegadorDetalle = PruebaCRUD.findname(desktopSession, 27, 1);
            var ElementList1 = PruebaCRUD.NavClass(desktopSession);
            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Insertar"].X, botonesNavegadorDetalle["Insertar"].Y);
            desktopSession.Mouse.Click(null);

            CrudKfdplcur.CrudDetalle1(desktopSession, 1, crudDetalle1, file);

            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
            desktopSession.Mouse.Click(null);

            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);

            CrudKfdplcur.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
            ////////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);


            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Cancelar"].X, botonesNavegadorDetalle["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);

            ////ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);

            CrudKfdplcur.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
            //////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);

            desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            //Thread.Sleep(20000000);


            if (motor == "SQL")
            {
                Thread.Sleep(1000);
                List<string> crudDetalle2 = new List<string> { "1" };
                ////AGREGAR DETALLE 2
                Dictionary<string, Point> botonesNavegadorDetalle2 = new Dictionary<string, Point>();
                botonesNavegadorDetalle2 = PruebaCRUD.findname(desktopSession, 27, 2);
                var ElementList2 = PruebaCRUD.NavClass(desktopSession);
                desktopSession.Mouse.MouseMove(ElementList2[2].Coordinates, botonesNavegadorDetalle2["Insertar"].X, botonesNavegadorDetalle2["Insertar"].Y);
                desktopSession.Mouse.Click(null);

                CrudKfdplcur.CrudDetalle2(desktopSession, 1, crudDetalle1, file);

                desktopSession.Mouse.MouseMove(ElementList2[2].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                desktopSession.Mouse.Click(null);

                /*desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

                CrudKfdplcur.CrudDetalle2(desktopSession, 2, crudDetalle1, file);
                ////////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);


                desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDetalle2["Cancelar"].X, botonesNavegadorDetalle2["Cancelar"].Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

                ////ACEPTAR EDICION
                desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

                CrudKfdplcur.CrudDetalle2(desktopSession, 2, crudDetalle1, file);
                //////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);

                desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                desktopSession.Mouse.Click(null);
            //Thread.Sleep(20000000);*/
        }

            ////// LUPA
            newFunctions_3.LupaAud(desktopSession, "1", file);

            //4->Informacion contractual
            CrudKfdplcur.ClickButtonExterno(desktopSession, 4);


            Thread.Sleep(1000);
            ////AGREGAR DETALLE 3
            Dictionary<string, Point> botonesNavegadorDetalle3 = new Dictionary<string, Point>();
            botonesNavegadorDetalle3 = PruebaCRUD.findname(desktopSession, 27, 1);
            var ElementList3 = PruebaCRUD.NavClass(desktopSession);
            desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorDetalle3["Insertar"].X, botonesNavegadorDetalle3["Insertar"].Y);
            desktopSession.Mouse.Click(null);

            CrudKfdplcur.CrudDetalle3(desktopSession, 1, file);

            desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorDetalle3["Aplicar"].X, botonesNavegadorDetalle3["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            //------------------------------------------------------------------------------------------

            /*desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);

            CrudKfdplcur.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
            ////////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);


            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Cancelar"].X, botonesNavegadorDetalle["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);

            ////ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);

            CrudKfdplcur.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
            //////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);

            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
            desktopSession.Mouse.Click(null);*/
            //Thread.Sleep(20000000);
            //-----------------------------------------------------------------------------------------
             // 5->Sesiones o Talleres
             CrudKfdplcur.ClickButtonExterno(desktopSession, 5);

             //1 -> Definición Sesión
             CrudKfdplcur.ClickButtonInterno(desktopSession, 1);

             Thread.Sleep(1000);
             ////AGREGAR DETALLE 4
             List<string> crudDetalle4 = new List<string> { "1", "1", "1", "PRUEBA LUGAR", "5" };
             Dictionary<string, Point> botonesNavegadorDetalle4 = new Dictionary<string, Point>();
             botonesNavegadorDetalle4 = PruebaCRUD.findname(desktopSession, 28, 1);
             var ElementList4 = PruebaCRUD.NavClass(desktopSession);
             desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Insertar"].X, botonesNavegadorDetalle4["Insertar"].Y);
             desktopSession.Mouse.Click(null);

             CrudKfdplcur.CrudDetalle4(desktopSession, 1, crudDetalle4, file);

             desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Aplicar"].X, botonesNavegadorDetalle4["Aplicar"].Y);
             desktopSession.Mouse.Click(null);

             desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Editar"].X, botonesNavegadorDetalle4["Editar"].Y);
             desktopSession.Mouse.Click(null);
             Thread.Sleep(1000);

             CrudKfdplcur.CrudDetalle4(desktopSession, 2, crudDetalle4, file);
             ////////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);


             desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Cancelar"].X, botonesNavegadorDetalle4["Cancelar"].Y);
             desktopSession.Mouse.Click(null);
             Thread.Sleep(1000);

             ////ACEPTAR EDICION
             desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Editar"].X, botonesNavegadorDetalle4["Editar"].Y);
             desktopSession.Mouse.Click(null);
             Thread.Sleep(1000);

             CrudKfdplcur.CrudDetalle4(desktopSession, 2, crudDetalle4, file);
             //////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);

             desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Aplicar"].X, botonesNavegadorDetalle4["Aplicar"].Y);
             desktopSession.Mouse.Click(null);
             //Thread.Sleep(20000000);

             //2->LogÃ­stica
             CrudKfdplcur.ClickButtonInterno(desktopSession, 2);

             ////// LUPA
             newFunctions_3.LupaAud(desktopSession, "1", file);

             Thread.Sleep(1000);
             ////AGREGAR DETALLE 5
             List<string> crudDetalle5 = new List<string> { "10", "2", "S", "S", "N" };
             Dictionary<string, Point> botonesNavegadorDetalle5 = new Dictionary<string, Point>();
             botonesNavegadorDetalle5 = PruebaCRUD.findname(desktopSession, 30, 1);
             var ElementList5 = PruebaCRUD.NavClass(desktopSession);
             desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorDetalle5["Insertar"].X, botonesNavegadorDetalle5["Insertar"].Y);
             desktopSession.Mouse.Click(null);

             CrudKfdplcur.CrudDetalle5(desktopSession, 1, crudDetalle5, file);

             desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorDetalle5["Aplicar"].X, botonesNavegadorDetalle5["Aplicar"].Y);
             desktopSession.Mouse.Click(null);
             Thread.Sleep(2000);

             desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorDetalle5["Editar"].X, botonesNavegadorDetalle5["Editar"].Y);
             desktopSession.Mouse.Click(null);
             Thread.Sleep(1000);

             CrudKfdplcur.CrudDetalle5(desktopSession, 2, crudDetalle5, file);
             ////////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);


             desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorDetalle5["Cancelar"].X, botonesNavegadorDetalle5["Cancelar"].Y);
             desktopSession.Mouse.Click(null);
             Thread.Sleep(1000);

             ////ACEPTAR EDICION
             desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorDetalle5["Editar"].X, botonesNavegadorDetalle5["Editar"].Y);
             desktopSession.Mouse.Click(null);
             Thread.Sleep(1000);

             CrudKfdplcur.CrudDetalle5(desktopSession, 2, crudDetalle5, file);
             //////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);

             desktopSession.Mouse.MouseMove(ElementList5[1].Coordinates, botonesNavegadorDetalle5["Aplicar"].X, botonesNavegadorDetalle5["Aplicar"].Y);
             desktopSession.Mouse.Click(null);

             CrudKfdplcur.Ventanana(desktopSession);

             //3 -> Inscritos
             CrudKfdplcur.ClickButtonInterno(desktopSession, 3);
             ////// LUPA
             newFunctions_3.LupaAud(desktopSession, "1", file);

             Thread.Sleep(1000);
             ////AGREGAR DETALLE 6
             List<string> crudDetalle6 = new List<string> { "12/06/2017", "15/06/2017" };
             Dictionary<string, Point> botonesNavegadorDetalle6 = new Dictionary<string, Point>();
             botonesNavegadorDetalle6 = PruebaCRUD.findname(desktopSession, 30, 1);
             var ElementList6 = PruebaCRUD.NavClass(desktopSession);
             desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegadorDetalle6["Insertar"].X, botonesNavegadorDetalle6["Insertar"].Y);
             desktopSession.Mouse.Click(null);

             CrudKfdplcur.CrudDetalle6(desktopSession, 1, crudDetalle6, file);

             desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegadorDetalle6["Aplicar"].X, botonesNavegadorDetalle6["Aplicar"].Y);
             desktopSession.Mouse.Click(null);
             Thread.Sleep(2000);

             desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegadorDetalle6["Editar"].X, botonesNavegadorDetalle6["Editar"].Y);
             desktopSession.Mouse.Click(null);
             Thread.Sleep(1000);

             CrudKfdplcur.CrudDetalle6(desktopSession, 2, crudDetalle6, file);
             ////////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);


             desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegadorDetalle6["Cancelar"].X, botonesNavegadorDetalle6["Cancelar"].Y);
             desktopSession.Mouse.Click(null);
             Thread.Sleep(1000);

             ////ACEPTAR EDICION
             desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegadorDetalle6["Editar"].X, botonesNavegadorDetalle6["Editar"].Y);
             desktopSession.Mouse.Click(null);
             Thread.Sleep(1000);

             CrudKfdplcur.CrudDetalle6(desktopSession, 2, crudDetalle6, file);
             //////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);

             desktopSession.Mouse.MouseMove(ElementList6[1].Coordinates, botonesNavegadorDetalle6["Aplicar"].X, botonesNavegadorDetalle6["Aplicar"].Y);
             desktopSession.Mouse.Click(null);

             //Thread.Sleep(1000000000);
             //4 -> Documentos
             CrudKfdplcur.ClickButtonInterno(desktopSession, 4);

             Thread.Sleep(1000);
             ////AGREGAR DETALLE 7
             List<string> crudDetalle7 = new List<string> { "NOMBRE DOCUMENTO", "1" };
             Dictionary<string, Point> botonesNavegadorDetalle7 = new Dictionary<string, Point>();
             botonesNavegadorDetalle7 = CrudKfdplcur.findname1(desktopSession, 30, 2);
             var ElementList7 = PruebaCRUD.NavClass(desktopSession);
             desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorDetalle7["Insertar"].X, botonesNavegadorDetalle7["Insertar"].Y);
             desktopSession.Mouse.Click(null);

             CrudKfdplcur.CrudDetalle7(desktopSession, 1, crudDetalle7, file);

             desktopSession.Mouse.MouseMove(ElementList7[2].Coordinates, botonesNavegadorDetalle7["Aplicar"].X, botonesNavegadorDetalle7["Aplicar"].Y);
             desktopSession.Mouse.Click(null);

             // 6 -> Cronograma
             CrudKfdplcur.ClickButtonExterno(desktopSession, 6);

             Thread.Sleep(1000);
             ////AGREGAR DETALLE 8
             List<string> crudDetalle8 = new List<string> { "25/06/2017", "1", "28/06/2017", "30/06/2017" };
             Dictionary<string, Point> botonesNavegadorDetalle8 = new Dictionary<string, Point>();
             botonesNavegadorDetalle8 = PruebaCRUD.findname(desktopSession, 27, 1);
             var ElementList8 = PruebaCRUD.NavClass(desktopSession);
             desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegadorDetalle8["Insertar"].X, botonesNavegadorDetalle8["Insertar"].Y);
             desktopSession.Mouse.Click(null);

             CrudKfdplcur.CrudDetalle8(desktopSession, 1, crudDetalle8, file);

             desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegadorDetalle8["Aplicar"].X, botonesNavegadorDetalle8["Aplicar"].Y);
             desktopSession.Mouse.Click(null);

             desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegadorDetalle8["Editar"].X, botonesNavegadorDetalle8["Editar"].Y);
             desktopSession.Mouse.Click(null);
             Thread.Sleep(1000);

             CrudKfdplcur.CrudDetalle8(desktopSession, 2, crudDetalle8, file);
             ////////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);


             desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegadorDetalle8["Cancelar"].X, botonesNavegadorDetalle8["Cancelar"].Y);
             desktopSession.Mouse.Click(null);
             Thread.Sleep(1000);

             ////ACEPTAR EDICION
             desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegadorDetalle8["Editar"].X, botonesNavegadorDetalle8["Editar"].Y);
             desktopSession.Mouse.Click(null);
             Thread.Sleep(1000);

             CrudKfdplcur.CrudDetalle8(desktopSession, 2, crudDetalle8, file);
             //////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);

             desktopSession.Mouse.MouseMove(ElementList8[1].Coordinates, botonesNavegadorDetalle8["Aplicar"].X, botonesNavegadorDetalle8["Aplicar"].Y);
             desktopSession.Mouse.Click(null);
             //Thread.Sleep(20000000);


             ////// LUPA
             newFunctions_3.LupaAud(desktopSession, "1", file);

             // 6 -> Cronograma
             CrudKfdplcur.ClickButtonExterno(desktopSession, 6);
             //ELIMINAR DETALLE 8
             PruebaCRUD.EliminarRegistro(desktopSession, ElementList8[1], botonesNavegadorDetalle8, file);

             // 5->Sesiones o Talleres
             CrudKfdplcur.ClickButtonExterno(desktopSession, 5);
             //4 -> Documentos
             CrudKfdplcur.ClickButtonInterno(desktopSession, 4);
             //ELIMINAR DETALLE 7
             PruebaCRUD.EliminarRegistro(desktopSession, ElementList7[2], botonesNavegadorDetalle7, file);

             //3 -> Inscritos
             CrudKfdplcur.ClickButtonInterno(desktopSession, 3);
             //ELIMINAR DETALLE 6
             PruebaCRUD.EliminarRegistro(desktopSession, ElementList6[1], botonesNavegadorDetalle6, file);

             //2->LogÃ­stica
             CrudKfdplcur.ClickButtonInterno(desktopSession, 2);
             //ELIMINAR DETALLE 5
             PruebaCRUD.EliminarRegistro(desktopSession, ElementList5[1], botonesNavegadorDetalle5, file);

             //1 -> DefiniciÃ³n SesiÃ³n
             CrudKfdplcur.ClickButtonInterno(desktopSession, 1);
             //ELIMINAR DETALLE 4
             PruebaCRUD.EliminarRegistro(desktopSession, ElementList4[1], botonesNavegadorDetalle4, file);

             //4->Informacion contractual
             CrudKfdplcur.ClickButtonExterno(desktopSession, 4);
             //ELIMINAR DETALLE 3
             PruebaCRUD.EliminarRegistro(desktopSession, ElementList3[1], botonesNavegadorDetalle3, file);

             //3->EvaluaciÃ³n curso
             CrudKfdplcur.ClickButtonExterno(desktopSession, 3);
             //ELIMINAR DETALLE 2
             //PruebaCRUD.EliminarRegistro(desktopSession, ElementList2[2], botonesNavegadorDetalle2, file);
             //ELIMINAR DETALLE 1
             PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[1], botonesNavegadorDetalle, file);
             CrudKfdplcur.Ventanana(desktopSession);

             //1 -> PlaneaciÃ³n Curso
             CrudKfdplcur.ClickButtonExterno(desktopSession, 1);
             //ELIMINAR REGISTRO
             PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);

             
            //-------------------------------------------------------------------------------------

            ////CLick Button
            //CrudKslarxcc.ClickButton(desktopSession, 2);

            //-------------------------------------------------------------------------------------

            /*  ////// LUPA
              newFunctions_3.LupaAud(desktopSession, "1", file);
              Thread.Sleep(1000);
              List<string> crudDetalle2 = new List<string> { "1", "2" };
              ////AGREGAR DETALLE 2
              Dictionary<string, Point> botonesNavegadorDetalle2 = new Dictionary<string, Point>();
              botonesNavegadorDetalle2 = PruebaCRUD.findname(desktopSession, 27, 2);
              ElementList = PruebaCRUD.NavClass(desktopSession);
              desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDetalle2["Insertar"].X, botonesNavegadorDetalle2["Insertar"].Y);
              desktopSession.Mouse.Click(null);

              CrudKed3perf.CrudDetalle23Perf(desktopSession, 1, crudDetalle2, file);
              Thread.Sleep(1000);
              desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
              desktopSession.Mouse.Click(null);

              desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
              desktopSession.Mouse.Click(null);
              Thread.Sleep(1000);

              CrudKed3perf.CrudDetalle23Perf(desktopSession, 2, crudDetalle2, file);
              ////////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);


              desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDetalle2["Cancelar"].X, botonesNavegadorDetalle2["Cancelar"].Y);
              desktopSession.Mouse.Click(null);
              Thread.Sleep(1000);

              ////ACEPTAR EDICION
              desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
              desktopSession.Mouse.Click(null);
              Thread.Sleep(1000);

              CrudKed3perf.CrudDetalle23Perf(desktopSession, 2, crudDetalle2, file);
              //////////CrudKgnwayud.CRUD(desktopSession, 2, crudVariables,file);

              desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
              desktopSession.Mouse.Click(null);

              //-------------------------------------------------------------------------------------
              ////Click Button
              //CrudKslarxcc.ClickButton(desktopSession, 3);

              //Thread.Sleep(1000000);
              //////AGREGAR DETALLE 3
              //botonesNavegadorDetalle = PruebaCRUD.findname1(desktopSession, 27, 1);
              //ElementList = PruebaCRUD.NavClass(desktopSession);
              //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Insertar"].X, botonesNavegadorDetalle["Insertar"].Y);
              //desktopSession.Mouse.Click(null);

              ////EL numero 1 hace referencia al detalle 1, el 2 al detalle 2 y el 3 al detalle 3
              //CrudKslarxcc.CrudDetalle(desktopSession, 3, file);

              //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
              //desktopSession.Mouse.Click(null);

              //COnfiguro el QBE
              //FuncionesGlobales.QBEQry(desktopSession, "0", "", file);

              //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
              //desktopSession.Mouse.Click(null);

              //CrudKgnmseri.CrudDetalle1(desktopSession, 2, crudDetalle1, file);

              //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
              //desktopSession.Mouse.Click(null);

              //////ACEPTAR EDICION
              //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
              //desktopSession.Mouse.Click(null);

              //CrudKgnmseri.CrudDetalle1(desktopSession, 2, crudDetalle1, file);

              //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
              //desktopSession.Mouse.Click(null);

              //////////CrudKfdcurso.EliminarDetalle(desktopSession, ElementList[1], botonesNavegador, file);
              //////// LUPA
              //////newFunctions_3.LupaAud(desktopSession, "1", file);
              /////
              //-------------------------------------------------------------------------------------

              ////// LUPA
              newFunctions_3.LupaAud(desktopSession, "1", file);
              //ELIMINAR DETALLE 2
              PruebaCRUD.EliminarRegistro(desktopSession, ElementList[2], botonesNavegadorDetalle2, file);
              Thread.Sleep(500);
              ////// LUPA
              newFunctions_3.LupaAud(desktopSession, "1", file);
              //ELIMINAR DETALLE 1
              PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorDetalle, file);
              Thread.Sleep(500);
              ////// LUPA
              newFunctions_3.LupaAud(desktopSession, "1", file);
              //ELIMINAR REGISTRO
              PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);*/
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

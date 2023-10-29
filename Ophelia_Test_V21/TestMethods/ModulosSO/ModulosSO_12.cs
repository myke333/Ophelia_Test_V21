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
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;

namespace Ophelia_Test_V21.TestMethods.ModulosSO
{
    /// <summary>
    /// Descripción resumida de ModulosSO_12
    /// </summary>
    [TestClass]
    public class ModulosSO_12 : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;

        public ModulosSO_12() { }

        [TestMethod]
        public void KSoActinCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_12.KSoActinCheckList")
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
                       
                                rows["Codigo"].ToString().Length != 0 && rows["Codigo"].ToString() != null &&
                                rows["NombreActo"].ToString().Length != 0 && rows["NombreActo"].ToString() != null &&
                                rows["Descripcion"].ToString().Length != 0 && rows["Descripcion"].ToString() != null &&
                                rows["EditarDescripcion"].ToString().Length != 0 && rows["EditarDescripcion"].ToString() != null
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

                                string Codigo = rows["Codigo"].ToString();
                                string NombreActo = rows["NombreActo"].ToString();
                                string Descripcion = rows["Descripcion"].ToString();
                                string EditarDescripcion = rows["EditarDescripcion"].ToString();

                                List<string> crudVariables = new List<string>() { Codigo, NombreActo, Descripcion, EditarDescripcion };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);
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
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, Campo, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

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

                                        CrudKsoactin.CRUD(desktopSession, 1, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKsoactin.CRUD(desktopSession, 2, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Descripcion, "DES_ACTO", c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKsoactin.CRUD(desktopSession, 2, crudVariables, file);
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

                                            CrudKsoactin.CRUD(desktopSession, 2, crudVariables, file);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, EditarDescripcion, "DES_ACTO", c, ErrorValidacion, "El registro no se editó correctamente", 1);

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
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, Codigo, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        Thread.Sleep(1000);

                                        ////Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);
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
        public void KSoEpactCheckList()
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
            ///TableOrder = "ktes1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_12.KSoEpactCheckList")
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
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&
                                //CRUD
                                rows["CodigoInterno"].ToString().Length != 0 && rows["CodigoInterno"].ToString() != null &&
                                rows["NumeroInforme"].ToString().Length != 0 && rows["NumeroInforme"].ToString() != null &&
                                rows["FechaAtel"].ToString().Length != 0 && rows["FechaAtel"].ToString() != null &&
                                rows["HoraAccidente"].ToString().Length != 0 && rows["HoraAccidente"].ToString() != null &&
                                rows["EditarNumInforme"].ToString().Length != 0 && rows["EditarNumInforme"].ToString() != null &&
                                //DETALLE
                                rows["ParteCuerpo"].ToString().Length != 0 && rows["ParteCuerpo"].ToString() != null &&
                                rows["AgenteAccidente"].ToString().Length != 0 && rows["AgenteAccidente"].ToString() != null &&
                                rows["FormaAccidente"].ToString().Length != 0 && rows["FormaAccidente"].ToString() != null &&
                                rows["EditFormaAccidente"].ToString().Length != 0 && rows["EditFormaAccidente"].ToString() != null &&
                                //DETALLE 0
                                rows["TipoLesion"].ToString().Length != 0 && rows["TipoLesion"].ToString() != null &&
                                rows["Observacion"].ToString().Length != 0 && rows["Observacion"].ToString() != null &&
                                rows["EditObservacion"].ToString().Length != 0 && rows["EditObservacion"].ToString() != null &&
                                //DETALLE 1
                                rows["FactorRiesgo"].ToString().Length != 0 && rows["FactorRiesgo"].ToString() != null &&
                                rows["DetalleFactor"].ToString().Length != 0 && rows["DetalleFactor"].ToString() != null &&
                                rows["EditarDetalleFactor"].ToString().Length != 0 && rows["EditarDetalleFactor"].ToString() != null &&
                                //DETALLE 2
                                rows["ParteAfectada"].ToString().Length != 0 && rows["ParteAfectada"].ToString() != null &&
                                rows["Lado"].ToString().Length != 0 && rows["Lado"].ToString() != null &&
                                rows["EditarLado"].ToString().Length != 0 && rows["EditarLado"].ToString() != null &&
                                //DETALLE 3
                                rows["CodActoInseguro"].ToString().Length != 0 && rows["CodActoInseguro"].ToString() != null &&
                                //DETALLE 4
                                rows["Recomendaciones"].ToString().Length != 0 && rows["Recomendaciones"].ToString() != null &&
                                rows["EditRecomendaciones"].ToString().Length != 0 && rows["EditRecomendaciones"].ToString() != null &&


                                //VALIDACIONES
                                rows["tablaDetalle1"].ToString().Length != 0 && rows["tablaDetalle1"].ToString() != null &&
                                rows["campoDetalle1"].ToString().Length != 0 && rows["campoDetalle1"].ToString() != null &&
                                rows["editCampoD1"].ToString().Length != 0 && rows["editCampoD1"].ToString() != null &&

                                rows["tablaDetalle2"].ToString().Length != 0 && rows["tablaDetalle2"].ToString() != null &&
                                rows["campoDetalle2"].ToString().Length != 0 && rows["campoDetalle2"].ToString() != null &&
                                rows["editCampoD2"].ToString().Length != 0 && rows["editCampoD2"].ToString() != null &&

                                rows["tablaDetalle3"].ToString().Length != 0 && rows["tablaDetalle3"].ToString() != null &&
                                rows["campoDetalle3"].ToString().Length != 0 && rows["campoDetalle3"].ToString() != null &&
                                rows["editCampoD3"].ToString().Length != 0 && rows["editCampoD3"].ToString() != null &&

                                rows["tablaDetalle4"].ToString().Length != 0 && rows["tablaDetalle4"].ToString() != null &&
                                rows["campoDetalle4"].ToString().Length != 0 && rows["campoDetalle4"].ToString() != null &&
                                rows["editCampoD4"].ToString().Length != 0 && rows["editCampoD4"].ToString() != null &&

                                rows["tablaDetalle5"].ToString().Length != 0 && rows["tablaDetalle5"].ToString() != null &&
                                rows["campoDetalle5"].ToString().Length != 0 && rows["campoDetalle5"].ToString() != null &&
                                //rows["editCampoD5"].ToString().Length != 0 && rows["editCampoD5"].ToString() != null &&

                                rows["tablaDetalle6"].ToString().Length != 0 && rows["tablaDetalle6"].ToString() != null &&
                                rows["campoDetalle6"].ToString().Length != 0 && rows["campoDetalle6"].ToString() != null &&
                                rows["editCampoD6"].ToString().Length != 0 && rows["editCampoD6"].ToString() != null &&

                                rows["Estado"].ToString().Length != 0 && rows["Estado"].ToString() != null 
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

                                string CodigoInterno = rows["CodigoInterno"].ToString();
                                string NumeroInforme = rows["NumeroInforme"].ToString();
                                string FechaAtel = rows["FechaAtel"].ToString();
                                string HoraAccidente = rows["HoraAccidente"].ToString();
                                string EditarNumInforme = rows["EditarNumInforme"].ToString();
                                
                                
                                string ParteCuerpo = rows["ParteCuerpo"].ToString();
                                string AgenteAccidente = rows["AgenteAccidente"].ToString();
                                string FormaAccidente = rows["FormaAccidente"].ToString();
                                string EditFormaAccidente = rows["EditFormaAccidente"].ToString();

                                string TipoLesion = rows["TipoLesion"].ToString();
                                string Observacion = rows["Observacion"].ToString();
                                string EditObservacion = rows["EditObservacion"].ToString();

                                string FactorRiesgo = rows["FactorRiesgo"].ToString();
                                string DetalleFactor = rows["DetalleFactor"].ToString();
                                string EditarDetalleFactor = rows["EditarDetalleFactor"].ToString();

                                string ParteAfectada = rows["ParteAfectada"].ToString();
                                string Lado = rows["Lado"].ToString();
                                string EditarLado = rows["EditarLado"].ToString();

                                string CodActoInseguro = rows["CodActoInseguro"].ToString();

                                string Recomendaciones = rows["Recomendaciones"].ToString();
                                string EditRecomendaciones = rows["EditRecomendaciones"].ToString();

                                string tablaDetalle1 = rows["tablaDetalle1"].ToString();
                                string campoDetalle1 = rows["campoDetalle1"].ToString();
                                string editCampoD1 = rows["editCampoD1"].ToString();
                                string tablaDetalle2 = rows["tablaDetalle2"].ToString();
                                string campoDetalle2 = rows["campoDetalle2"].ToString();
                                string editCampoD2 = rows["editCampoD2"].ToString();
                                string tablaDetalle3 = rows["tablaDetalle3"].ToString();
                                string campoDetalle3 = rows["campoDetalle3"].ToString();
                                string editCampoD3 = rows["editCampoD3"].ToString();
                                string tablaDetalle4 = rows["tablaDetalle4"].ToString();
                                string campoDetalle4 = rows["campoDetalle4"].ToString();
                                string editCampoD4 = rows["editCampoD4"].ToString();
                                string tablaDetalle5 = rows["tablaDetalle5"].ToString();
                                string campoDetalle5 = rows["campoDetalle5"].ToString();
                                //string editCampoD5 = rows["editCampoD5"].ToString();
                                string tablaDetalle6 = rows["tablaDetalle6"].ToString();
                                string campoDetalle6 = rows["campoDetalle6"].ToString();
                                string editCampoD6 = rows["editCampoD6"].ToString();

                                string Estado = rows["Estado"].ToString();


                                List<string> crudVariables = new List<string>() { CodigoInterno, NumeroInforme, FechaAtel, HoraAccidente, EditarNumInforme };
                                List<string> crudDetalle= new List<string>() { ParteCuerpo, AgenteAccidente, FormaAccidente, EditFormaAccidente};
                                List<string> crudDetalle0 = new List<string>() { TipoLesion, Observacion, EditObservacion };
                                List<string> crudDetalle1 = new List<string>() { FactorRiesgo, DetalleFactor, EditarDetalleFactor };
                                List<string> crudDetalle2 = new List<string>() { ParteAfectada, Lado, EditarLado };
                                List<string> crudDetalle3 = new List<string>() { CodActoInseguro };
                                List<string> crudDetalle4 = new List<string>() { Recomendaciones, EditRecomendaciones };


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);
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

                                        //ELIMINAR REGISTRO PEGADO DETALLE 6
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle6, user, motor, campoDetalle6, Estado, Estado, campoDetalle6, c, ErrorValidacion, "No se agregó correctamente 6", 0,"1");

                                        //ELIMINAR REGISTRO PEGADO DETALLE 5
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle5, user, motor, campoDetalle5, CodActoInseguro, CodActoInseguro, campoDetalle5, c, ErrorValidacion, "No se agregó correctamente 5", 0,"1");

                                        //ELIMINAR REGISTRO PEGADO DETALLE 4
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle4, user, motor, campoDetalle4, ParteAfectada, ParteAfectada, campoDetalle4, c, ErrorValidacion, "No se agregó correctamente 4", 0,"1");

                                        //ELIMINAR REGISTRO PEGADO DETALLE 3
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, FactorRiesgo, FactorRiesgo, campoDetalle3, c, ErrorValidacion, "No se agregó correctamente 3", 0,"1");

                                        //ELIMINAR REGISTRO PEGADO DETALLE 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, TipoLesion, TipoLesion, campoDetalle2, c, ErrorValidacion, "No se agregó correctamente 2", 0,"1");

                                        //ELIMINAR REGISTRO PEGADO DETALLE 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, ParteCuerpo, ParteCuerpo, campoDetalle1, c, ErrorValidacion, "No se agregó correctamente 1", 0,"1");

                                        //ELIMINAR REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CodigoInterno, CodigoInterno, Campo, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

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

                                        CrudKsoepact.CRUD(desktopSession, 1, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CodigoInterno, CodigoInterno, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKsoepact.CRUD(desktopSession, 2, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CodigoInterno, NumeroInforme, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKsoepact.CRUD(desktopSession, 2, crudVariables, file);
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

                                            CrudKsoepact.CRUD(desktopSession, 2, crudVariables, file);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, CodigoInterno, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CodigoInterno, EditarNumInforme, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);

                                        //2 -> Datos del Evento
                                        CrudKsoepact.ClickButtonExterno(desktopSession, 2);

                                        //1-> Parte de Cuerpo Afectada / Agente / Mecanismo del Accidente
                                        CrudKsoepact.ClickButtonInterno(desktopSession, 1);

                                        Thread.Sleep(1000);

                                        ////AGREGAR DETALLE 
                                        Dictionary<string, Point> botonesNavegadorDetalle = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle = PruebaCRUD.findname(desktopSession, 27, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Insertar"].X, botonesNavegadorDetalle["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKsoepact.CrudDetalle(desktopSession, 1, crudDetalle, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle", true, file);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, ParteCuerpo, ParteCuerpo, campoDetalle1, c, ErrorValidacion, "No se agregó correctamente 1", 0);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle", true, file);

                                        Thread.Sleep(1000);

                                        CrudKsoepact.CrudDetalle(desktopSession, 2, crudDetalle, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Cancelar"].X, botonesNavegadorDetalle["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle", true, file);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, ParteCuerpo, FormaAccidente, editCampoD1, c, ErrorValidacion, "No se edito correctamente 1", 1);

                                        Thread.Sleep(1000);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle", true, file);

                                        Thread.Sleep(1000);

                                        CrudKsoepact.CrudDetalle(desktopSession, 2, crudDetalle, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(3000);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, ParteCuerpo, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, ParteCuerpo, EditFormaAccidente, editCampoD1, c, ErrorValidacion, "El registro no se editó correctamente 1", 1);

                                        //2 -> Sitio y Lesión
                                        CrudKsoepact.ClickButtonInterno(desktopSession, 2);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 0
                                        Dictionary<string, Point> botonesNavegadorDetalle0 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle0 = PruebaCRUD.findname(desktopSession, 27, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle0["Insertar"].X, botonesNavegadorDetalle0["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKsoepact.CrudDetalle0(desktopSession, 1, crudDetalle0, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 0", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle0["Aplicar"].X, botonesNavegadorDetalle0["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 0", true, file);

                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, TipoLesion, TipoLesion, campoDetalle2, c, ErrorValidacion, "No se agregó correctamente 2", 0);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle0["Editar"].X, botonesNavegadorDetalle0["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 0", true, file);

                                        Thread.Sleep(1000);

                                        CrudKsoepact.CrudDetalle0(desktopSession, 2, crudDetalle0, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 0", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle0["Cancelar"].X, botonesNavegadorDetalle0["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 0", true, file);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, TipoLesion, Observacion, editCampoD2, c, ErrorValidacion, "No se edito correctamente 2", 1);

                                        Thread.Sleep(1000);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle0["Editar"].X, botonesNavegadorDetalle0["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 0", true, file);
  
                                        Thread.Sleep(1000);

                                        CrudKsoepact.CrudDetalle0(desktopSession, 2, crudDetalle0, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 0", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle0["Aplicar"].X, botonesNavegadorDetalle0["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 0", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(3000);
                                        ////Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, TipoLesion, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, TipoLesion, EditObservacion, editCampoD2, c, ErrorValidacion, "El registro no se editó correctamente 2", 1);

                                        //3 -> Adicionales
                                        CrudKsoepact.ClickButtonExterno(desktopSession, 3);

                                        //3-> Factores
                                        CrudKsoepact.ClickButtonInterno(desktopSession, 3);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 1
                                        Dictionary<string, Point> botonesNavegadorDetalle1 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle1 = PruebaCRUD.findname(desktopSession, 27, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle1["Insertar"].X, botonesNavegadorDetalle1["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKsoepact.CrudDetalle1(desktopSession, 1, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle1["Aplicar"].X, botonesNavegadorDetalle1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 1", true, file);

                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, FactorRiesgo, FactorRiesgo, campoDetalle3, c, ErrorValidacion, "No se agregó correctamente 3", 0);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle1["Editar"].X, botonesNavegadorDetalle1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                        Thread.Sleep(1000);

                                        CrudKsoepact.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle1["Cancelar"].X, botonesNavegadorDetalle1["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 1", true, file);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, FactorRiesgo, DetalleFactor, editCampoD3, c, ErrorValidacion, "No se edito correctamente 3", 1);

                                        Thread.Sleep(1000);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle1["Editar"].X, botonesNavegadorDetalle1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                        Thread.Sleep(1000);

                                        CrudKsoepact.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle1["Aplicar"].X, botonesNavegadorDetalle1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 1", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(3000);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle3, motor, user, campoDetalle3, FactorRiesgo, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, FactorRiesgo, EditarDetalleFactor, editCampoD3, c, ErrorValidacion, "El registro no se editó correctamente 3", 1);

                                        //4-> Partes Afectadas
                                        CrudKsoepact.ClickButtonInterno(desktopSession, 4);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 2
                                        Dictionary<string, Point> botonesNavegadorDetalle2 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle2 = PruebaCRUD.findname(desktopSession, 23, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Insertar"].X, botonesNavegadorDetalle2["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKsoepact.CrudDetalle2(desktopSession, 1, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 2", true, file);

                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle4, user, motor, campoDetalle4, ParteAfectada, ParteAfectada, campoDetalle4, c, ErrorValidacion, "No se agregó correctamente 4", 0);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);

                                        Thread.Sleep(1000);

                                        CrudKsoepact.CrudDetalle2(desktopSession, 2, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Cancelar"].X, botonesNavegadorDetalle2["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 2", true, file);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle4, user, motor, campoDetalle4, ParteAfectada, Lado, editCampoD4, c, ErrorValidacion, "No se edito correctamente 4", 1);

                                        Thread.Sleep(1000);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);

                                        Thread.Sleep(1000);

                                        CrudKsoepact.CrudDetalle2(desktopSession, 2, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 2", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(3000);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle4, motor, user, campoDetalle4, ParteAfectada, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle4, user, motor, campoDetalle4, ParteAfectada, EditarLado, editCampoD4, c, ErrorValidacion, "El registro no se editó correctamente 4", 1);

                                        //5 -> Actos Inseguros y/o Causas Accidente 
                                        CrudKsoepact.ClickButtonInterno(desktopSession, 5);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 3
                                        Dictionary<string, Point> botonesNavegadorDetalle3 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle3 = PruebaCRUD.findname(desktopSession, 27, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle3["Insertar"].X, botonesNavegadorDetalle3["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKsoepact.CrudDetalle3(desktopSession, 1, crudDetalle3, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle3["Aplicar"].X, botonesNavegadorDetalle3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 3", true, file);

                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle5, user, motor, campoDetalle5, CodActoInseguro, CodActoInseguro, campoDetalle5, c, ErrorValidacion, "No se agregó correctamente 5", 0);

                                        //4 -> Recomendacionesidente 
                                        CrudKsoepact.ClickButtonExterno(desktopSession, 4);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 4
                                        Dictionary<string, Point> botonesNavegadorDetalle4 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle4 = CrudKsoepact.findname(desktopSession, 33, 1);
                                        var ElementList4 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Insertar"].X, botonesNavegadorDetalle4["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKsoepact.CrudDetalle4(desktopSession, 1, crudDetalle4, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 4", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Aplicar"].X, botonesNavegadorDetalle4["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 4", true, file);

                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle6, user, motor, campoDetalle6, Estado, Estado, campoDetalle6, c, ErrorValidacion, "No se agregó correctamente 6", 0);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Editar"].X, botonesNavegadorDetalle4["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 4", true, file);

                                        Thread.Sleep(1000);

                                        CrudKsoepact.CrudDetalle4(desktopSession, 2, crudDetalle4, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 4", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Cancelar"].X, botonesNavegadorDetalle4["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 4", true, file);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle6, user, motor, campoDetalle6, Estado, Recomendaciones, editCampoD6, c, ErrorValidacion, "No se edito correctamente 6", 1);
 
                                        Thread.Sleep(1000);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Editar"].X, botonesNavegadorDetalle4["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 4", true, file);

                                        Thread.Sleep(1000);

                                        CrudKsoepact.CrudDetalle4(desktopSession, 2, crudDetalle4, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 4", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Aplicar"].X, botonesNavegadorDetalle4["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 4", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(3000);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle6, motor, user, campoDetalle6, Estado, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle6, user, motor, campoDetalle6, Estado, EditRecomendaciones, editCampoD6, c, ErrorValidacion, "El registro no se editó correctamente 6", 1);

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
                                        errors = CrudKsoepact.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1, 1);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        string pdf2 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = CrudKsoepact.ReportePreliminar(desktopSession, BanderaPreli, file, pdf2, 2);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        string pdf3 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = CrudKsoepact.ReportePreliminar(desktopSession, BanderaPreli, file, pdf3, 3);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        string pdf4 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = CrudKsoepact.ReportePreliminar(desktopSession, BanderaPreli, file, pdf4, 4);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //RECTIFICACION DE CAMPOS DE EXCEL  
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, CodigoInterno, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //3 -> Adicionales
                                        CrudKsoepact.ClickButtonExterno(desktopSession, 3);

                                        //3-> Factores
                                        CrudKsoepact.ClickButtonInterno(desktopSession, 3);

                                        //ELIMINAR DETALLE 1
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorDetalle1, file);

                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle3, motor, user, campoDetalle3, FactorRiesgo, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, FactorRiesgo, "", campoDetalle3, c, ErrorValidacion, "los datos no se eliminaron correctamente 3", 2);

                                        //4-> Partes Afectadas
                                        CrudKsoepact.ClickButtonInterno(desktopSession, 4);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR DETALLE 2
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorDetalle2, file);

                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle4, motor, user, campoDetalle4, ParteAfectada, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle4, user, motor, campoDetalle4, ParteAfectada, "", campoDetalle4, c, ErrorValidacion, "los datos no se eliminaron correctamente 4", 2);

                                        // 5 -> Actos Inseguros y/o Causas Accidente
                                        CrudKsoepact.ClickButtonInterno(desktopSession, 5);

                                        //ELIMINAR DETALLE 3
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorDetalle3, file);
                                        Thread.Sleep(500);

                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle5, motor, user, campoDetalle5, CodActoInseguro, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle5, user, motor, campoDetalle5, CodActoInseguro, "", campoDetalle5, c, ErrorValidacion, "los datos no se eliminaron correctamente 5", 2);

                                        //4 -> Recomendacionesidente 
                                        CrudKsoepact.ClickButtonExterno(desktopSession, 4);

                                        //ELIMINAR DETALLE 4
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList4[1], botonesNavegadorDetalle4, file);

                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle6, motor, user, campoDetalle6, Estado, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle6, user, motor, campoDetalle6, Estado, "", campoDetalle6, c, ErrorValidacion, "los datos no se eliminaron correctamente 6", 2);
                                        
                                        //2 -> Datos del Evento
                                        CrudKsoepact.ClickButtonExterno(desktopSession, 2);
                                        //1-> Parte de Cuerpo Afectada / Agente / Mecanismo del Accidente
                                        CrudKsoepact.ClickButtonInterno(desktopSession, 1);

                                        //ELIMINAR DETALLE 
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorDetalle, file);

                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, ParteCuerpo, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, ParteCuerpo, "", campoDetalle1, c, ErrorValidacion, "los datos no se eliminaron correctamente 1", 2);

                                        //2 -> Sitio y Lesión
                                        CrudKsoepact.ClickButtonInterno(desktopSession, 2);

                                        //ELIMINAR DETALLE  0
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorDetalle0, file);

                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, TipoLesion, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, TipoLesion, "", campoDetalle2, c, ErrorValidacion, "los datos no se eliminaron correctamente 2", 2);

                                        CrudKsoepact.ClickButtonExterno(desktopSession, 1);

                                        //ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        Thread.Sleep(500);

                                        ////Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, CodigoInterno, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, CodigoInterno, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);
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
        public void KSoTcomiCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_12.KSoTcomiCheckList")
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

                                rows["TipoComite"].ToString().Length != 0 && rows["TipoComite"].ToString() != null &&
                                rows["Nombre"].ToString().Length != 0 && rows["Nombre"].ToString() != null &&
                                rows["EditarNombre"].ToString().Length != 0 && rows["EditarNombre"].ToString() != null &&

                                rows["TipoRepresentante"].ToString().Length != 0 && rows["TipoRepresentante"].ToString() != null &&
                                rows["Funciones"].ToString().Length != 0 && rows["Funciones"].ToString() != null &&
                                rows["EditarFunciones"].ToString().Length != 0 && rows["EditarFunciones"].ToString() != null &&
                                rows["tablaDetalle1"].ToString().Length != 0 && rows["tablaDetalle1"].ToString() != null &&
                                rows["campoDetalle1"].ToString().Length != 0 && rows["campoDetalle1"].ToString() != null && 
                                rows["editCampoD1"].ToString().Length != 0 && rows["editCampoD1"].ToString() != null &&

                                rows["CentroCostos"].ToString().Length != 0 && rows["CentroCostos"].ToString() != null &&
                                rows["Actividad"].ToString().Length != 0 && rows["Actividad"].ToString() != null &&
                                rows["EditarActividad"].ToString().Length != 0 && rows["EditarActividad"].ToString() != null &&
                                rows["tablaDetalle2"].ToString().Length != 0 && rows["tablaDetalle2"].ToString() != null &&
                                rows["campoDetalle2"].ToString().Length != 0 && rows["campoDetalle2"].ToString() != null &&
                                rows["editCampoD2"].ToString().Length != 0 && rows["editCampoD2"].ToString() != null 
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

                                string TipoComite = rows["TipoComite"].ToString();
                                string Nombre = rows["Nombre"].ToString();
                                string EditarNombre = rows["EditarNombre"].ToString();
                                string EditCampo = rows["EditCampo"].ToString();

                                string TipoRepresentante = rows["TipoRepresentante"].ToString();
                                string Funciones = rows["Funciones"].ToString();
                                string EditarFunciones = rows["EditarFunciones"].ToString();
                                string tablaDetalle1 = rows["tablaDetalle1"].ToString();
                                string campoDetalle1 = rows["campoDetalle1"].ToString();
                                string editCampoD1 = rows["editCampoD1"].ToString();

                                string CentroCostos = rows["CentroCostos"].ToString();
                                string Actividad = rows["Actividad"].ToString();
                                string EditarActividad = rows["EditarActividad"].ToString();
                                string tablaDetalle2 = rows["tablaDetalle2"].ToString();
                                string campoDetalle2 = rows["campoDetalle2"].ToString();
                                string editCampoD2 = rows["editCampoD2"].ToString();

                                List<string> crudVariables = new List<string>() { TipoComite, Nombre, EditarNombre };
                                List<string> crudDetalle1 = new List<string>() { TipoRepresentante, Funciones, EditarFunciones };
                                List<string> crudDetalle2 = new List<string>() { CentroCostos, Actividad, EditarActividad };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);
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

                                        //ELIMINAR REGISTRO PEGADO DETALLE 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, CentroCostos, CentroCostos, campoDetalle2, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //ELIMINAR REGISTRO PEGADO DETALLE 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, TipoRepresentante, TipoRepresentante, campoDetalle1, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //ELIMINAR REGISTRO PEGADO
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, TipoComite, TipoComite, Campo, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(2000);

                                        ////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 27, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKsotcomi.CRUD(desktopSession, 1, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(2000);
                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, TipoComite, TipoComite, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKsotcomi.CRUD(desktopSession, 2, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, TipoComite, Nombre, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKsotcomi.CRUD(desktopSession, 2, crudVariables, file);
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

                                            CrudKsotcomi.CRUD(desktopSession, 2, crudVariables, file);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, TipoComite, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, TipoComite, EditarNombre, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);

                                        //1 -> Tipos Representantes 
                                        CrudKsotcomi.ClickButtonExterno(desktopSession, 1);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 1
                                        Dictionary<string, Point> botonesNavegadorDetalle1 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle1 = PruebaCRUD.findname(desktopSession, 27, 1);
                                        var ElementList1 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Insertar"].X, botonesNavegadorDetalle1["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKsotcomi.CrudDetalle1(desktopSession, 1, crudDetalle1, file);
                                        CrudKsotcomi.AcomodarDetalle(desktopSession);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Aplicar"].X, botonesNavegadorDetalle1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 1", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, TipoRepresentante, TipoRepresentante, campoDetalle1, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Editar"].X, botonesNavegadorDetalle1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                        Thread.Sleep(1000);

                                        CrudKsotcomi.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
                                        CrudKsotcomi.AcomodarDetalle(desktopSession);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Cancelar"].X, botonesNavegadorDetalle1["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 1", true, file);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, TipoRepresentante, Funciones, editCampoD1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        Thread.Sleep(1000);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Editar"].X, botonesNavegadorDetalle1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                        Thread.Sleep(1000);

                                        CrudKsotcomi.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
                                        CrudKsotcomi.AcomodarDetalle(desktopSession);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Aplicar"].X, botonesNavegadorDetalle1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 1", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(3000);
                                        // //Debugger.Launch();
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, TipoRepresentante, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, TipoRepresentante, EditarFunciones, editCampoD1, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //2 -> Población para comité y votaciones 
                                        CrudKsotcomi.ClickButtonExterno(desktopSession, 2);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 2
                                        Dictionary<string, Point> botonesNavegadorDetalle2 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle2 = PruebaCRUD.findname(desktopSession, 27, 1);
                                        var ElementList2 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Insertar"].X, botonesNavegadorDetalle2["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKsotcomi.CrudDetalle2(desktopSession, 1, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 2", true, file);

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, CentroCostos, CentroCostos, campoDetalle2, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);

                                        Thread.Sleep(1000);

                                        CrudKsotcomi.CrudDetalle2(desktopSession, 2, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Cancelar"].X, botonesNavegadorDetalle2["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 2", true, file);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, CentroCostos, Actividad, editCampoD2, c, ErrorValidacion, "No se edito correctamente", 1);

                                        Thread.Sleep(1000);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);

                                        Thread.Sleep(1000);

                                        CrudKsotcomi.CrudDetalle2(desktopSession, 2, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 2", true, file);
                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, CentroCostos, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, CentroCostos, EditarActividad, editCampoD2, c, ErrorValidacion, "El registro no se editó correctamente", 1);

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
                                        /*string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, TipoComite, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));*/

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //2 -> Población para comité y votaciones 
                                        CrudKsotcomi.ClickButtonExterno(desktopSession, 2);

                                        //ELIMINAR DETALLE 2
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList2[1], botonesNavegadorDetalle2, file);
                                        Thread.Sleep(500);

                                        ////Validacion Eliminar Detalle 2
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, CentroCostos, "D", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, CentroCostos, "", campoDetalle2, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //1 -> Tipos Representantes 
                                        CrudKsotcomi.ClickButtonExterno(desktopSession, 1);

                                        //ELIMINAR DETALLE 1
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[1], botonesNavegadorDetalle1, file);
                                        Thread.Sleep(500);

                                        ////Validacion Eliminar Detalle 1
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, TipoRepresentante, "D", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, TipoRepresentante, "", campoDetalle1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);

                                        ////Validacion Eliminar Registro
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, TipoComite, "D", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, TipoComite, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);
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
        public void KSoTipvaCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosSO.ModulosSO_12.KSoTipvaCheckList")
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

                                rows["Codigo"].ToString().Length != 0 && rows["Codigo"].ToString() != null &&
                                rows["Vacuna"].ToString().Length != 0 && rows["Vacuna"].ToString() != null &&
                                rows["EditarVacuna"].ToString().Length != 0 && rows["EditarVacuna"].ToString() != null && 
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&

                                rows["Dosis"].ToString().Length != 0 && rows["Dosis"].ToString() != null &&
                                rows["Fecha"].ToString().Length != 0 && rows["Fecha"].ToString() != null &&
                                rows["NumTiempo"].ToString().Length != 0 && rows["NumTiempo"].ToString() != null &&
                                rows["Tiempo"].ToString().Length != 0 && rows["Tiempo"].ToString() != null &&
                                rows["EditarTiempo"].ToString().Length != 0 && rows["EditarTiempo"].ToString() != null &&
                                rows["tablaDetalle1"].ToString().Length != 0 && rows["tablaDetalle1"].ToString() != null &&
                                rows["campoDetalle1"].ToString().Length != 0 && rows["campoDetalle1"].ToString() != null &&
                                rows["editCampoD1"].ToString().Length != 0 && rows["editCampoD1"].ToString() != null &&

                                rows["CodigoD2"].ToString().Length != 0 && rows["CodigoD2"].ToString() != null &&
                                rows["EditarCodigo"].ToString().Length != 0 && rows["EditarCodigo"].ToString() != null &&
                                rows["tablaDetalle2"].ToString().Length != 0 && rows["tablaDetalle2"].ToString() != null &&
                                rows["campoDetalle2"].ToString().Length != 0 && rows["campoDetalle2"].ToString() != null &&
                                rows["editCampoD2"].ToString().Length != 0 && rows["editCampoD2"].ToString() != null

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

                                string Codigo = rows["Codigo"].ToString();
                                string Vacuna = rows["Vacuna"].ToString();
                                string EditarVacuna = rows["EditarVacuna"].ToString(); 
                                string EditCampo = rows["EditCampo"].ToString();

                                string Dosis = rows["Dosis"].ToString();
                                string Fecha = rows["Fecha"].ToString();
                                string NumTiempo = rows["NumTiempo"].ToString();
                                string Tiempo = rows["Tiempo"].ToString();
                                string EditarTiempo = rows["EditarTiempo"].ToString();
                                string tablaDetalle1 = rows["tablaDetalle1"].ToString();
                                string campoDetalle1 = rows["campoDetalle1"].ToString();
                                string editCampoD1 = rows["editCampoD1"].ToString();

                                string CodigoD2 = rows["CodigoD2"].ToString();
                                string EditarCodigo = rows["EditarCodigo"].ToString();
                                string tablaDetalle2 = rows["tablaDetalle2"].ToString();
                                string campoDetalle2 = rows["campoDetalle2"].ToString();
                                string editCampoD2 = rows["editCampoD2"].ToString();

                                List<string> crudVariables = new List<string>() { Codigo, Vacuna, EditarVacuna };
                                List<string> crudDetalle1 = new List<string>() { Dosis, Fecha, NumTiempo, Tiempo, EditarTiempo };
                                List<string> crudDetalle2 = new List<string>() { CodigoD2, EditarCodigo};

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "SO", motor);
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

                                        //ELIMINAR REGISTRO PEGADO DETALLE 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, Codigo, Codigo, campoDetalle2, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //ELIMINAR REGISTRO PEGADO DETALLE 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, Dosis, Dosis, campoDetalle1, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //ELIMINAR REGISTRO PEGADO
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, Campo, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

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

                                        CrudKsotipva.CRUD(desktopSession, 1, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKsotipva.CRUD(desktopSession, 2, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Vacuna, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKsotipva.CRUD(desktopSession, 2, crudVariables, file);
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

                                            CrudKsotipva.CRUD(desktopSession, 2, crudVariables, file);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, EditarVacuna, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);

                                        //1 -> Dosis
                                        CrudKsotipva.ClickButtonExterno(desktopSession, 1);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 1
                                        Dictionary<string, Point> botonesNavegadorDetalle1 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle1 = PruebaCRUD.findname(desktopSession, 25, 1);
                                        var ElementList1 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Insertar"].X, botonesNavegadorDetalle1["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKsotipva.CrudDetalle1(desktopSession, 1, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Aplicar"].X, botonesNavegadorDetalle1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 1", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, Dosis, Dosis, campoDetalle1, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Editar"].X, botonesNavegadorDetalle1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                        Thread.Sleep(1000);

                                        CrudKsotipva.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Cancelar"].X, botonesNavegadorDetalle1["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 1", true, file);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, Dosis, Tiempo, editCampoD1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        Thread.Sleep(1000);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Editar"].X, botonesNavegadorDetalle1["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                        Thread.Sleep(1000);

                                        CrudKsotipva.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle1["Aplicar"].X, botonesNavegadorDetalle1["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 1", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(3000);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, Dosis, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, Dosis, EditarTiempo, editCampoD1, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //2 -> Cargos
                                        CrudKsotipva.ClickButtonExterno(desktopSession, 2);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 2
                                        Dictionary<string, Point> botonesNavegadorDetalle2 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle2 = PruebaCRUD.findname(desktopSession, 26, 1);
                                        var ElementList2 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Insertar"].X, botonesNavegadorDetalle2["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKsotipva.CrudDetalle2(desktopSession, 1, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 2", true, file);

                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, Codigo, Codigo, campoDetalle2, c, ErrorValidacion, "No se agregó correctamente", 0);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);

                                        Thread.Sleep(1000);

                                        CrudKsotipva.CrudDetalle2(desktopSession, 2, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Cancelar"].X, botonesNavegadorDetalle2["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 2", true, file);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, Codigo, CodigoD2, editCampoD2, c, ErrorValidacion, "No se edito correctamente", 1);

                                        Thread.Sleep(1000);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);

                                        Thread.Sleep(1000);

                                        CrudKsotipva.CrudDetalle2(desktopSession, 2, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 2", true, file);
                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, Codigo, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, Codigo, EditarCodigo, editCampoD2, c, ErrorValidacion, "El registro no se editó correctamente", 1);

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
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, Codigo, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //2 -> Cargos
                                        CrudKsotipva.ClickButtonExterno(desktopSession, 2);

                                        //ELIMINAR DETALLE 2
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList2[1], botonesNavegadorDetalle2, file);
                                        Thread.Sleep(500);

                                        ////Validacion Eliminar Detalle 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, Codigo, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, Codigo, "", campoDetalle2, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //1 -> Dosis
                                        CrudKsotipva.ClickButtonExterno(desktopSession, 1);

                                        //ELIMINAR DETALLE 1
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[1], botonesNavegadorDetalle1, file);
                                        Thread.Sleep(500);

                                        ////Validacion Eliminar Detalle 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, Dosis, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, Dosis, "", campoDetalle1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);

                                        ////Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);
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

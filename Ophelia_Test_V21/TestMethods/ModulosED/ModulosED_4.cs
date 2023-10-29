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
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;

namespace Ophelia_Test_V21.TestMethods.ModulosED
{
    /// <summary>
    /// Descripción resumida de ModulosED_4
    /// </summary>
    [TestClass]
    public class ModulosED_4 : FuncionesVitales
    {

        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;

        public ModulosED_4() { }

        [TestMethod]
        public void KEdFomeiCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosED.ModulosED_4.KEdFomeiCheckList")
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
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null

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

                                //List<string> crudVariables = new List<string>() { Observaciones, Fecha };


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "ED", motor);
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
                                        //errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////VALIDACION NOMBRE DEL PROGRAMA
                                        //errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //CrudKnmaumge.Ventana(desktopSession);
                                        //newFunctions_4.ScreenshotNuevo("Ventana Controlada", true, file);
                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

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
                                        //CrudKnmctpre.Preliminar(desktopSession, file, pdf1, BanderaPreli);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        /*
                                        /if (motor == "SQL")
                                        {
                                            //RECTIFICACION DE CAMPOS DE EXCEL  
                                            string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, QbeFilter, c);
                                            string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                            Thread.Sleep(3000);
                                            Celda = Celda_Excel(ruta, nom_empr);
                                            Celda.ForEach(u => ErrorValidacion.Add(u));
                                        }*/
                                       
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
        public void KEd3TevaCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosED.ModulosED_4.KEd3TevaCheckList")
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

                                //@NombreServicio @URLServ @Interface @EditarURLServ                         
                                rows["TipoEscala"].ToString().Length != 0 && rows["TipoEscala"].ToString() != null &&
                                rows["ProtEvaluacionNum"].ToString().Length != 0 && rows["ProtEvaluacionNum"].ToString() != null &&
                                rows["ProtEvaluacionPal"].ToString().Length != 0 && rows["ProtEvaluacionPal"].ToString() != null &&
                                rows["PorEvaluacion"].ToString().Length != 0 && rows["PorEvaluacion"].ToString() != null &&
                                rows["EditPorEvaluacion"].ToString().Length != 0 && rows["EditPorEvaluacion"].ToString() != null &&

                                rows["Tipo"].ToString().Length != 0 && rows["Tipo"].ToString() != null &&
                                rows["PorPeso"].ToString().Length != 0 && rows["PorPeso"].ToString() != null &&
                                rows["EditPorPeso"].ToString().Length != 0 && rows["EditPorPeso"].ToString() != null &&

                                rows["ValInicial"].ToString().Length != 0 && rows["ValInicial"].ToString() != null &&
                                rows["ValFinal"].ToString().Length != 0 && rows["ValFinal"].ToString() != null &&
                                rows["Descripcion"].ToString().Length != 0 && rows["Descripcion"].ToString() != null &&
                                rows["EditDescripcion"].ToString().Length != 0 && rows["EditDescripcion"].ToString() != null &&

                                rows["NombreVJ"].ToString().Length != 0 && rows["NombreVJ"].ToString() != null &&
                                rows["DescripcionVJ"].ToString().Length != 0 && rows["DescripcionVJ"].ToString() != null &&
                                rows["EditDescripcionVJ"].ToString().Length != 0 && rows["EditDescripcionVJ"].ToString() != null &&

                                //Validacion
                                rows["tablaDetalle1"].ToString().Length != 0 && rows["tablaDetalle1"].ToString() != null &&
                                rows["campoDetalle1"].ToString().Length != 0 && rows["campoDetalle1"].ToString() != null &&
                                rows["editCampoD1"].ToString().Length != 0 && rows["editCampoD1"].ToString() != null &&
                                rows["tablaDetalle2"].ToString().Length != 0 && rows["tablaDetalle2"].ToString() != null &&
                                rows["campoDetalle2"].ToString().Length != 0 && rows["campoDetalle2"].ToString() != null &&
                                rows["editCampoD2"].ToString().Length != 0 && rows["editCampoD2"].ToString() != null &&
                                rows["tablaDetalle3"].ToString().Length != 0 && rows["tablaDetalle3"].ToString() != null &&
                                rows["campoDetalle3"].ToString().Length != 0 && rows["campoDetalle3"].ToString() != null &&
                                rows["editCampoD3"].ToString().Length != 0 && rows["editCampoD3"].ToString() != null
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

                                string TipoEscala = rows["TipoEscala"].ToString();
                                string ProtEvaluacionNum = rows["ProtEvaluacionNum"].ToString();
                                string ProtEvaluacionPal = rows["ProtEvaluacionPal"].ToString();
                                string PorEvaluacion = rows["PorEvaluacion"].ToString();
                                string EditPorEvaluacion = rows["EditPorEvaluacion"].ToString();

                                string Tipo = rows["Tipo"].ToString();
                                string PorPeso = rows["PorPeso"].ToString();
                                string EditPorPeso = rows["EditPorPeso"].ToString();

                                string ValInicial = rows["ValInicial"].ToString();
                                string ValFinal = rows["ValFinal"].ToString();
                                string Descripcion = rows["Descripcion"].ToString();
                                string EditDescripcion = rows["EditDescripcion"].ToString();

                                string NombreVJ = rows["NombreVJ"].ToString();
                                string DescripcionVJ = rows["DescripcionVJ"].ToString();
                                string EditDescripcionVJ = rows["EditDescripcionVJ"].ToString();

                                string tablaDetalle1 = rows["tablaDetalle1"].ToString();
                                string campoDetalle1 = rows["campoDetalle1"].ToString();
                                string editCampoD1 = rows["editCampoD1"].ToString();
                                string tablaDetalle2 = rows["tablaDetalle2"].ToString();
                                string campoDetalle2 = rows["campoDetalle2"].ToString();
                                string editCampoD2 = rows["editCampoD2"].ToString();
                                string tablaDetalle3 = rows["tablaDetalle3"].ToString();
                                string campoDetalle3 = rows["campoDetalle3"].ToString();
                                string editCampoD3 = rows["editCampoD3"].ToString();

                                List<string> crudVariables = new List<string>() { TipoEscala, ProtEvaluacionNum, ProtEvaluacionPal, PorEvaluacion, EditPorEvaluacion };
                                List<string> crudDetalle1 = new List<string>() { Tipo, PorPeso, EditPorPeso };
                                List<string> crudDetalle2 = new List<string>() { ValInicial, ValFinal, Descripcion, EditDescripcion };
                                List<string> crudDetalle3 = new List<string>() { NombreVJ, DescripcionVJ, EditDescripcionVJ };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "ED", motor);
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

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 3
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, NombreVJ, NombreVJ, campoDetalle3, c, ErrorValidacion, "No se agregó correctamente 3", 0,"1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, ValInicial, ValInicial, campoDetalle2, c, ErrorValidacion, "No se agregó correctamente 2", 0,"1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, Tipo, Tipo, campoDetalle1, c, ErrorValidacion, "No se agregó correctamente 1", 0,"1");

                                        //ELIMINAR REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, TipoEscala, TipoEscala, Campo, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

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

                                        CrudKed3teva.CRUD(desktopSession, 1, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, TipoEscala, TipoEscala, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);
                                        
                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKed3teva.CRUD(desktopSession, 2, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, TipoEscala, PorEvaluacion, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKed3teva.CRUD(desktopSession, 2, crudVariables, file);
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

                                            CrudKed3teva.CRUD(desktopSession, 2, crudVariables, file);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, TipoEscala, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, TipoEscala, EditPorEvaluacion, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 1
                                        Dictionary<string, Point> botonesNavegadorDetalle = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle = PruebaCRUD.findname(desktopSession, 27, 1);
                                        var ElementList1 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Insertar"].X, botonesNavegadorDetalle["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKed3teva.CrudDetalle1(desktopSession, 1, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 1", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, Tipo, Tipo, campoDetalle1, c, ErrorValidacion, "No se agregó correctamente 1", 0);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);
                                        Thread.Sleep(1000);

                                        CrudKed3teva.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Cancelar"].X, botonesNavegadorDetalle["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Edición Detalle 1", true, file);
                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, Tipo, PorPeso, editCampoD1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);
                                        Thread.Sleep(1000);

                                        CrudKed3teva.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Aplicados Detalle 1", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(3000);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, Tipo, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, Tipo, EditPorPeso, editCampoD1, c, ErrorValidacion, "El registro no se editó correctamente", 1);
                                        
                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 2
                                        Dictionary<string, Point> botonesNavegadorDetalle2 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle2 = PruebaCRUD.findname(desktopSession, 27, 2);
                                        var ElementList2 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList2[2].Coordinates, botonesNavegadorDetalle2["Insertar"].X, botonesNavegadorDetalle2["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKed3teva.CrudDetalle2(desktopSession, 1, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 2", true, file);

                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList2[2].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 2", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, ValInicial, ValInicial, campoDetalle2, c, ErrorValidacion, "No se agregó correctamente 2", 0);

                                        desktopSession.Mouse.MouseMove(ElementList2[2].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);

                                        Thread.Sleep(1000);

                                        CrudKed3teva.CrudDetalle2(desktopSession, 2, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[2].Coordinates, botonesNavegadorDetalle2["Cancelar"].X, botonesNavegadorDetalle2["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Edición Detalle 2", true, file);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, ValInicial, Descripcion, editCampoD2, c, ErrorValidacion, "No se edito correctamente", 1);

                                        Thread.Sleep(1000);
                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[2].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);

                                        Thread.Sleep(1000);

                                        CrudKed3teva.CrudDetalle2(desktopSession, 2, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[2].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Aplicados Detalle 2", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, ValInicial, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, ValInicial, EditDescripcion, editCampoD2, c, ErrorValidacion, "El registro no se editó correctamente", 1);
                                       
                                        // 1 - Detalles  2 - Ventana Johari
                                        CrudKed3teva.ClickButton(desktopSession, 2);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 3
                                        Dictionary<string, Point> botonesNavegadorDetalle3 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle3 = PruebaCRUD.findname(desktopSession, 27, 1);
                                        var ElementList3 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorDetalle3["Insertar"].X, botonesNavegadorDetalle3["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKed3teva.CrudDetalle3(desktopSession, 1, crudDetalle3, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorDetalle3["Aplicar"].X, botonesNavegadorDetalle3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 3", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, NombreVJ, NombreVJ, campoDetalle3, c, ErrorValidacion, "No se agregó correctamente 3", 0);

                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorDetalle3["Editar"].X, botonesNavegadorDetalle3["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 3", true, file);
                                        Thread.Sleep(1000);

                                        CrudKed3teva.CrudDetalle3(desktopSession, 2, crudDetalle3, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorDetalle3["Cancelar"].X, botonesNavegadorDetalle3["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Edición Detalle 3", true, file);
                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, NombreVJ, DescripcionVJ, editCampoD3, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorDetalle3["Editar"].X, botonesNavegadorDetalle3["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 3", true, file);
                                        Thread.Sleep(1000);

                                        CrudKed3teva.CrudDetalle3(desktopSession, 2, crudDetalle3, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorDetalle3["Aplicar"].X, botonesNavegadorDetalle3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Aplicados Detalle 3", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle3, motor, user, campoDetalle3, NombreVJ, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, NombreVJ, EditDescripcionVJ, editCampoD3, c, ErrorValidacion, "El registro no se editó correctamente", 1);

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
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, TipoEscala, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        /*Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));*/

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        // 1 - Detalles  2 - Ventana Johari
                                        CrudKed3teva.ClickButton(desktopSession, 2);

                                        //ELIMINAR DETALLE 3
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList3[1], botonesNavegadorDetalle3, file);
                                        Thread.Sleep(500);

                                        ////Validacion Eliminar Detalle 3
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle3, motor, user, campoDetalle3, NombreVJ, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, NombreVJ, "", campoDetalle3, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        // 1 - Detalles  2 - Ventana Johari
                                        CrudKed3teva.ClickButton(desktopSession, 1);

                                        //ELIMINAR DETALLE 2
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList2[2], botonesNavegadorDetalle2, file);
                                        Thread.Sleep(500);

                                        ////Validacion Eliminar Detalle 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, ValInicial, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, ValInicial, "", campoDetalle2, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //ELIMINAR DETALLE 1
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[1], botonesNavegadorDetalle, file);
                                        Thread.Sleep(500);

                                        ////Validacion Eliminar Detalle 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, Tipo, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, Tipo, "", campoDetalle1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        CrudKed3teva.VentanaEliminar(desktopSession);
                                        Thread.Sleep(1000);

                                        ////Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, TipoEscala, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, TipoEscala, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);
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
        public void KEd3CrolCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosED.ModulosED_4.KEd3CrolCheckList")
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

                                rows["RolNum"].ToString().Length != 0 && rows["RolNum"].ToString() != null &&
                                rows["RolPal"].ToString().Length != 0 && rows["RolPal"].ToString() != null &&
                                rows["EditarRol"].ToString().Length != 0 && rows["EditarRol"].ToString() != null &&
                                rows["Cargo"].ToString().Length != 0 && rows["Cargo"].ToString() != null &&
                                rows["EditarCargo"].ToString().Length != 0 && rows["EditarCargo"].ToString() != null &&
                                rows["CargoSiguiente"].ToString().Length != 0 && rows["CargoSiguiente"].ToString() != null &&
                                rows["EditarCargoSig"].ToString().Length != 0 && rows["EditarCargoSig"].ToString() != null &&

                                rows["CentroCostos"].ToString().Length != 0 && rows["CentroCostos"].ToString() != null &&
                                rows["EditCentroCostos"].ToString().Length != 0 && rows["EditCentroCostos"].ToString() != null &&
                                rows["CargoRoles"].ToString().Length != 0 && rows["CargoRoles"].ToString() != null &&
                                rows["EditCargoRoles"].ToString().Length != 0 && rows["EditCargoRoles"].ToString() != null &&

                                //Validacion
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
                                rows["editCampoD4"].ToString().Length != 0 && rows["editCampoD4"].ToString() != null 
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

                                string RolNum = rows["RolNum"].ToString();
                                string RolPal = rows["RolPal"].ToString();
                                string EditarRol = rows["EditarRol"].ToString();
                                string Cargo = rows["Cargo"].ToString();
                                string EditarCargo = rows["EditarCargo"].ToString();
                                string CargoSiguiente = rows["CargoSiguiente"].ToString();
                                string EditarCargoSig = rows["EditarCargoSig"].ToString();

                                string CentroCostos = rows["CentroCostos"].ToString();
                                string EditCentroCostos = rows["EditCentroCostos"].ToString();
                                string CargoRoles = rows["CargoRoles"].ToString();
                                string EditCargoRoles = rows["EditCargoRoles"].ToString();

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

                                List<string> crudVariables = new List<string>() { RolNum, RolPal, EditarRol };
                                List<string> crudDetalle1 = new List<string>() { Cargo, EditarCargo };
                                List<string> crudDetalle2 = new List<string>() { CargoSiguiente, EditarCargoSig };
                                List<string> crudDetalle3 = new List<string>() { CentroCostos, EditCentroCostos };
                                List<string> crudDetalle4 = new List<string>() { CargoRoles, EditCargoRoles };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "ED", motor);
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

                                        //ELIMINAR REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, RolNum, RolNum, Campo, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

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

                                        CrudKed3crol.CRUD(desktopSession, 1, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);

                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, RolNum, RolNum, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);
                                        
                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKed3crol.CRUD(desktopSession, 2, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, RolNum, RolPal, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKed3crol.CRUD(desktopSession, 2, crudVariables, file);
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

                                            CrudKed3crol.CRUD(desktopSession, 2, crudVariables, file);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, RolNum, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, RolNum, EditarRol, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 1
                                        Dictionary<string, Point> botonesNavegadorDetalle = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle = PruebaCRUD.findname(desktopSession, 27, 2);
                                        var ElementList1 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList1[2].Coordinates, botonesNavegadorDetalle["Insertar"].X, botonesNavegadorDetalle["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKed3crol.CrudDetalle1(desktopSession, 1, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList1[2].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 1", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, RolNum, RolNum, campoDetalle1, c, ErrorValidacion, "No se agregó correctamente 1", 0);

                                        desktopSession.Mouse.MouseMove(ElementList1[2].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);
                                        Thread.Sleep(1000);

                                        CrudKed3crol.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList1[2].Coordinates, botonesNavegadorDetalle["Cancelar"].X, botonesNavegadorDetalle["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Edición Detalle 1", true, file);
                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, RolNum, Cargo, editCampoD1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList1[2].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);
                                        Thread.Sleep(1000);

                                        CrudKed3crol.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList1[2].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Aplicados Detalle 1", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(3000);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, RolNum, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, RolNum, EditarCargo, editCampoD1, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 2
                                        Dictionary<string, Point> botonesNavegadorDetalle2 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle2 = PruebaCRUD.findname(desktopSession, 27, 1);
                                        var ElementList2 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Insertar"].X, botonesNavegadorDetalle2["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKed3crol.CrudDetalle2(desktopSession, 1, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 2", true, file);

                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 2", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, RolNum, RolNum, campoDetalle2, c, ErrorValidacion, "No se agregó correctamente 2", 0);

                                        
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);
    
                                        Thread.Sleep(1000);

                                        CrudKed3crol.CrudDetalle2(desktopSession, 2, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Cancelar"].X, botonesNavegadorDetalle2["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Edición Detalle 2", true, file);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, RolNum, CargoSiguiente, editCampoD2, c, ErrorValidacion, "No se edito correctamente", 1);

                                        Thread.Sleep(1000);
                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);

                                        Thread.Sleep(1000);

                                        CrudKed3crol.CrudDetalle2(desktopSession, 2, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Aplicados Detalle 2", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, RolNum, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, RolNum, EditarCargoSig, editCampoD2, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        // 2 - Centros de Costo 
                                        CrudKed3crol.CLickButton(desktopSession, 2);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 3
                                        Dictionary<string, Point> botonesNavegadorDetalle3 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle3 = PruebaCRUD.findname(desktopSession, 27, 1);
                                        var ElementList3 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorDetalle3["Insertar"].X, botonesNavegadorDetalle3["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKed3crol.CrudDetalle3(desktopSession, 1, crudDetalle3, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorDetalle3["Aplicar"].X, botonesNavegadorDetalle3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 3", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, RolNum, RolNum, campoDetalle3, c, ErrorValidacion, "No se agregó correctamente 3", 0);

                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorDetalle3["Editar"].X, botonesNavegadorDetalle3["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 3", true, file);
                                        Thread.Sleep(1000);

                                        CrudKed3crol.CrudDetalle3(desktopSession, 2, crudDetalle3, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorDetalle3["Cancelar"].X, botonesNavegadorDetalle3["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Edición Detalle 3", true, file);
                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, RolNum, CentroCostos, editCampoD3, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorDetalle3["Editar"].X, botonesNavegadorDetalle3["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 3", true, file);
                                        Thread.Sleep(1000);

                                        CrudKed3crol.CrudDetalle3(desktopSession, 2, crudDetalle3, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegadorDetalle3["Aplicar"].X, botonesNavegadorDetalle3["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Aplicados Detalle 3", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle3, motor, user, campoDetalle3, RolNum, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, RolNum, EditCentroCostos, editCampoD3, c, ErrorValidacion, "El registro no se editó correctamente", 1);


                                        // 3 - Roles de Planta
                                        CrudKed3crol.CLickButton(desktopSession, 3);

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 4
                                        Dictionary<string, Point> botonesNavegadorDetalle4 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle4 = PruebaCRUD.findname(desktopSession, 27, 1);
                                        var ElementList4 = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Insertar"].X, botonesNavegadorDetalle4["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKed3crol.CrudDetalle4(desktopSession, 1, crudDetalle4, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 4", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Aplicar"].X, botonesNavegadorDetalle4["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 4", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle4, user, motor, campoDetalle4, RolNum, RolNum, campoDetalle4, c, ErrorValidacion, "No se agregó correctamente 3", 0);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Editar"].X, botonesNavegadorDetalle4["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 4", true, file);
                                        Thread.Sleep(1000);

                                        CrudKed3crol.CrudDetalle4(desktopSession, 2, crudDetalle4, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 4", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Cancelar"].X, botonesNavegadorDetalle4["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Edición Detalle 4", true, file);
                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle4, user, motor, campoDetalle4, RolNum, CargoRoles, editCampoD4, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Editar"].X, botonesNavegadorDetalle4["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 4", true, file);
                                        Thread.Sleep(1000);

                                        CrudKed3crol.CrudDetalle4(desktopSession, 2, crudDetalle4, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 4", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList4[1].Coordinates, botonesNavegadorDetalle4["Aplicar"].X, botonesNavegadorDetalle4["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Aplicados Detalle 4", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle3, motor, user, campoDetalle4, RolNum, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle4, user, motor, campoDetalle4, RolNum, EditCargoRoles, editCampoD4, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

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
                                        //CrudKnmctpre.Preliminar(desktopSession, file, pdf1, BanderaPreli);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        // LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //RECTIFICACION DE CAMPOS DE EXCEL  
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, RolNum, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        // 3 - Roles de Planta
                                        CrudKed3crol.CLickButton(desktopSession, 3);

                                        //ELIMINAR DETALLE 4
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList4[1], botonesNavegadorDetalle4, file);
                                        Thread.Sleep(500);

                                        ////Validacion Eliminar Detalle 4
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle4, motor, user, campoDetalle4, RolNum, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle4, user, motor, campoDetalle4, RolNum, "", campoDetalle4, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        // 2 - Centros de Costo 
                                        CrudKed3crol.CLickButton(desktopSession, 2);

                                        //ELIMINAR DETALLE 3
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList3[1], botonesNavegadorDetalle3, file);
                                        Thread.Sleep(500);

                                        ////Validacion Eliminar Detalle 3
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle3, motor, user, campoDetalle3, RolNum, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, RolNum, "", campoDetalle3, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        // 1 - Cargo 
                                        CrudKed3crol.CLickButton(desktopSession, 1);

                                        //ELIMINAR DETALLE 2
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorDetalle2, file);
                                        Thread.Sleep(500);

                                        ////Validacion Eliminar Detalle 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, RolNum, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, RolNum, "", campoDetalle2, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR DETALLE 1
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[2], botonesNavegadorDetalle, file);
                                        Thread.Sleep(500);

                                        ////Validacion Eliminar Detalle 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, RolNum, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, RolNum, "", campoDetalle1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        ////Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, RolNum, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, RolNum, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);
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
        public void KEd3PerfCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosED.ModulosED_4.KEd3PerfCheckList")
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

                                rows["Rol"].ToString().Length != 0 && rows["Rol"].ToString() != null &&
                                rows["DescripcionEval"].ToString().Length != 0 && rows["DescripcionEval"].ToString() != null &&
                                rows["EditDescripcionEval"].ToString().Length != 0 && rows["EditDescripcionEval"].ToString() != null &&
                                rows["Prototipo"].ToString().Length != 0 && rows["Prototipo"].ToString() != null &&
                                rows["Porcentaje"].ToString().Length != 0 && rows["Porcentaje"].ToString() != null &&
                                rows["EditPorcentaje"].ToString().Length != 0 && rows["EditPorcentaje"].ToString() != null &&
                                rows["Item"].ToString().Length != 0 && rows["Item"].ToString() != null &&
                                rows["EditarItem"].ToString().Length != 0 && rows["EditarItem"].ToString() != null &&

                                rows["tablaDetalle1"].ToString().Length != 0 && rows["tablaDetalle1"].ToString() != null &&
                                rows["campoDetalle1"].ToString().Length != 0 && rows["campoDetalle1"].ToString() != null &&
                                rows["editCampoD1"].ToString().Length != 0 && rows["editCampoD1"].ToString() != null &&
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
                                string EditCampo = rows["EditCampo"].ToString();

                                string Rol = rows["Rol"].ToString();
                                string DescripcionEval = rows["DescripcionEval"].ToString();
                                string EditDescripcionEval = rows["EditDescripcionEval"].ToString();
                                string Prototipo = rows["Prototipo"].ToString();
                                string Porcentaje = rows["Porcentaje"].ToString();
                                string EditPorcentaje = rows["EditPorcentaje"].ToString();
                                string Item = rows["Item"].ToString();
                                string EditarItem = rows["EditarItem"].ToString();

                                string tablaDetalle1 = rows["tablaDetalle1"].ToString();
                                string campoDetalle1 = rows["campoDetalle1"].ToString();
                                string editCampoD1 = rows["editCampoD1"].ToString();
                                string tablaDetalle2 = rows["tablaDetalle2"].ToString();
                                string campoDetalle2 = rows["campoDetalle2"].ToString();
                                string editCampoD2 = rows["editCampoD2"].ToString();

                                List<string> crudVariables = new List<string>() { Rol, DescripcionEval, EditDescripcionEval };
                                List<string> crudDetalle1 = new List<string>() { Prototipo, Porcentaje, EditPorcentaje };
                                List<string> crudDetalle2 = new List<string>() { Item, EditarItem };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "ED", motor);
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

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, Rol, Rol, campoDetalle2, c, ErrorValidacion, "No se agregó correctamente 2", 0,"1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, Prototipo, Prototipo, campoDetalle1, c, ErrorValidacion, "No se agregó correctamente 1", 0,"1");

                                        //ELIMINAR REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Rol, Rol, Campo, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

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

                                        CrudKed3perf.CRUD(desktopSession, 1, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);

                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Rol, Rol, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);
                                        
                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKed3perf.CRUD(desktopSession, 2, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);
                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Rol, DescripcionEval, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKed3perf.CRUD(desktopSession, 2, crudVariables, file);
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

                                            CrudKed3perf.CRUD(desktopSession, 2, crudVariables, file);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(3000);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);
                                        
                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Rol, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Rol, EditDescripcionEval, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 1
                                        Dictionary<string, Point> botonesNavegadorDetalle = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle = PruebaCRUD.findname(desktopSession, 27, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Insertar"].X, botonesNavegadorDetalle["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKed3perf.CrudDetalle1(desktopSession, 1, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 1", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, Prototipo, Prototipo, campoDetalle1, c, ErrorValidacion, "No se agregó correctamente 1", 0);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);
                                        Thread.Sleep(1000);
                                        
                                        CrudKed3perf.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Cancelar"].X, botonesNavegadorDetalle["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Cancelados Detalle 1", true, file);


                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, Prototipo, Porcentaje, editCampoD1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        Thread.Sleep(1000);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Editar"].X, botonesNavegadorDetalle["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                        Thread.Sleep(1000);

                                        CrudKed3perf.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegadorDetalle["Aplicar"].X, botonesNavegadorDetalle["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Aplicados Detalle 1", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(3000);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, Prototipo, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, Prototipo, EditPorcentaje, editCampoD1, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        ///LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        Thread.Sleep(1000);
                                        ////AGREGAR DETALLE 2
                                        Dictionary<string, Point> botonesNavegadorDetalle2 = new Dictionary<string, Point>();
                                        botonesNavegadorDetalle2 = PruebaCRUD.findname(desktopSession, 27, 2);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDetalle2["Insertar"].X, botonesNavegadorDetalle2["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKed3perf.CrudDetalle2(desktopSession, 1, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 2", true, file);

                                        Thread.Sleep(1000);
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 2", true, file); ;

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, Rol, Rol, campoDetalle2, c, ErrorValidacion, "No se agregó correctamente 2", 0);

                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);
                                        Thread.Sleep(1000);

                                        CrudKed3perf.CrudDetalle2(desktopSession, 2, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDetalle2["Cancelar"].X, botonesNavegadorDetalle2["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Cancelados Detalle 2", true, file);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, Rol, Item, editCampoD2, c, ErrorValidacion, "No se edito correctamente", 1);

                                        Thread.Sleep(1000);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDetalle2["Editar"].X, botonesNavegadorDetalle2["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);

                                        Thread.Sleep(1000);

                                        CrudKed3perf.CrudDetalle2(desktopSession, 2, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegadorDetalle2["Aplicar"].X, botonesNavegadorDetalle2["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Aplicados Detalle 2", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, Rol, "E", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, Rol, EditarItem, editCampoD2, c, ErrorValidacion, "El registro no se editó correctamente", 1);

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
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, Rol, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //ELIMINAR DETALLE 2
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[2], botonesNavegadorDetalle2, file);
                                        Thread.Sleep(500);

                                        ////Validacion Eliminar Detalle 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, Prototipo, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, Prototipo, "", campoDetalle2, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //ELIMINAR DETALLE 1
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegadorDetalle, file);
                                        Thread.Sleep(500);

                                        ////Validacion Eliminar Detalle 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, Rol, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, Rol, "", campoDetalle1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);

                                        ////Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Rol, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Rol, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);
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

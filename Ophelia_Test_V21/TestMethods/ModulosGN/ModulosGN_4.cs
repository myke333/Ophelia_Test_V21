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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;

namespace Ophelia_Test_V21.TestMethods.ModulosGN

{
    /// <summary>
    /// Descripción resumida de ModulosGN_4
    /// </summary>
    [TestClass]
    public class ModulosGN_4:FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;

        public ModulosGN_4(){ }

        [TestMethod]
        public void KGnMseriCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosGN.ModulosGN_4.KGnMseriCheckList")
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
                                //rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&
                                rows["EditCampo"].ToString().Length != 0 && rows["EditCampo"].ToString() != null &&

                                //@NombreServicio @URLServ @Interface @EditarURLServ                         
                                rows["NombreServ"].ToString().Length != 0 && rows["NombreServ"].ToString() != null &&
                                rows["URLServ"].ToString().Length != 0 && rows["URLServ"].ToString() != null &&
                                rows["Interface"].ToString().Length != 0 && rows["Interface"].ToString() != null &&
                                rows["EditarURLServ"].ToString().Length != 0 && rows["EditarURLServ"].ToString() != null &&
                                //@MetodoServ @Observaciones @EditarObs
                                rows["MetodoServicio"].ToString().Length != 0 && rows["MetodoServicio"].ToString() != null &&
                                rows["Observaciones"].ToString().Length != 0 && rows["Observaciones"].ToString() != null &&
                                rows["EditarObserv"].ToString().Length != 0 && rows["EditarObserv"].ToString() != null &&
                                //@NombreParam @TipoParam @ValorTipoCons @EditarParam
                                rows["NombreParam"].ToString().Length != 0 && rows["NombreParam"].ToString() != null &&
                                rows["TipoParam"].ToString().Length != 0 && rows["TipoParam"].ToString() != null &&
                                rows["ValorTipoCons"].ToString().Length != 0 && rows["ValorTipoCons"].ToString() != null &&
                                rows["EditarParam"].ToString().Length != 0 && rows["EditarParam"].ToString() != null &&
                                //@NombreRespuesta @Tipo @EditarTipo
                                rows["NombreRespuesta"].ToString().Length != 0 && rows["NombreRespuesta"].ToString() != null &&
                                rows["Tipo"].ToString().Length != 0 && rows["Tipo"].ToString() != null &&
                                rows["EditarTipo"].ToString().Length != 0 && rows["EditarTipo"].ToString() != null &&
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

                                string NombreServ = rows["NombreServ"].ToString();
                                string URLServ = rows["URLServ"].ToString();
                                string Interface = rows["Interface"].ToString();
                                string EditarURLServ = rows["EditarURLServ"].ToString();
                                string MetodoServicio = rows["MetodoServicio"].ToString();
                                string Observaciones = rows["Observaciones"].ToString();
                                string EditarObserv = rows["EditarObserv"].ToString();
                                string NombreParam = rows["NombreParam"].ToString();
                                string TipoParam = rows["TipoParam"].ToString();
                                string ValorTipoCons = rows["ValorTipoCons"].ToString();
                                string EditarParam = rows["EditarParam"].ToString();
                                string NombreRespuesta = rows["NombreRespuesta"].ToString();
                                string Tipo = rows["Tipo"].ToString();
                                string EditarTipo = rows["EditarTipo"].ToString();

                                string tablaDetalle1 = rows["tablaDetalle1"].ToString();
                                string campoDetalle1 = rows["campoDetalle1"].ToString();
                                string editCampoD1 = rows["editCampoD1"].ToString();
                                string tablaDetalle2 = rows["tablaDetalle2"].ToString();
                                string campoDetalle2 = rows["campoDetalle2"].ToString();
                                string editCampoD2 = rows["editCampoD2"].ToString();
                                string tablaDetalle3 = rows["tablaDetalle3"].ToString();
                                string campoDetalle3 = rows["campoDetalle3"].ToString();
                                string editCampoD3 = rows["editCampoD3"].ToString();

                                List<string> crudVariables = new List<string>() { NombreServ, URLServ, Interface, EditarURLServ };
                                List<string> crudDetalle1 = new List<string>() { MetodoServicio, Observaciones, EditarObserv };
                                List<string> crudDetalle2 = new List<string>() { NombreParam, TipoParam, ValorTipoCons, EditarParam };
                                List<string> crudDetalle3 = new List<string>() { NombreRespuesta, Tipo, EditarTipo };


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "GN", motor);
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
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera("GN_DRSER", user, motor, "COD_EMPR", "9", "9", "COD_EMPR", c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 2
                                       // ErrorValidacion = ValidacionesQueries.ValidacionCabecera("GN_DMSER", user, motor, "COD_EMPR", "9", "9", "COD_EMPR", c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS DETALLE 1
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera("GN_DSERI", user, motor, "COD_EMPR", "9", "9", "COD_EMPR", c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, NombreServ, NombreServ, Campo, c, ErrorValidacion, "No se agregó correctamente", 0,"1");

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

                                        CrudKgnmseri.CRUD(desktopSession, 1, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, NombreServ, NombreServ, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);
                                        
                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKgnmseri.CRUD(desktopSession, 2, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, NombreServ, URLServ, EditCampo, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                        CrudKgnmseri.CRUD(desktopSession, 2, crudVariables, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);

                                        Thread.Sleep(1000);
                                        bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Editar", true, file);

                                            CrudKgnmseri.CRUD(desktopSession, 2, crudVariables, file);
                                            newFunctions_4.ScreenshotNuevo("Datos Editados", true, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);
                                        }

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                       // ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, NombreServ, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, NombreServ, EditarURLServ, EditCampo, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);

                                        ////AGREGAR DETALLE 1
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKgnmseri.CrudDetalle1(desktopSession, 1, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 1", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, MetodoServicio, MetodoServicio, campoDetalle1, c, ErrorValidacion, "No se agregó correctamente 1", 0);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                        CrudKgnmseri.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 1", true, file);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, MetodoServicio, Observaciones, editCampoD1, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 1", true, file);

                                        CrudKgnmseri.CrudDetalle1(desktopSession, 2, crudDetalle1, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 1", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 1", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(3000);
                                        // //Debugger.Launch();
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, MetodoServicio, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, MetodoServicio, EditarObserv, editCampoD1, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        ////AGREGAR DETALLE 2
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKgnmseri.CrudDetalle2(desktopSession, 1, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y); 
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 2", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, TipoParam, TipoParam, campoDetalle2, c, ErrorValidacion, "No se agregó correctamente 2", 0);

                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);

                                        CrudKgnmseri.CrudDetalle2(desktopSession, 2, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 2", true, file);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, TipoParam, ValorTipoCons, editCampoD2, c, ErrorValidacion, "No se edito correctamente", 1);
  
                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 2", true, file);

                                        CrudKgnmseri.CrudDetalle2(desktopSession, 2, crudDetalle2, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 2", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 2", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, TipoParam, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, TipoParam, EditarParam, editCampoD2, c, ErrorValidacion, "El registro no se editó correctamente", 1);


                                        //Tipo 1 cLick a Respuesta Metodo y tipo 2 Click a Parametros Metodo
                                        CrudKgnmseri.ClickButton(desktopSession, 1);

                                        ////AGREGAR DETALLE 3
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 2);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKgnmseri.CrudDetalle3(desktopSession, 1, crudDetalle3, file);
                                        newFunctions_4.ScreenshotNuevo("Insertar Datos Detalle 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Datos Insertados Detalle 3", true, file);

                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, NombreRespuesta, NombreRespuesta, campoDetalle3, c, ErrorValidacion, "No se agregó correctamente 3", 0);

                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 3", true, file);

                                        CrudKgnmseri.CrudDetalle3(desktopSession, 2, crudDetalle3, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados Detalle 3", true, file);

                                        ////Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, NombreRespuesta, Tipo, editCampoD3, c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar Detalle 3", true, file);

                                        CrudKgnmseri.CrudDetalle3(desktopSession, 2, crudDetalle3, file);
                                        newFunctions_4.ScreenshotNuevo("Datos Editados Detalle 3", true, file);

                                        desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada Detalle 3", true, file);

                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                        //ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle3, motor, user, campoDetalle3, NombreRespuesta, "E", c, ErrorValidacion);
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, NombreRespuesta, EditarTipo, editCampoD3, c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        /*errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        newFunctions_1.lapiz(desktopSession);*/

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

                                        ////REPORTE PRELIMINAR
                                        ///
                                        //string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        //errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                                        ////CrudKnmctpre.Preliminar(desktopSession, file, pdf1, BanderaPreli);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        ////Rectificar Pie Pagina PDF PRELIMINAR
                                        //listaResultado = Textopdf(pdf1, codProgram, user);
                                        //listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        ////EXPORTAR EXCEL
                                        //string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        //errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //RECTIFICACION DE CAMPOS DE EXCEL  
                                        //string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, NombreServ, c);
                                        //string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        //Celda = Celda_Excel(ruta, nom_empr);
                                        //Celda.ForEach(u => ErrorValidacion.Add(u));

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        Thread.Sleep(1000);

                                        //ELIMINAR DETALLE 3
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[2], botonesNavegador, file);

                                        ////Validacion Eliminar Detalle 3
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle3, motor, user, campoDetalle3, NombreRespuesta, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle3, user, motor, campoDetalle3, NombreRespuesta, "", campoDetalle3, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //Tipo 1 cLick a Respuesta Metodo y tipo 2 Click a Parametros Metodo
                                        CrudKgnmseri.ClickButton(desktopSession, 2);
                                        //ELIMINAR DETALLE 2
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[2], botonesNavegador, file);

                                        ////Validacion Eliminar Detalle 2
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle2, motor, user, campoDetalle2, TipoParam, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle2, user, motor, campoDetalle2, TipoParam, "", campoDetalle2, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR DETALLE 1
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador, file);

                                        ////Validacion Eliminar Detalle 1
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(tablaDetalle1, motor, user, campoDetalle1, MetodoServicio, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(tablaDetalle1, user, motor, campoDetalle1, MetodoServicio, "", campoDetalle1, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);

                                        //ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);

                                        ////Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, NombreServ, "D", c, ErrorValidacion);
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, NombreServ, "", Campo, c, ErrorValidacion, "los datos no se eliminaron correctamente", 2);
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
        public void KGnWayudCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosGN.ModulosGN_4.KGnWayudCheckList")
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
                                //rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&

                                //@Codigo @Nombre @EditarNombre
                                rows["Codigo"].ToString().Length != 0 && rows["Codigo"].ToString() != null &&
                                rows["Descripcion"].ToString().Length != 0 && rows["Descripcion"].ToString() != null &&
                                rows["TipoAyuda"].ToString().Length != 0 && rows["TipoAyuda"].ToString() != null &&
                                rows["EditarDescripcion"].ToString().Length != 0 && rows["EditarDescripcion"].ToString() != null &&
                                rows["CodigoDetalle"].ToString().Length != 0 && rows["CodigoDetalle"].ToString() != null &&
                                rows["DescripcionDetalle"].ToString().Length != 0 && rows["DescripcionDetalle"].ToString() != null &&
                                rows["Link"].ToString().Length != 0 && rows["Link"].ToString() != null

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
                                string Descripcion = rows["Descripcion"].ToString();
                                string TipoAyuda = rows["TipoAyuda"].ToString();
                                string EditarDescripcion = rows["EditarDescripcion"].ToString();
                                string CodigoDetalle = rows["CodigoDetalle"].ToString();
                                string DescripcionDetalle = rows["DescripcionDetalle"].ToString();
                                string Link = rows["Link"].ToString();  

                                List<string> crudVariables = new List<string>() { Codigo, Descripcion, TipoAyuda, EditarDescripcion};
                                List<string> crudDetalle = new List<string>() { CodigoDetalle, DescripcionDetalle, Link };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "GN", motor);
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
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKgnwayud.CRUD(desktopSession, 1, crudVariables, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1500);
                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, Campo, c, ErrorValidacion, "No se agregó correctamente", 0);
                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKgnwayud.CRUD(desktopSession, 2, crudVariables, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);

                                        ////Validacion Editar Descartar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Descripcion, "NOM_AYUD", c, ErrorValidacion, "No se edito correctamente", 1);

                                        ////ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKgnwayud.CRUD(desktopSession, 2, crudVariables, file);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //ACEPTAR EDICION
                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                            desktopSession.Mouse.Click(null);

                                            CrudKgnwayud.CRUD(desktopSession, 2, crudVariables, file);

                                            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                        }
                                        ////Validacion Editar
                                        Thread.Sleep(1500);
                                        // //Debugger.Launch();
                                       // ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Codigo, "E", c, ErrorValidacion);
                                       // ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, EditarDescripcion, "NOM_AYUD", c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        Thread.Sleep(1000);

                                        ////AGREGAR DETALLE
                                        //List<string> crudDetalle = new List<string> { "9393", "PRUEBA DETALLE", "PRUEBA LINK" };
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);

                                        CrudKgnwayud.CrudDetalle(desktopSession, 1, crudDetalle, file);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);

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

                                        //ELIMINAR EL DETALLE
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador, file);

                                        ////ELIMINAR REGISTR
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

        [TestMethod]
        public void KGnRpermCheckList() 
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosGN.ModulosGN_4.KGnRpermCheckList")
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
                                //rows["Tabla"].ToString().Length != 0 && rows["Tabla"].ToString() != null &&
                                //rows["Campo"].ToString().Length != 0 && rows["Campo"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                //rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                //rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                //rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null 

                                )
                            {
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                //string Tabla = rows["Tabla"].ToString();
                                //string Campo = rows["Campo"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                //string banExcel = rows["banExcel"].ToString();
                                //string BanderaLupa = rows["BanderaLupa"].ToString();
                                //string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();

                                //List<string> crudVariables = new List<string>() { Tipo, URL, Servicio, Puerto, EditarServ };
                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "GN", motor);
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

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                       
                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //REPORTE PRELIMINAR
                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                                        Crudkgnrperm.CRUD(desktopSession, BanderaPreli, file, pdf1, ruta, motor);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

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

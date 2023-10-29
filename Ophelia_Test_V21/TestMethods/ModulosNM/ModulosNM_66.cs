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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosPn;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_7;

namespace Ophelia_Test_V21.TestMethods.ModulosNM
{
    /// <summary>
    /// Descripción resumida de ModulosNM_66
    /// </summary>
    [TestClass]
    public class ModulosNM_66 : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();

        public ModulosNM_66()
        { }

        private TestContext testContextInstance;
        [TestMethod]
        public void abrirprograma()
        {
            //////ARRANCAR EN SQL
            string motor = "SQL";
            AbrirPrograma a = new AbrirPrograma("KNmLdifa", "UCalida1");
            string file = CrearDocumentoWordDinamico("KNmLdifa", "PruebasCrud", motor);
            List<string> crudVariables = new List<string> { "222333444555" };

            ////ARRANCAR EN ORA
            //string motor = "ORA";
            //AbrirPrograma a = new AbrirPrograma("KNmDifin", "ODESAR");
            //string file = CrearDocumentoWordDinamico("KNmDifin", "PruebasCrud", motor);
            ////List<string> crudVariablesLupa = new List<string> {  };
            //List<string> crudVariables = new List<string> { "PRUEBA", "19/01/2021", "19/05/2021", "ACTUALIZACION", "ABC11111", "c", "1200000" };

            ////EJECUTAR EN EL NAVEGADOR
            desktopSession = a.Start2(motor, "");
            Thread.Sleep(20000000);

            //////////////AGREGAR REGISTRO
            //Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
            //botonesNavegador = PruebaCRUD.findname(desktopSession, 27, 0);
            //var ElementList = PruebaCRUD.NavClass(desktopSession);

            ////////////INSERTAR
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            //desktopSession.Mouse.Click(null);

            //////////////////APLICAR
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);

            //CrudKnmmcpar.OPCIONES(desktopSession, "Temporalidad", crudVariables, file);

            //////////AGREGAR REGISTRO DETALLE 1
            //Dictionary<string, Point> botonesNavegador1 = new Dictionary<string, Point>();
            //botonesNavegador1 = PruebaCRUD.findname(desktopSession, 29, 1);
            //var ElementList1 = PruebaCRUD.NavClass(desktopSession);

            ////////INSERTAR DETALLE 1
            //desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Insertar"].X, botonesNavegador1["Insertar"].Y);
            //desktopSession.Mouse.Click(null);

            //CrudKnmmcpar.CRUD_DETALLE2(desktopSession, "0", crudVariables, file);
            //Thread.Sleep(1000);

            //////////////APLICAR DETALLE 1
            //desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);

            ////////////////EDITAR CAMBIOS DETALLE 1 (EDITAR - CANCELAR)
            //desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
            //desktopSession.Mouse.Click(null);

            //CrudKnmmcpar.CRUD_DETALLE2(desktopSession, "1", crudVariables, file);
            //Thread.Sleep(1000);

            //////////////CANCELAR EDITAR DETALLE 1
            //desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Cancelar"].X, botonesNavegador1["Cancelar"].Y);
            //desktopSession.Mouse.Click(null);

            //////////////EDITAR CAMBIOS DETALLE 1 (EDITAR - APLICAR)
            //desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
            //desktopSession.Mouse.Click(null);

            //CrudKnmmcpar.CRUD_DETALLE2(desktopSession, "1", crudVariables, file);
            //Thread.Sleep(1000);

            //////////////APLICAR DETALLE 1
            //desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);

            //CrudKnmmcpar.OPCIONES(desktopSession, "SubComisiones", crudVariables, file);

            //////////AGREGAR REGISTRO DETALLE 2
            //Dictionary<string, Point> botonesNavegador2 = new Dictionary<string, Point>();
            //botonesNavegador2 = PruebaCRUD.findname(desktopSession, 29, 1);
            //var ElementList2 = PruebaCRUD.NavClass(desktopSession);

            ////////INSERTAR DETALLE 2
            //desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Insertar"].X, botonesNavegador2["Insertar"].Y);
            //desktopSession.Mouse.Click(null);

            //CrudKnmmcpar.CRUD_DETALLE3(desktopSession, "0", crudVariables, file);
            //Thread.Sleep(1000);

            //////////////APLICAR DETALLE 2
            //desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);

            ////////////////EDITAR CAMBIOS DETALLE 2 (EDITAR - CANCELAR)
            //desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
            //desktopSession.Mouse.Click(null);

            //CrudKnmmcpar.CRUD_DETALLE3(desktopSession, "1", crudVariables, file);
            //Thread.Sleep(1000);

            //////////////CANCELAR EDITAR DETALLE 2
            //desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Cancelar"].X, botonesNavegador2["Cancelar"].Y);
            //desktopSession.Mouse.Click(null);

            //////////////EDITAR CAMBIOS DETALLE 2 (EDITAR - APLICAR)
            //desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Editar"].X, botonesNavegador2["Editar"].Y);
            //desktopSession.Mouse.Click(null);

            //CrudKnmmcpar.CRUD_DETALLE3(desktopSession, "1", crudVariables, file);
            //Thread.Sleep(1000);

            //////////////APLICAR DETALLE 2
            //desktopSession.Mouse.MouseMove(ElementList2[1].Coordinates, botonesNavegador2["Aplicar"].X, botonesNavegador2["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);

            //CrudKnmmcpar.OPCIONES(desktopSession, "Variables", crudVariables, file);

            //////////AGREGAR REGISTRO DETALLE 3
            //Dictionary<string, Point> botonesNavegador3 = new Dictionary<string, Point>();
            //botonesNavegador3 = PruebaCRUD.findname(desktopSession, 29, 1);
            //var ElementList3 = PruebaCRUD.NavClass(desktopSession);

            ////////INSERTAR DETALLE 3
            //desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Insertar"].X, botonesNavegador3["Insertar"].Y);
            //desktopSession.Mouse.Click(null);

            //CrudKnmmcpar.CRUD_DETALLE4(desktopSession, "0", crudVariables, file);
            //Thread.Sleep(1000);

            //////////////APLICAR DETALLE 3
            //desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Aplicar"].X, botonesNavegador3["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);

            ////////////////EDITAR CAMBIOS DETALLE 3 (EDITAR - CANCELAR)
            //desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Editar"].X, botonesNavegador3["Editar"].Y);
            //desktopSession.Mouse.Click(null);

            //CrudKnmmcpar.CRUD_DETALLE4(desktopSession, "1", crudVariables, file);
            //Thread.Sleep(1000);

            //////////////CANCELAR EDITAR DETALLE 3
            //desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Cancelar"].X, botonesNavegador3["Cancelar"].Y);
            //desktopSession.Mouse.Click(null);

            //////////////EDITAR CAMBIOS DETALLE 3 (EDITAR - APLICAR)
            //desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Editar"].X, botonesNavegador3["Editar"].Y);
            //desktopSession.Mouse.Click(null);

            //CrudKnmmcpar.CRUD_DETALLE4(desktopSession, "1", crudVariables, file);
            //Thread.Sleep(1000);

            //////////////APLICAR DETALLE 3
            //desktopSession.Mouse.MouseMove(ElementList3[1].Coordinates, botonesNavegador3["Aplicar"].X, botonesNavegador3["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);

            //CrudKnmmcpar.OPCIONES(desktopSession, "Generales", crudVariables, file);

            //////////////////EDITAR CAMBIOS (EDITAR - CANCELAR)
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            //desktopSession.Mouse.Click(null);

            //CrudKnmmcpar.CRUD_DETALLE1(desktopSession, "1", crudVariables, file);
            //Thread.Sleep(1000);

            ////////////////CANCELAR EDITAR
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            //desktopSession.Mouse.Click(null);

            //newFunctions_3.LupaAud(desktopSession, "0", file);

            //////////////////EDITAR CAMBIOS (EDITAR - APLICAR)
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            //desktopSession.Mouse.Click(null);

            //CrudKnmmcpar.CRUD_DETALLE1(desktopSession, "1", crudVariables, file);
            //Thread.Sleep(1000);

            ////////////////APLICAR EDITAR
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            //desktopSession.Mouse.Click(null);

            //CrudKnmmcpar.OPCIONES(desktopSession, "Variables", crudVariables, file);

            ///////////////// ELIMINAR REGISTRO DETALLE 3
            //CrudKnmmcpar.EliminarRegistro(desktopSession, ElementList3[1], botonesNavegador3, file);

            //CrudKnmmcpar.OPCIONES(desktopSession, "SubComisiones", crudVariables, file);

            ///////////////// ELIMINAR REGISTRO DETALLE 2
            //CrudKnmmcpar.EliminarRegistro(desktopSession, ElementList2[1], botonesNavegador2, file);

            //CrudKnmmcpar.OPCIONES(desktopSession, "Temporalidad", crudVariables, file);

            ///////////////// ELIMINAR REGISTRO DETALLE 1
            //CrudKnmmcpar.EliminarRegistro(desktopSession, ElementList1[1], botonesNavegador1, file);

            //CrudKnmmcpar.OPCIONES(desktopSession, "Generales", crudVariables, file);
            //newFunctions_3.LupaAud(desktopSession, "0", file);

            ///////////// ELIMINAR REGISTRO
            //PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
        }
        [TestMethod]
        public void KNmLdifaCheckList()
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
            //TableOrder = "ktest1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_66.KNmLdifaCheckList")
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
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null 
                                
                                )
                            {
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                

                                List<string> crudVariables = new List<string>() { };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);
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

                                        ////VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);
   
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
        public void KNmLssecCheckList()
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
            //TableOrder = "ktest1";
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosNM.ModulosNM_66.KNmLssecCheckList")
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
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null 
                               
                                )
                            {
                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                

                                List<string> crudVariables = new List<string>() {};

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                string file = CrearDocumentoWordDinamico(codProgram, "NM", motor);
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

                                        ////VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);
                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

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
    }
}

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
using Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi;
using OpenQA.Selenium.Interactions;
using System.Windows.Forms;
using OpenQA.Selenium;

namespace Ophelia_Test_V21.TestMethods.ModulosBI
{
    /// <summary>
    /// Descripción resumida de ModulosBi_8
    /// </summary>
    [TestClass]
    public class ModulosBi_8 : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();
        public ModulosBi_8() { }

        [TestMethod]
        public void AbrirPrograma()
        {
            string programa = "KBiRoles";
            AbrirPrograma a = new AbrirPrograma(programa, "UNatalia");
            desktopSession = a.Start2("SQL", "");
            //  AbrirPrograma a = new AbrirPrograma(programa, "TPRUEBAS");
            //  desktopSession = a.Start2("ORA", "");
        }

        [TestMethod]
        public void KBiRolesCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosBI.ModulosBi_8.KBiRolesCheckList")
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
                            //int velocidad = 10;

                            //Playback.PlaybackSettings.DelayBetweenActions = velocidad;

                            //@User @Motor @CodProgram @NomProgram @TipoQbe @QbeFilter @banExcel @BanderaLupa @BanderaSum @BanderaPreli
                            //@Identificacion @Concepto @Ausencia @Observaciones

                            if (
                                //Datos Login
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&

                                rows["L1Nivel"].ToString().Length != 0 && rows["L1Nivel"].ToString() != null &&
                                rows["Rol"].ToString().Length != 0 && rows["Rol"].ToString() != null &&
                                rows["Ncargos"].ToString().Length != 0 && rows["Ncargos"].ToString() != null &&
                                rows["FechaCre"].ToString().Length != 0 && rows["FechaCre"].ToString() != null &&
                                rows["NcargosEditar"].ToString().Length != 0 && rows["NcargosEditar"].ToString() != null &&
                                rows["Descri"].ToString().Length != 0 && rows["Descri"].ToString() != null &&

                                rows["Mision"].ToString().Length != 0 && rows["Mision"].ToString() != null &&

                                rows["DesFunciones"].ToString().Length != 0 && rows["DesFunciones"].ToString() != null &&
                                rows["DesFuncionesEditar"].ToString().Length != 0 && rows["DesFuncionesEditar"].ToString() != null &&

                                rows["Requisitos"].ToString().Length != 0 && rows["Requisitos"].ToString() != null &&

                                rows["Perfil"].ToString().Length != 0 && rows["Perfil"].ToString() != null &&
                                rows["PerfilEditar"].ToString().Length != 0 && rows["PerfilEditar"].ToString() != null &&

                                 rows["FechI"].ToString().Length != 0 && rows["FechI"].ToString() != null &&
                                rows["FechH"].ToString().Length != 0 && rows["FechH"].ToString() != null &&
                                rows["SalarioMinME"].ToString().Length != 0 && rows["SalarioMinME"].ToString() != null &&
                                rows["SalarioME"].ToString().Length != 0 && rows["SalarioME"].ToString() != null &&
                                rows["SalarioMiaxME"].ToString().Length != 0 && rows["SalarioMiaxME"].ToString() != null &&
                                rows["SalarioMinMM"].ToString().Length != 0 && rows["SalarioMinMM"].ToString() != null &&
                                rows["SalarioMM"].ToString().Length != 0 && rows["SalarioMM"].ToString() != null &&
                                rows["SalarioMaxMM"].ToString().Length != 0 && rows["SalarioMaxMM"].ToString() != null &&
                                rows["SalarioMinAE"].ToString().Length != 0 && rows["SalarioMinAE"].ToString() != null &&
                                rows["SalarioAE"].ToString().Length != 0 && rows["SalarioAE"].ToString() != null &&
                                rows["SalarioMaxAE"].ToString().Length != 0 && rows["SalarioMaxAE"].ToString() != null &&
                                rows["SalarioMinAM"].ToString().Length != 0 && rows["SalarioMinAM"].ToString() != null &&
                                rows["SalarioAM"].ToString().Length != 0 && rows["SalarioAM"].ToString() != null &&
                                rows["SalarioMaxAM"].ToString().Length != 0 && rows["SalarioMaxAM"].ToString() != null &&
                                rows["FechEdi"].ToString().Length != 0 && rows["FechEdi"].ToString() != null &&

                                rows["Descripcion"].ToString().Length != 0 && rows["Descripcion"].ToString() != null &&
                                rows["DescripcionEditar"].ToString().Length != 0 && rows["DescripcionEditar"].ToString() != null &&

                                rows["ConocimientosB"].ToString().Length != 0 && rows["ConocimientosB"].ToString() != null &&
                                rows["ConocimientosBEditar"].ToString().Length != 0 && rows["ConocimientosBEditar"].ToString() != null 
                                )
                            {
                                //VARIABLES GENERALES
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
                                //VARIABLES CRUD                                
                                string L1Nivel = rows["L1Nivel"].ToString();  
                                string Rol = rows["Rol"].ToString();
                                string Ncargos = rows["Ncargos"].ToString();
                                string FechaCre = rows["FechaCre"].ToString();
                                string NcargosEditar = rows["NcargosEditar"].ToString();
                                string Descri = rows["Descri"].ToString();
                                //VARIABLES CRUD DETALLE 1                            
                                string Mision = rows["Mision"].ToString();
                                //VARIABLES CRUD DETALLE 2                   
                                string DesFunciones = rows["DesFunciones"].ToString();
                                string DesFuncionesEditar = rows["DesFuncionesEditar"].ToString();
                                //VARIABLES CRUD DETALLE 3                            
                                string Requisitos = rows["Requisitos"].ToString();
                                //VARIABLES CRUD DETALLE 4                
                                string Perfil = rows["Perfil"].ToString();
                                string PerfilEditar = rows["PerfilEditar"].ToString();
                                //VARIABLES CRUD DETALLE 5                          
                                string FechI = rows["FechI"].ToString();
                                string FechH = rows["FechH"].ToString();
                                string SalarioMinME = rows["SalarioMinME"].ToString();
                                string SalarioME = rows["SalarioME"].ToString();
                                string SalarioMiaxME = rows["SalarioMiaxME"].ToString();
                                string SalarioMinMM = rows["SalarioMinMM"].ToString();   
                                string SalarioMM = rows["SalarioMM"].ToString();
                                string SalarioMaxMM = rows["SalarioMaxMM"].ToString();
                                string SalarioMinAE = rows["SalarioMinAE"].ToString();
                                string SalarioAE = rows["SalarioAE"].ToString();
                                string SalarioMaxAE = rows["SalarioMaxAE"].ToString();
                                string SalarioMinAM = rows["SalarioMinAM"].ToString();
                                string SalarioAM = rows["SalarioAM"].ToString();
                                string SalarioMaxAM = rows["SalarioMaxAM"].ToString();
                                string FechEdi = rows["FechEdi"].ToString();
                                //VARIABLES CRUD DETALLE 6              
                                string Descripcion = rows["Descripcion"].ToString();
                                string DescripcionEditar = rows["DescripcionEditar"].ToString();
                                //VARIABLES CRUD DETALLE 7
                                string ConocimientosB = rows["ConocimientosB"].ToString();
                                string ConocimientosBEditar = rows["ConocimientosBEditar"].ToString();


                                List<string> crudVariables = new List<string>() { L1Nivel ,Rol ,Ncargos ,FechaCre ,NcargosEditar,Descri };
                                List<string> crudVariablesDetalle1 = new List<string>() { Mision };
                                List<string> crudVariablesDetalle2 = new List<string>() { DesFunciones , DesFuncionesEditar };
                                List<string> crudVariablesDetalle3 = new List<string>() { Requisitos };
                                List<string> crudVariablesDetalle4 = new List<string>() { Perfil, PerfilEditar };
                                List<string> crudVariablesDetalle5 = new List<string>() { FechI ,FechH,SalarioMinME ,SalarioME ,SalarioMiaxME,SalarioMinMM ,SalarioMM,SalarioMaxMM ,SalarioMinAE  ,SalarioAE ,SalarioMaxAE ,SalarioMinAM ,SalarioAM ,SalarioMaxAM  ,FechEdi };
                                List<string> crudVariablesDetalle6 = new List<string>() { Descripcion,DescripcionEditar };
                                List<string> crudVariablesDetalle7 = new List<string>() { ConocimientosB ,ConocimientosBEditar };


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                string file = CrearDocumentoWordDinamico(codProgram, "GH", motor);

                                try
                                {

                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();
                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
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

                                        WindowsDriver<WindowsElement> rootSession = null;

                                        ////////////////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbiroles.CRUDKBiRoles(desktopSession, 1, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);                           
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);

                                        //////Validacion Agregar
                                        ////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbiroles.CRUDKBiRoles(desktopSession, 2, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);


                                        ////Validacion Editar Descartar
                                        ////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, "OBS_ERVA", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbiroles.CRUDKBiRoles(desktopSession, 2, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        //PASAR A DETALLE 
                                        var Element1 = desktopSession.FindElementsByClassName("TPageControl");
                                        desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 7);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);

                                        ///////////BARRA DETALLE 1//////
                                        var Element2 = desktopSession.FindElementsByClassName("TTabSheet");
                                        var Element3 = Element2[0].FindElementsByClassName("TKNavegador");
                                        int Pasod = 1;

                                        CrudKbiroles.CRUDDetalle1KBiRoles(desktopSession, crudVariablesDetalle1, file);
                                        Thread.Sleep(500);

                                        //bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        rootSession = PruebaCRUD.RootSession();
                                        rootSession = PruebaCRUD.ReloadSession(rootSession, "TReconcileErrorForm");
                                        //var ventana2 = rootSession.FindElementsByClassName("TReconcileErrorForm");
                                        bool ValEdit = false;
                                        //if (ventana2.Count > 0) { rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter); ValEdit = true; }
                                        if (ValEdit == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //EDITAR
                                            desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 5, 15);
                                            desktopSession.Mouse.Click(null);

                                            CrudKbiroles.CRUDDetalle1KBiRoles(desktopSession, crudVariablesDetalle1, file);

                                            //ACEPTAR
                                            desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 30, 15);
                                            desktopSession.Mouse.Click(null);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Agregar detalle 1", true, file);

                                        //PASAR A DETALLE 2
                                        Element1 = desktopSession.FindElementsByClassName("TPageControl");
                                        desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 7);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);

                                        /////////BARRA DETALLE 2//////
                                        Element2 = desktopSession.FindElementsByClassName("TTabSheet");
                                        Element3 = Element2[0].FindElementsByClassName("TKNavegador");

                                        desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 110, 15);
                                        desktopSession.Mouse.Click(null);

                                        CrudKbiroles.CRUDDetalle2KBiRoles(desktopSession, crudVariablesDetalle2, file);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Agregar detalle 2", true, file);

                                        //PASAR A DETALLE 3
                                        Element1 = desktopSession.FindElementsByClassName("TPageControl");
                                        desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 7);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Mouse.Click(null);
                                        for (Pasod = 1; Pasod <= 3; Pasod++)
                                        {
                                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        }

                                        ///////BARRA DETALLE 3//////
                                        Element2 = desktopSession.FindElementsByClassName("TTabSheet");
                                        Element3 = Element2[0].FindElementsByClassName("TKNavegador");

                                        CrudKbiroles.CRUDDetalle3KBiRoles(desktopSession, crudVariablesDetalle3, file);
                                        bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
                                        if (valEditOra == true)
                                        {
                                            //LUPA
                                            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            //EDITAR
                                            desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 5, 15);
                                            desktopSession.Mouse.Click(null);
                                            CrudKbiroles.CRUDDetalle3KBiRoles(desktopSession, crudVariablesDetalle3, file);
                                            //ACEPTAR
                                            desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 30, 15);
                                            desktopSession.Mouse.Click(null);
                                        }
                                        newFunctions_4.ScreenshotNuevo("Agregar detalle 3", true, file);

                                        //PASAR A DETALLE 4
                                        Element1 = desktopSession.FindElementsByClassName("TPageControl");
                                        desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 7);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        for (Pasod = 1; Pasod <= 4; Pasod++)
                                        {
                                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        }

                                        ///////BARRA DETALLE 4//////
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 25, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbiroles.CRUDDetalle4KBiRoles(desktopSession, 1, crudVariablesDetalle4, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);   
                                        

                                        rootSession = PruebaCRUD.RootSession();
                                        rootSession = PruebaCRUD.ReloadSession(rootSession, "TMessageForm");
                                        var okboton = rootSession.FindElementsByName("OK");
                                        rootSession.Mouse.MouseMove(okboton[0].Coordinates, 35, 10);
                                        rootSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Agregar detalle 4", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        //rootSession = PruebaCRUD.RootSession();
                                        //rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                                        //var allFrame = rootSession.FindElementsByClassName("IEFrame");
                                        //new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 15, (allFrame[0].Size.Height / 2) + 30).DoubleClick().Click().Perform();
                                        CrudKbiroles.CRUDDetalle4KBiRoles(desktopSession, 2, crudVariablesDetalle4, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar detalle 4", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbiroles.CRUDDetalle4KBiRoles(desktopSession, 2, crudVariablesDetalle4, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        //rootSession = PruebaCRUD.RootSession();
                                        //rootSession = PruebaCRUD.ReloadSession(rootSession, "TMessageForm");
                                        //var okboton2 = rootSession.FindElementsByName("OK");
                                        //rootSession.Mouse.MouseMove(okboton2[0].Coordinates, 35, 10);
                                        //rootSession.Mouse.Click(null);
                                        //Thread.Sleep(500);
                                        Thread.Sleep(500);
                                        CrudKbiroles.ventanaok(desktopSession);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar detalle 4", true, file);

                                        //PASAR A DETALLE 5
                                        Element1 = desktopSession.FindElementsByClassName("TPageControl");
                                        desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 10);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        for (Pasod = 1; Pasod <= 5; Pasod++)
                                        {
                                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        }

                                        ///////BARRA DETALLE 5//////
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 25, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbiroles.CRUDDetalle5KBiRoles(desktopSession, 1, crudVariablesDetalle5, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Agregar detalle 5", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbiroles.CRUDDetalle5KBiRoles(desktopSession, 2, crudVariablesDetalle5, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar detalle 5", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbiroles.CRUDDetalle5KBiRoles(desktopSession, 2, crudVariablesDetalle5, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar detalle 5", true, file);

                                        //PASAR A DETALLE 6
                                        Element1 = desktopSession.FindElementsByClassName("TPageControl");
                                        desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 7);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        for (Pasod = 1; Pasod <= 6; Pasod++)
                                        {
                                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        }

                                        ///////BARRA DETALLE 6//////
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 25, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbiroles.CRUDDetalle6KBiRoles(desktopSession, 1, crudVariablesDetalle6, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Agregar detalle 6", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbiroles.CRUDDetalle6KBiRoles(desktopSession, 2, crudVariablesDetalle6, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar detalle 6", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbiroles.CRUDDetalle6KBiRoles(desktopSession, 2, crudVariablesDetalle6, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar detalle 6", true, file);

                                        //PASAR A DETALLE 7
                                        Element1 = desktopSession.FindElementsByClassName("TPageControl");
                                        desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 7);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        for (Pasod = 1; Pasod <= 7; Pasod++)
                                        {
                                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        }

                                        ///////BARRA DETALLE 7//////
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 25, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbiroles.CRUDDetalle7KBiRoles(desktopSession, 1, crudVariablesDetalle6, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Agregar detalle 7", true, file);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbiroles.CRUDDetalle7KBiRoles(desktopSession, 2, crudVariablesDetalle6, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar detalle 7", true, file);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbiroles.CRUDDetalle7KBiRoles(desktopSession, 2, crudVariablesDetalle6, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar detalle 7", true, file);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////SUMATORIA
                                        //errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //newFunctions_1.lapiz(desktopSession);

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
                                        newFunctions_1.lapiz(desktopSession);

                                        //ELIMINAR REGISTROS
                                        Element1 = desktopSession.FindElementsByClassName("TPageControl");
                                        desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 7);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        for (Pasod = 1; Pasod <= 7; Pasod++)
                                        {
                                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        }
                                        int deta = 7;
                                        for (Pasod = 1; Pasod <= 4; Pasod++)
                                        {
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(500);
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador, file);
                                            Thread.Sleep(500);
                                            newFunctions_4.ScreenshotNuevo("Eliminar detalle "+deta, true, file);
                                            deta -= 1;

                                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Left);
                                        }
                                        Thread.Sleep(500);
                                        CrudKbiroles.ventanaok(desktopSession);
                                        Thread.Sleep(500);
                                        rootSession = PruebaCRUD.RootSession();
                                        rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                                        //allFrame = rootSession.FindElementsByClassName("IEFrame");
                                        //new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 15, (allFrame[0].Size.Height / 2) + 30).DoubleClick().Click().Perform();
                                        rootSession = PruebaCRUD.RootSession();

                                        Element1 = desktopSession.FindElementsByClassName("TPageControl");
                                        desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 7);
                                        desktopSession.Mouse.Click(null);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        for (Pasod = 1; Pasod <= 2; Pasod++)
                                        {
                                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                                        }
                                        //Thread.Sleep(500);
                                        //CrudKbiroles.ventanaok(desktopSession);
                                        //Thread.Sleep(500);
                                        //ELIMINAR REGISTRO DETALLE 3
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador, file);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Eliminar detalle 3", true, file);

                                        //ELIMINAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        Thread.Sleep(500);
                                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                        Thread.Sleep(500);
                                        rootSession = PruebaCRUD.RootSession();
                                        rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                                        //allFrame = rootSession.FindElementsByClassName("IEFrame");
                                        //new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 15, (allFrame[0].Size.Height / 2) + 110).DoubleClick().Click().Perform();
                                        rootSession = PruebaCRUD.RootSession();
                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        Thread.Sleep(500);
                                        Thread.Sleep(500);
                                        newFunctions_4.ScreenshotNuevo("Eliminar Registro", true, file);

                                        //Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Rol, "", "COD_EMPL", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);


                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Rol, "D", c, ErrorValidacion);

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
        public void KBiRolesAbrirPrograma()
        {
            string file = CrearDocumentoWordDinamico("KBiRoles", "GH", "SQL");
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();

            //string TipoQbe = "0";
            //string QbeFilter = "123456";
            //Thread.Sleep(1000);
            //errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
            //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            //Thread.Sleep(1000);
            //newFunctions_3.BotonImprimir(desktopSession, file);
            //Thread.Sleep(1000);

            AbrirPrograma a = new AbrirPrograma("KBiRoles", "UCalida1");
            desktopSession = a.Start2("SQL", "");

            List<string> crudVariables = new List<string>() { "1", "0808", "1", "08/02/2020", "2" };
            List<string> crudVariablesDetalle1 = new List<string>() { "Pruebad1" };
            List<string> crudVariablesDetalle2 = new List<string>() { "Pruebad2", "Pruebad2Editar" };
            List<string> crudVariablesDetalle3 = new List<string>() { "Pruebad3" };
            List<string> crudVariablesDetalle4 = new List<string>() { "0125", "08" };
            List<string> crudVariablesDetalle5 = new List<string>() { "02/08/2020", "02/09/2020", "123456", "67890", "10203", "20301", "30102", "50406", "60450", "90874", "10478", "60326", "70152", "23456", "07/10/2020" };
            List<string> crudVariablesDetalle6 = new List<string>() { "1", "5" };
            List<string> crudVariablesDetalle7 = new List<string>() { "2", "17" };


            WindowsDriver<WindowsElement> rootSession = null;

            string BanderaLupa = "1";
            //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
            //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

            ////////////////AGREGAR REGISTRO
            Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
            botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
            var ElementList = PruebaCRUD.NavClass(desktopSession);

            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            desktopSession.Mouse.Click(null);

            CrudKbiroles.CRUDKBiRoles(desktopSession, 1, crudVariables, file);

            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(200);
            ////Validacion Agregar
            //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);

            CrudKbiroles.CRUDKBiRoles(desktopSession, 2, crudVariables, file);

            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);

            ////Validacion Editar Descartar
            ////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, "OBS_ERVA", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);

            CrudKbiroles.CRUDKBiRoles(desktopSession, 2, crudVariables, file);

            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);

            //PASAR A DETALLE 
            var Element1 = desktopSession.FindElementsByClassName("TPageControl");
            desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 7);
            desktopSession.Mouse.Click(null);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);

            ///////////BARRA DETALLE 1//////
            var Element2 = desktopSession.FindElementsByClassName("TTabSheet");
            var Element3 = Element2[0].FindElementsByClassName("TKNavegador");
            int Pasod = 1;

            CrudKbiroles.CRUDDetalle1KBiRoles(desktopSession, crudVariablesDetalle1, file);
            Thread.Sleep(500);

            //bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "TReconcileErrorForm");
            var ventana2 = rootSession.FindElementsByClassName("TReconcileErrorForm");
            bool ValEdit = false;

            if (ventana2.Count > 0) { rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter); ValEdit = true; }
            if (ValEdit == true)
            {
                //LUPA
                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                //EDITAR
                desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 5, 15);
                desktopSession.Mouse.Click(null);

                CrudKbiroles.CRUDDetalle1KBiRoles(desktopSession, crudVariablesDetalle1, file);

                //ACEPTAR
                desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 30, 15);
                desktopSession.Mouse.Click(null);
            }

            //PASAR A DETALLE 2
            Element1 = desktopSession.FindElementsByClassName("TPageControl");
            desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 7);
            desktopSession.Mouse.Click(null);
            desktopSession.Mouse.Click(null);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);

            /////////BARRA DETALLE 2//////
            Element2 = desktopSession.FindElementsByClassName("TTabSheet");
            Element3 = Element2[0].FindElementsByClassName("TKNavegador");

            desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 110, 15);
            desktopSession.Mouse.Click(null);

            CrudKbiroles.CRUDDetalle2KBiRoles(desktopSession, crudVariablesDetalle2, file);

            //PASAR A DETALLE 3
            Element1 = desktopSession.FindElementsByClassName("TPageControl");
            desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 7);
            desktopSession.Mouse.Click(null);
            desktopSession.Mouse.Click(null);
            for (Pasod = 1; Pasod <= 3; Pasod++)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
            }

            ///////BARRA DETALLE 3//////
            Element2 = desktopSession.FindElementsByClassName("TTabSheet");
            Element3 = Element2[0].FindElementsByClassName("TKNavegador");

            CrudKbiroles.CRUDDetalle3KBiRoles(desktopSession, crudVariablesDetalle3, file);

            bool valEditOra = PruebaCRUD.ValEditOra(desktopSession);
            if (valEditOra == true)
            {
                //LUPA
                errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                //EDITAR
                desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 5, 15);
                desktopSession.Mouse.Click(null);

                CrudKbiroles.CRUDDetalle3KBiRoles(desktopSession, crudVariablesDetalle3, file);

                //ACEPTAR
                desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 30, 15);
                desktopSession.Mouse.Click(null);
            }

            //PASAR A DETALLE 4
            Element1 = desktopSession.FindElementsByClassName("TPageControl");
            desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 7);
            desktopSession.Mouse.Click(null);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            for (Pasod = 1; Pasod <= 4; Pasod++)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
            }

            ///////BARRA DETALLE 4//////
            botonesNavegador = PruebaCRUD.findname(desktopSession, 25, 1);
            ElementList = PruebaCRUD.NavClass(desktopSession);

            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            desktopSession.Mouse.Click(null);

            CrudKbiroles.CRUDDetalle4KBiRoles(desktopSession, 1, crudVariablesDetalle4, file);

            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);

            Thread.Sleep(1000);
            //rootSession = PruebaCRUD.RootSession();
            //rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
            //var allFrame = rootSession.FindElementsByClassName("IEFrame");
            //new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 15, (allFrame[0].Size.Height / 2) + 30).DoubleClick().Click().Perform();
            //rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "TMessageForm");
            var okboton = rootSession.FindElementsByName("OK");
            rootSession.Mouse.MouseMove(okboton[0].Coordinates, 35, 10);
            rootSession.Mouse.Click(null);
            Thread.Sleep(500);

            //Validacion Agregar
            //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);


            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
            var allFrame = rootSession.FindElementsByClassName("IEFrame");
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 15, (allFrame[0].Size.Height / 2) + 30).DoubleClick().Click().Perform();

            CrudKbiroles.CRUDDetalle4KBiRoles(desktopSession, 2, crudVariablesDetalle4, file);

            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);

            ////Validacion Editar Descartar
            //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, "OBS_ERVA", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);

            CrudKbiroles.CRUDDetalle4KBiRoles(desktopSession, 2, crudVariablesDetalle4, file);

            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);

            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "TMessageForm");
            okboton = rootSession.FindElementsByName("OK");
            rootSession.Mouse.MouseMove(okboton[0].Coordinates, 35, 10);
            rootSession.Mouse.Click(null);
            Thread.Sleep(500);

            //PASAR A DETALLE 5
            Element1 = desktopSession.FindElementsByClassName("TPageControl");
            desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 10);
            desktopSession.Mouse.Click(null);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            for (Pasod = 1; Pasod <= 5; Pasod++)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
            }

            ///////BARRA DETALLE 5//////
            botonesNavegador = PruebaCRUD.findname(desktopSession, 25, 1);
            ElementList = PruebaCRUD.NavClass(desktopSession);

            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            desktopSession.Mouse.Click(null);

            CrudKbiroles.CRUDDetalle5KBiRoles(desktopSession, 1, crudVariablesDetalle5, file);

            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);

            //Validacion Agregar
            //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);

            CrudKbiroles.CRUDDetalle5KBiRoles(desktopSession, 2, crudVariablesDetalle5, file);

            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);

            ////Validacion Editar Descartar
            ////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, "OBS_ERVA", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);

            CrudKbiroles.CRUDDetalle5KBiRoles(desktopSession, 2, crudVariablesDetalle5, file);

            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);

            //PASAR A DETALLE 6
            Element1 = desktopSession.FindElementsByClassName("TPageControl");
            desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 7);
            desktopSession.Mouse.Click(null);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            for (Pasod = 1; Pasod <= 6; Pasod++)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
            }

            ///////BARRA DETALLE 6//////
            botonesNavegador = PruebaCRUD.findname(desktopSession, 25, 1);
            ElementList = PruebaCRUD.NavClass(desktopSession);

            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            desktopSession.Mouse.Click(null);

            CrudKbiroles.CRUDDetalle6KBiRoles(desktopSession, 1, crudVariablesDetalle6, file);

            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);

            ////Validacion Agregar
            //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);

            CrudKbiroles.CRUDDetalle6KBiRoles(desktopSession, 2, crudVariablesDetalle6, file);

            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);

            ////Validacion Editar Descartar
            ////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, "OBS_ERVA", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);

            CrudKbiroles.CRUDDetalle6KBiRoles(desktopSession, 2, crudVariablesDetalle6, file);

            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);

            //PASAR A DETALLE 7
            Element1 = desktopSession.FindElementsByClassName("TPageControl");
            desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 7);
            desktopSession.Mouse.Click(null);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            for (Pasod = 1; Pasod <= 7; Pasod++)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
            }

            ///////BARRA DETALLE 7//////
            botonesNavegador = PruebaCRUD.findname(desktopSession, 25, 1);
            ElementList = PruebaCRUD.NavClass(desktopSession);

            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            desktopSession.Mouse.Click(null);

            CrudKbiroles.CRUDDetalle7KBiRoles(desktopSession, 1, crudVariablesDetalle6, file);

            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);

            ////Validacion Agregar
            //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Codigo, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);

            //DESCARTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);

            CrudKbiroles.CRUDDetalle7KBiRoles(desktopSession, 2, crudVariablesDetalle6, file);

            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);

            ////Validacion Editar Descartar
            ////ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Codigo, Nombre, "OBS_ERVA", c, ErrorValidacion, "No se edito correctamente", 1);

            //ACEPTAR EDICION
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
            desktopSession.Mouse.Click(null);

            CrudKbiroles.CRUDDetalle7KBiRoles(desktopSession, 2, crudVariablesDetalle6, file);

            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            desktopSession.Mouse.Click(null);

            //ELIMINAR REGISTROS
            Element1 = desktopSession.FindElementsByClassName("TPageControl");
            desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 7);
            desktopSession.Mouse.Click(null);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            for (Pasod = 1; Pasod <= 7; Pasod++)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
            }

            for (Pasod = 1; Pasod <= 4; Pasod++)
            {
                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador, file);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Left);
            }

            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
            allFrame = rootSession.FindElementsByClassName("IEFrame");
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 15, (allFrame[0].Size.Height / 2) + 30).DoubleClick().Click().Perform();
            rootSession = PruebaCRUD.RootSession();

            Element1 = desktopSession.FindElementsByClassName("TPageControl");
            desktopSession.Mouse.MouseMove(Element1[0].Coordinates, 25, 7);
            desktopSession.Mouse.Click(null);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            for (Pasod = 1; Pasod <= 2; Pasod++)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
            }
            
            //ELIMINAR REGISTRO DETALLE 3
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador, file);
            Thread.Sleep(1000);


            //ELIMINAR REGISTRO
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
            Thread.Sleep(500);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(500);
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
            allFrame = rootSession.FindElementsByClassName("IEFrame");
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 15, (allFrame[0].Size.Height / 2) + 110).DoubleClick().Click().Perform();
            rootSession = PruebaCRUD.RootSession();
            //LUPA
            errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
            Thread.Sleep(500);
        }

        [TestMethod]
        public void KBiEmtalCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosBI.ModulosBi_8.KBiEmtalCheckList")
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
                            //int velocidad = 10;

                            //Playback.PlaybackSettings.DelayBetweenActions = velocidad;

                            //@User @Motor @CodProgram @NomProgram @TipoQbe @QbeFilter @banExcel @BanderaLupa @BanderaSum @BanderaPreli
                            //@Identificacion @Concepto @Ausencia @Observaciones

                            if (
                                //Datos Login
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&

                                rows["Identificacion"].ToString().Length != 0 && rows["Identificacion"].ToString() != null &&
                                rows["TallaP"].ToString().Length != 0 && rows["TallaP"].ToString() != null 
                                )
                            {
                                //VARIABLES GENERALES
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
                                //VARIABLES CRUD                                
                                string Identificacion = rows["Identificacion"].ToString();
                                string TallaP = rows["TallaP"].ToString();

                                List<string> crudVariables = new List<string>() { Identificacion ,TallaP };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                string file = CrearDocumentoWordDinamico(codProgram, "BI", motor);

                                try
                                {

                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();
                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
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
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Identificacion, Identificacion, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0,"1");

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        ////////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbiemtal.CRUDKBiEmtal(desktopSession, 1, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

                                        //Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Identificacion, Identificacion, "COD_EMPL", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbiemtal.CRUDKBiEmtal(desktopSession, 2, crudVariables, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        //Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Identificacion, "10", "COD_TALL", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(500);
                                        //CrudKbicarpe.ventanaSum(desktopSession);
                                        //Thread.Sleep(500);
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
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //RECTIFICACION DE CAMPOS DE EXCEL       
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, Identificacion, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        /*Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));*/

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        
                                        //ELIMINAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        Thread.Sleep(1000);

                                        //Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Identificacion, "", "COD_EMPL", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);


                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Identificacion, "D", c, ErrorValidacion);

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
        public void KBiGruimCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosBI.ModulosBi_8.KBiGruimCheckList")
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
                            //int velocidad = 10;

                            //Playback.PlaybackSettings.DelayBetweenActions = velocidad;

                            //@User @Motor @CodProgram @NomProgram @TipoQbe @QbeFilter @banExcel @BanderaLupa @BanderaSum @BanderaPreli
                            //@Identificacion @Concepto @Ausencia @Observaciones

                            if (
                                //Datos Login
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&

                                rows["Grupo"].ToString().Length != 0 && rows["Grupo"].ToString() != null &&
                                rows["GrupoDescrip"].ToString().Length != 0 && rows["GrupoDescrip"].ToString() != null &&

                                rows["Implemento"].ToString().Length != 0 && rows["Implemento"].ToString() != null &&
                                rows["ImplementoEditar"].ToString().Length != 0 && rows["ImplementoEditar"].ToString() != null
                                )
                            {
                                //VARIABLES GENERALES
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
                                //VARIABLES CRUD                                
                                string Grupo = rows["Grupo"].ToString();
                                string GrupoDescrip = rows["GrupoDescrip"].ToString();
                                string GrupoEdit = rows["GrupoEdit"].ToString();
                                //VARIABLES DETALLE 1                                
                                string Implemento = rows["Implemento"].ToString();
                                string ImplementoEditar = rows["ImplementoEditar"].ToString();

                                List<string> crudVariables = new List<string>() { Grupo, GrupoDescrip, GrupoEdit };
                                List<string> crudVariablesDetalle1 = new List<string>() { Implemento, ImplementoEditar };
                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                string file = CrearDocumentoWordDinamico(codProgram, "BI", motor);

                                try
                                {

                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();
                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
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

                                        //ELIMINAR REGISTROS PEGADOS DETALLE
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera("BI_DGRUI", user, motor, Campo, Grupo, Grupo, "COD_GRUP", c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //ELIMINAR REGISTROS PEGADOS
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Grupo, Grupo, "COD_GRUP", c, ErrorValidacion, "No se agregó correctamente", 0, "1");

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        ////////AGREGAR REGISTRO
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbigruim.CrudPrincipal(desktopSession, 0, crudVariables);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

                                        ////DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);
                                        Thread.Sleep(1500);
                                        CrudKbigruim.CrudPrincipal(desktopSession, 1, crudVariables);
                                        newFunctions_4.ScreenshotNuevo("Editar Datos Crud Principal", true, file);
                                        Thread.Sleep(1500);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados", true, file);

                                        Thread.Sleep(2000);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(2000);
                                        newFunctions_4.ScreenshotNuevo("Editar", true, file);
                                        CrudKbigruim.CrudPrincipal(desktopSession, 1, crudVariables);
                                        newFunctions_4.ScreenshotNuevo("Editar Datos Crud Principal", true, file);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(3000);
                                        newFunctions_4.ScreenshotNuevo("Edición Aplicada", true, file);


                                        ////Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Grupo, Grupo, "COD_GRUP", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        ///////////BARRA DETALLE 1//////
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                                        ElementList = PruebaCRUD.NavClass(desktopSession);

                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbigruim.CrudDetalle(desktopSession, 0, crudVariablesDetalle1, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 1", true, file);

                                        //DESCARTAR EDICION DETALLE 1
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbigruim.CrudDetalle(desktopSession, 1, crudVariablesDetalle1, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar Detalle 1", true, file);

                                        //ACEPTAR EDICION DETALLE 1
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbigruim.CrudDetalle(desktopSession, 1, crudVariablesDetalle1, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar Detalle 1", true, file);

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
                                        CrudKbigruim.PreliKBiGruim(desktopSession, file, pdf1, TipoQbe, QbeFilter, BanderaPreli);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        Thread.Sleep(1000);
                                        CrudKbigruim.Exceldistinto(desktopSession, 0, file);
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);
                                        CrudKbigruim.Exceldistinto(desktopSession, 1, file);
                                        string ruta2 = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta2);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //RECTIFICACION DE CAMPOS DE EXCEL       
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, Grupo, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //ELIMINAR DETALLE 1
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador, file);
                                        Thread.Sleep(1000);

                                        //ELIMINAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(500);
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        Thread.Sleep(1000);

                                        //Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Grupo, "", "COD_GRUP", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);

                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Grupo, "D", c, ErrorValidacion);

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
        public void KBiPaendCheckList()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosBI.ModulosBi_8.KBiPaendCheckList")
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
                            //int velocidad = 10;

                            //Playback.PlaybackSettings.DelayBetweenActions = velocidad;

                            //@User @Motor @CodProgram @NomProgram @TipoQbe @QbeFilter @banExcel @BanderaLupa @BanderaSum @BanderaPreli
                            //@Identificacion @Concepto @Ausencia @Observaciones

                            if (
                                //Datos Login
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["NomProgram"].ToString().Length != 0 && rows["NomProgram"].ToString() != null &&
                                //Datos Necesidades Formación    
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null &&

                                rows["Prototipo"].ToString().Length != 0 && rows["Prototipo"].ToString() != null &&
                                rows["TopeMaxSuel"].ToString().Length != 0 && rows["TopeMaxSuel"].ToString() != null &&
                                rows["DiasMinTrab"].ToString().Length != 0 && rows["DiasMinTrab"].ToString() != null &&
                                rows["DiasMinTrabEditar"].ToString().Length != 0 && rows["DiasMinTrabEditar"].ToString() != null
                                )
                            {
                                //VARIABLES GENERALES
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
                                //VARIABLES CRUD                                
                                string Prototipo = rows["Prototipo"].ToString();
                                string TopeMaxSuel = rows["TopeMaxSuel"].ToString();
                                string DiasMinTrab = rows["DiasMinTrab"].ToString();
                                string DiasMinTrabEditar = rows["DiasMinTrabEditar"].ToString();

                                List<string> crudVariables = new List<string>() { Prototipo, TopeMaxSuel, DiasMinTrab, DiasMinTrabEditar };
                                List<string> crudVariablesDetalle1 = new List<string>() { };

                                if (motor == "SQL")
                                {
                                    string TipoOp = rows["TipoOp"].ToString();
                                    string TipoOpera = rows["TipoOpera"].ToString();
                                    string NDocumento = rows["NDocumento"].ToString();
                                    string Bodega = rows["Bodega"].ToString();

                                    crudVariables = new List<string>() { Prototipo, TopeMaxSuel, DiasMinTrab, TipoOp, TipoOpera, NDocumento, Bodega, DiasMinTrabEditar };


                                    string Codigo1 = rows["Codigo1"].ToString();
                                    crudVariablesDetalle1 = new List<string>() { Codigo1 };
                                }

                                string codigo2 = rows["codigo2"].ToString();
                                string codigo2Editar = rows["codigo2Editar"].ToString();
                                List<string> crudVariablesDetalle1_2 = new List<string>() { codigo2, codigo2Editar };

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                string file = CrearDocumentoWordDinamico(codProgram, "GH", motor);

                                try
                                {

                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();
                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
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

                                        ////QBE
                                        //errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //newFunctions_1.lapiz(desktopSession);

                                        ////////Identificar Botoneras
                                        Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
                                        botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
                                        var ElementList = PruebaCRUD.NavClass(desktopSession);

                                       


                                        //Validacion Eliminar Registro
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Prototipo, "", "COD_PROT", c, ErrorValidacion, "los datos no se encontraron correctamente", 2);


                                        //CONSULTA DE DELETE DE AUDITORIA GN_LOGBO CABECERA
                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Prototipo, "D", c, ErrorValidacion);

                                        //////////AGREGAR REGISTRO
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbipaend.CRUDKBiPaend(desktopSession, 1, crudVariables, file, motor);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);

                                        //Validacion Agregar
                                        ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Prototipo, Prototipo, "COD_PROT", c, ErrorValidacion, "No se agregó correctamente", 0);

                                        //DESCARTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbipaend.CRUDKBiPaend(desktopSession, 2, crudVariables, file, motor);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        //Validacion Editar Descartar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Prototipo, DiasMinTrab, "VAL_CINA", c, ErrorValidacion, "No se edito correctamente", 1);

                                        //ACEPTAR EDICION
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbipaend.CRUDKBiPaend(desktopSession, 2, crudVariables, file, motor);
                                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);

                                        ErrorValidacion = ValidacionesQueries.ValidacionLogbo(Tabla, motor, user, Campo, Prototipo, "E", c, ErrorValidacion);

                                        //Validacion Editar Aceptar
                                        //ErrorValidacion = ValidacionesQueries.ValidacionCabecera(Tabla, user, motor, Campo, Prototipo, DiasMinTrabEditar, "VAL_CINA", c, ErrorValidacion, "El registro no se editó correctamente", 1);

                                        if (motor == "SQL")
                                        {
                                            ///////////BARRA DETALLE 1//////
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            CrudKbipaend.CRUDDetalle1KBiPaend(desktopSession, 1, crudVariablesDetalle1, file);
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(200);
                                            newFunctions_4.ScreenshotNuevo("Agregar Detalle 1", true, file);
                                        }

                                        ///////////BARRA DETALLE 1_2//////
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbipaend.CRUDDetalle1KBiPaend(desktopSession, 1, crudVariablesDetalle1_2, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(200);
                                        newFunctions_4.ScreenshotNuevo("Agregar Detalle 1", true, file);

                                        //DESCARTAR EDICION DETALLE 1_2
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbipaend.CRUDDetalle1KBiPaend(desktopSession, 2, crudVariablesDetalle1_2, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Descartar", true, file);

                                        //ACEPTAR EDICION DETALLE 1_2
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        CrudKbipaend.CRUDDetalle1KBiPaend(desktopSession, 2, crudVariablesDetalle1_2, file);
                                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                                        desktopSession.Mouse.Click(null);
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("Editar Aceptar", true, file);                                        

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
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
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //RECTIFICACION DE CAMPOS DE EXCEL       
                                        string COD_EMPRAUX = ValidacionesQueries.CodEmpresa(Tabla, user, motor, Campo, Prototipo, c);
                                        string nom_empr = ValidacionesQueries.NombreEmpresa(user, motor, COD_EMPRAUX);
                                        Celda = Celda_Excel(ruta, nom_empr);
                                        Celda.ForEach(u => ErrorValidacion.Add(u));
                                        newFunctions_1.lapiz(desktopSession);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        newFunctions_1.lapiz(desktopSession);

                                        if (motor == "SQL")
                                        {
                                            //ELIMINAR DETALLE 1
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador, file);
                                            Thread.Sleep(1000);

                                            //ELIMINAR DETALLE 1_2
                                            WindowsDriver<WindowsElement> rootSession = null;
                                            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
                                            desktopSession.Mouse.Click(null);
                                            Thread.Sleep(500);
                                            newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
                                            rootSession = PruebaCRUD.RootSession();
                                            rootSession = PruebaCRUD.ReloadSession(rootSession, "TMessageForm");
                                            var ok = rootSession.FindElementsByName("OK");
                                            ok[0].Click();
                                            Thread.Sleep(1000);
                                            newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
                                            Thread.Sleep(2000);
                                        }

                                        //ELIMINAR REGISTRO
                                        PruebaCRUD.EliminarRegistro(desktopSession, ElementList[0], botonesNavegador, file);
                                        Thread.Sleep(1000);

                                        if (motor == "ORA")
                                        {
                                            //ELIMINAR DETALLE 1
                                            PruebaCRUD.EliminarRegistro(desktopSession, ElementList[1], botonesNavegador, file);
                                            Thread.Sleep(1000);
                                        }
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

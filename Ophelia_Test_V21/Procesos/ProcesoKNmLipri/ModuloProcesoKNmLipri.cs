using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales;
using System.Text;
using System.Threading;
using System.Drawing;
using System.Diagnostics;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesoKNmLivac;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.ValidacionQueries;
using OpenQA.Selenium.Interactions;
using System.Data;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesoKNmLices;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesGlobalesProcesos;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesosKNmLinoe;

namespace Ophelia_Test_V21.Procesos.ProcesoKNmLipri
{
    /// <summary>
    /// Descripción resumida de UnitTest1
    /// </summary>
    [TestClass]
    public class ModuloProcesoKNmLipri : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> session;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();

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

        public ModuloProcesoKNmLipri()
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
        public void ProcesoKNmLipri()
        {
            //////ARRANCAR EN SQL
            string motor = "SQL";
            var prog = "KNmLipri";
            //Debugger.Launch();
            AbrirPrograma a = new AbrirPrograma(prog, "KOBA");

            desktopSession = a.Start2(motor, "");

            var TipoPrima= desktopSession.FindElementsByName("Servicios (Primer Semestre)");//"  Selección de Empleados  " "Primas Extralegales" "Servicios (Primer Semestre)"
            TipoPrima[0].Click();
            //Thread.Sleep(10000);
            var TipoPrima3 = desktopSession.FindElementsByName("Navidad (Segundo Semestre)");//"Servicios (Primer Semestre)" "Primas Extralegales" "Servicios (Primer Semestre)"
            TipoPrima3[0].Click();
            ////Thread.Sleep(10000);
            ////var TipoPrima2 = desktopSession.FindElementsByName("Primas Extralegales");//"  Selección de Empleados  " "Primas Extralegales" "Servicios (Primer Semestre)"
            ////TipoPrima2[0].Click();
            


            var CamposTEdit = desktopSession.FindElementsByClassName("TEdit");
            CamposTEdit[0].SendKeys("16/12/2021");//Fecha de Pago
            CamposTEdit[1].SendKeys("16/12/2021");//Fecha de Liquidacion
            CamposTEdit[2].SendKeys("1042");//Prototipo

            CamposTEdit[3].SendKeys("16/12/2021");//Fecha de Pago
            CamposTEdit[4].SendKeys("16/12/2021");//Fecha de Liquidacion
            Thread.Sleep(10000);

            var Sim = desktopSession.FindElementsByName("Simulación");
            Sim[0].Click();
            var LyI = desktopSession.FindElementsByName("Liquidar/Ingreso");
            LyI[0].Click();

            //Normal
            var Normal = desktopSession.FindElementsByName("Normal");
            Normal[0].Click();
            // Consolidada Reliquidación Consolidar Todo

            //int cont = 0;
            //for (int i = 0; i < 20; i++)
            //{
            //    Elementlist[i].SendKeys(i.ToString());
            //    cont = cont + 1;
            //}

            //string file = CrearDocumentoWordDinamico(prog, "PruebasCrud", motor);
            //List<string> crudVariables = new List<string> { "1", "Activo", "Inactivo" };

            //ARRANCAR EN ORA
            //string motor = "ORA";
            //var prog = "KNmCaban";
            //AbrirPrograma a = new AbrirPrograma(prog, "ODESAR");
            //string file = CrearDocumentoWordDinamico(prog, "PruebasCrud", motor);
            //List<string> crudVariables = new List<string> { "123", "BAN", "BBVA", "1" };
            //List<string> VariablesLupa = new List<string> { "BAN" };
            //List<string> crudVariables1 = new List<string> { "TELIZAB", "KADMCON" };
            //List<string> crudVariables2 = new List<string> { "1", "VF.", "FP." };
            ////EJECUTAR EN EL NAVEGADOR

            //fecha emision TEdit ==3
            //fechas final TEdit =4
            //fecha inicial TEdit = 5
            // TRM TEdit =1
            //CUNE Tedit = 0
            //TMemo = 0 observaciones TGroupBox
            // PREFIJO TEdit =2

            ////////AGREGAR REGISTRO
            //var ElementList = PruebaCRUD.NavClass(desktopSession);
            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 125, 15);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(1000);
            ////Thread.Sleep(50000);
            //var Elementlist1 = desktopSession.FindElementsByName("Servicios (Primer Semestre)");//TEdit
            //var TEdit = desktopSession.FindElementsByClassName("TEdit");//TEdit
            //Elementlist1[2].SendKeys(Convert.ToString("12584"));//identificacion
            ////Elementlist1[3].SendKeys(Convert.ToString("3"));
            //Elementlist1[4].SendKeys(Convert.ToString("12584"));//Codigo interno
            ////Elementlist1[5].SendKeys(Convert.ToString("5"));

            //TEdit[0].SendKeys(Convert.ToString("TEdit0"));//segundo nombre
            //TEdit[1].SendKeys(Convert.ToString("TEdit1"));//segundo apellido
            //TEdit[2].SendKeys(Convert.ToString("TEdit2"));//primer nombre
            //TEdit[3].SendKeys(Convert.ToString("TEdit3"));//primer apellido

            //Elementlist1[7].SendKeys(Convert.ToString("C")); //Tipo de documento
            ////Elementlist1[8].SendKeys(Convert.ToString("8"));
            ////Elementlist1[9].SendKeys(Convert.ToString("9"));

            //var Sexo = desktopSession.FindElementsByName("Masculino");
            //Sexo[0].Click();

            //var Nac = desktopSession.FindElementsByName("Nacional");
            //Nac[0].Click();

            ////Elementlist1[10].SendKeys(Convert.ToString("10"));//pais extrajeros

            ////Elementlist1[11].SendKeys(Convert.ToString("11"));
            //Elementlist1[12].SendKeys(Convert.ToString("12/04/1994"));//nacimiento
            ////Elementlist1[13].SendKeys(Convert.ToString("13"));//distrito
            ////Elementlist1[14].SendKeys(Convert.ToString("14"));//numero
            //Elementlist1[15].SendKeys(Convert.ToString("12/04/2014"));//fecha expedicion
            ////Elementlist1[16].SendKeys(Convert.ToString("16"));

            //Elementlist1[23].SendKeys(Convert.ToString("169"));//pais
            //Elementlist1[22].SendKeys(Convert.ToString("11"));//departamento
            //Elementlist1[21].SendKeys(Convert.ToString("1"));//municipio
            //Elementlist1[17].SendKeys(Convert.ToString("1"));//localidad
            ////Elementlist1[18].SendKeys(Convert.ToString("18"));
            ////Elementlist1[19].SendKeys(Convert.ToString("19"));
            ////Elementlist1[20].SendKeys(Convert.ToString("20"));


            ////Elementlist1[24].SendKeys(Convert.ToString("24"));
            ////Elementlist1[25].SendKeys(Convert.ToString("25"));

            ////Datos Personales Pesta;a
            //var PestaDatos = desktopSession.FindElementsByName("Datos Personales ");
            //PestaDatos[0].Click();

            //var ElementlistPestana = desktopSession.FindElementsByClassName("TDBEdit");//TDBEdit
            //ElementlistPestana[23].SendKeys(Convert.ToString("169"));//pais lugar de residencia
            //ElementlistPestana[22].SendKeys(Convert.ToString("11"));//departamento lugar de residencia
            //ElementlistPestana[21].SendKeys(Convert.ToString("1"));// muynicipio lugar de residenci
            //ElementlistPestana[17].SendKeys(Convert.ToString("1"));// localidad

            //ElementlistPestana[13].SendKeys(Convert.ToString("1"));// municipio lugar de nacimiento 
            //ElementlistPestana[14].SendKeys(Convert.ToString("11"));// departamento lusgar de nacimiento
            //ElementlistPestana[15].SendKeys(Convert.ToString("169"));// lugar de nacimiento pais
            //ElementlistPestana[9].SendKeys(Convert.ToString("1"));// localidad


            ////ElementlistPestana[27].SendKeys(Convert.ToString("BG"));//al lado de barrio
            ////ElementlistPestana[28].SendKeys(Convert.ToString("BA"));//barrio
            ////ElementlistPestana[24].SendKeys(Convert.ToString("1"));//NumCasa
            //ElementlistPestana[25].Click();//click en dir
            ////ENONTRAR EN VENTANA EMERGENTE TEdit

            //WindowsDriver<WindowsElement> rootSession = null;
            //Thread.Sleep(1000);
            //rootSession = PruebaCRUD.RootSession();
            //rootSession = PruebaCRUD.ReloadSession(rootSession, "TFrmDireccion");
            //var allFrame = rootSession.FindElementsByClassName("TFrmDireccion");
            //Thread.Sleep(1000);

            //var TEdit2 = rootSession.FindElementsByClassName("TEdit");
            //TEdit2[3].SendKeys(Convert.ToString("TEdit0"));
            //TEdit2[4].SendKeys(Convert.ToString("TEdit4"));

            ////CLICK EN LE BOTON ACEPTAR TBitBtn
            ////var BOTTONaCEPTAR = rootSession.FindElementsByClassName("TBitBtn");
            //new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) - 100, (allFrame[0].Size.Height / 2) + 100).DoubleClick().Perform();
            ////TEdit2[5].SendKeys(Convert.ToString("TEdit5"));
            ////TEdit2[6].SendKeys(Convert.ToString("TEdit6"));
            ////TEdit2[7].SendKeys(Convert.ToString("TEdit7"));
            ////TEdit[4].SendKeys(Convert.ToString("TEdit4"));
            ////int cont = 0;

            ////var Elementlist = desktopSession.FindElementsByClassName("TDBEdit");
            ////int cont = 0;
            ////for (int i = 0; i < 20; i++ )
            ////{
            ////    Elementlist[i].SendKeys(i.ToString());
            ////    cont = cont + 1;
            ////}
            ////var CajaObserv = desktopSession.FindElementsByClassName("TMemo");
            //// Vacaciones e Incapacidades

            ////List<string> errors = new List<string>();
            ////string file = CrearDocumentoWordDinamico("Noema", "NM", motor);
            //////2986749 // 1985163 // 1233907670 // 1233908779 // 1014226548 // 6106394 // 5975854
            ////errors = FuncionesGlobales.QBEQry(desktopSession, "0", "1233907670", file);
            //Thread.Sleep(50000);
            ///////////////////////////////////////////////////////////////


            ////var Empleado = desktopSession.FindElementsByName("Empleado");
            ////Empleado[0].Click();
            ////Thread.Sleep(500);
            ////// TDBEdit
            ////var CamposEmpleados = desktopSession.FindElementsByClassName("TDBEdit");
            ////Debugger.Launch();
            ////for (int i = 1; i < 10; i++) { string j = CamposEmpleados[i].Text; }


            ///////////////////////////////////////////////////////////////

            ////var Ingresos = desktopSession.FindElementsByName("Ingresos");
            ////Ingresos[0].Click();

            ////var tabVac = desktopSession.FindElementsByName("Vacaciones e Incapacidades");
            ////tabVac[0].Click();

            ////WindowsDriver<WindowsElement> rootSession = null;
            ////Thread.Sleep(100);
            ////rootSession = PruebaCRUD.RootSession();
            ////rootSession = PruebaCRUD.ReloadSession(rootSession, "TDBGrid");

            //////var CajaVacaciones = desktopSession.FindElementsByClassName("TGroupBox");
            ////var CajaVacaciones = desktopSession.FindElementsByClassName("TDBGrid");
            ////CajaVacaciones[0].Click();

            //////CajaVacaciones1[0].Click();

            ////List<string> var3 = new List<string>();

            //////desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowUp);
            ////desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowLeft);
            ////desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowLeft);
            ////desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowLeft);
            ////desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowLeft);
            ////desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowLeft);
            ////desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            ////var celdavaciones = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
            //////string var4 = celdavaciones[1].Text;

            ////var3.Add(celdavaciones[1].Text);

            ////for (int i = 0; i < 5; i++)
            ////{
            ////    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            ////    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            ////    var3.Add(celdavaciones[1].Text);
            ////}
            ////Debugger.Launch();


            ////var CajaVacaciones = desktopSession.FindElementsByClassName("TDBGrid");
            ////CajaVacaciones[0].Click();



            ////var celdavaciones = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
            ////celdavaciones[0].Click();
            ////TDBGridInplaceEdit
            ////var CajasDetalles = desktopSession.FindElementByXPath(" / Pane[@ClassName =\"#32769\"][@Name=\"Escritorio 1\"]/Window[@ClassName=\"IEFrame\"][@Name=\"KACTUS HR - KNMNOEMA - Internet Explorer\"]/Pane[@ClassName=\"Frame Tab\"]/Pane[@ClassName=\"TabWindowClass\"][@Name=\"KACTUS HR - KNMNOEMA - Internet Explorer\"]/Pane[@ClassName=\"Shell DocObject View\"]/Pane[@ClassName=\"Internet Explorer_Server\"][starts-with(@Name,\"http://opht1:8081/cgi/sgngener.dll/opcion2?KNMNOEMA&amp;/GARH/NM/BIN\")]/Pane[@ClassName=\"Shell Embedding\"]/Pane[@ClassName=\"Shell DocObject View\"]/Pane[@ClassName=\"DocObject_Top_Class\"]/Window[@ClassName=\"RootBrowserWindow\"]/Pane[@ClassName=\"TFrmNmNoema\"][@Name=\"FrmNmNoema - Maestro Nomina Electronica\"]/Pane[@ClassName=\"TPanel\"]/Pane[@ClassName=\"TPanel\"]/Pane[@ClassName=\"TPanel\"]/Pane[@ClassName=\"TScrollBox\"]/Tab[@ClassName=\"TPageControl\"]/Pane[@ClassName=\"TTabSheet\"][@Name=\"Ingresos\"]/Tab[@ClassName=\"TPageControl\"]/Pane[@ClassName=\"TTabSheet\"][@Name=\"Vacaciones e Incapacidades\"]/Pane[@ClassName=\"TGroupBox\"][@Name=\" Vacaciones \"]/Pane[@ClassName=\"TDBGrid\"]");
            ////
            ////CajaVacaciones[0].Click();
            ////Debugger.Launch();
            ////CajasDetalles.Click();
            ////WindowsDriver<Window sElement> rootSession = null;
            ////Thread.Sleep(100);
            ////rootSession = Prueba  CRUD.RootSession();
            ////rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
            ////var allFrame = rootSession.FindElementsByClassName("TDBGrid");
            ////Thread.Sleep(100);
            //////Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 5), (allFrame[0].Size.Height / 5)).Click().Perform();
            ////var CajaObserv = allFrame[0].FindElementsByClassName("TDBGridInplaceEdit");
            ////CajaObserv[0].Click();
            ////string Identificacion = celdavaciones[0].Text;

            ////Thread.Sleep(50000);

            ////int cont = 0;
            ////CajaObserv[0].SendKeys(Convert.ToString(cont));
            ////Elementlist[2].SendKeys(Convert.ToString(cont));
            ////Cajaprefijo[2].SendKeys("KOBE");
            ////foreach (var item in Cajaprefijo)
            ////{
            ////    if (cont == 3 || cont == 4 || cont == 5)
            ////    {
            ////        item.SendKeys("01/01/2021");
            ////    }
            ////    else { item.SendKeys(Convert.ToString(cont)); }

            ////    cont = cont + 1;

            ////}

            ////Thread.Sleep(20000);
            ////////////AGREGAR REGISTRO
            ////Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
            ////botonesNavegador = PruebaCRUD.findname(desktopSession, 27, 0);
            ////var ElementList = PruebaCRUD.NavClass(desktopSession);

            //////////INSERTAR
            ////desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
            ////desktopSession.Mouse.Click(null);


            ////CrudKnmcaban.CRUD(desktopSession, "0", crudVariables, VariablesLupa, file);
            ////CrudKnmcaban.CRUD(desktopSession, "1", crudVariables, VariablesLupa, file);

            //////////////APLICAR
            ////desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
            ////desktopSession.Mouse.Click(null);


            ////CrudKnmcaban.CRUD(desktopSession, "2", crudVariables, VariablesLupa, file);
            //////Thread.Sleep(200000);


            ///////CRUD DETALLES 
            ////////////AGREGAR REGISTRO
            ////Dictionary<string, Point> botonesNavegador1 = new Dictionary<string, Point>();
            ////botonesNavegador1 = PruebaCRUD.findname(desktopSession, 27, 0);
            ////var ElementList1 = PruebaCRUD.NavClass(desktopSession);
            ///////CRUD DETALLES 1
            ///////INSERTAR
            ////desktopSession.Mouse.MouseMove(ElementList1[2].Coordinates, botonesNavegador1["Insertar"].X, botonesNavegador1["Insertar"].Y);
            ////desktopSession.Mouse.Click(null);

            ////CrudKnmcaban.CRUDdetelles1(desktopSession, "0", crudVariables1, file);
            //////Aplicar
            ////desktopSession.Mouse.MouseMove(ElementList1[2].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
            ////desktopSession.Mouse.Click(null);
            //////Thread.Sleep(200000);

            ////////////EDITAR CAMBIOS (EDITAR - CANCELAR)
            ////desktopSession.Mouse.MouseMove(ElementList1[2].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
            ////desktopSession.Mouse.Click(null);

            ////CrudKnmcaban.CRUDdetelles1(desktopSession, "1", crudVariables1, file);
            //////Thread.Sleep(200000);

            ////////////CANCELAR EDITAR
            ////desktopSession.Mouse.MouseMove(ElementList1[2].Coordinates, botonesNavegador1["Cancelar"].X, botonesNavegador1["Cancelar"].Y);
            ////desktopSession.Mouse.Click(null);
            //////Thread.Sleep(200000);
            ////////LUPA DINAMICA
            //////newFunctions_3.LupaAud(desktopSession, "0", file);

            //////////EDITAR CAMBIOS (EDITAR - APLICAR)
            ////desktopSession.Mouse.MouseMove(ElementList1[2].Coordinates, botonesNavegador1["Editar"].X, botonesNavegador1["Editar"].Y);
            ////desktopSession.Mouse.Click(null);
            ////CrudKnmcaban.CRUDdetelles1(desktopSession, "1", crudVariables1, file);

            //////////////APLICAR EDITAR
            ////desktopSession.Mouse.MouseMove(ElementList1[2].Coordinates, botonesNavegador1["Aplicar"].X, botonesNavegador1["Aplicar"].Y);
            ////desktopSession.Mouse.Click(null);


            ///////CRUD DETALLES 2
            ////CrudKnmcaban.CRUDdetelles2(desktopSession, "0", crudVariables2, file);

            ////desktopSession.Mouse.MouseMove(ElementList1[1].Coordinates, botonesNavegador1["Insertar"].X, botonesNavegador1["Insertar"].Y);
            ////desktopSession.Mouse.Click(null);

            ////CrudKnmcaban.CRUDdetelles2(desktopSession, "1", crudVariables2, file);

            ////Thread.Sleep(200000);
            //////TDBGridInplaceEdit
            /////////// ELIMINAR REGISTRO
            ////////newFunctions_3.LupaAud(desktopSession, "0", file);
            ////PruebaCRUD.EliminarRegistro(desktopSession, ElementList1[2], botonesNavegador1, file);

        }

        [TestMethod]
        public void ValidacionesProcesoKNmLipri()
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            List<string> NomCamposValidarProcesoP = new List<string>();
            List<string> VarValorGlobalProcesoP = new List<string>();
            List<string> NumCamposTabProcesoP = new List<string>();
            string motor = string.Empty;
            string codProgram = string.Empty;
            int cont = 0;
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
            //TableOrder = "DW_A1705"; 
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.Procesos.ProcesoKNmLipri.ModulosProcesoKNmLipri.ValidacionesProcesoKNmLipri")
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

                        string Carpeta = "Proceso_KNmLipri";

                        foreach (DataRow rows in DataCase.Tables[0].Rows)
                        {

                            if (
                                //Datos Login
                                rows["User"].ToString().Length != 0 && rows["User"].ToString() != null &&
                                rows["Motor"].ToString().Length != 0 && rows["Motor"].ToString() != null &&
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null &&
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["QbeFilter1"].ToString().Length != 0 && rows["QbeFilter1"].ToString() != null &&
                                //rows["QbeFilter2"].ToString().Length != 0 && rows["QbeFilter2"].ToString() != null &&
                                rows["IdInicial"].ToString().Length != 0 && rows["IdInicial"].ToString() != null &&
                                rows["IdFinal"].ToString().Length != 0 && rows["IdFinal"].ToString() != null &&
                                //rows["CampoID"].ToString().Length != 0 && rows["CampoID"].ToString() != null &&
                                rows["CampoTN"].ToString().Length != 0 && rows["CampoTN"].ToString() != null &&
                                rows["CampoFC"].ToString().Length != 0 && rows["CampoFC"].ToString() != null
                                )
                            {

                                string user = rows["User"].ToString();
                                motor = rows["Motor"].ToString();
                                codProgram = rows["CodProgram"].ToString();
                                //string nomProgram = rows["NomProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();   //Valor del filtro QBE Identificacion
                                string QbeFilter1 = rows["QbeFilter1"].ToString(); //Valor del filtro QBE tipo de Nomina
                                string QbeFilter2 = rows["QbeFilter2"].ToString(); //Valor del filtro QBE Fecha Acumula
                                string IdInicial = rows["IdInicial"].ToString(); // ID por la que inicia la validacioon
                                string IdFinal = rows["IdFinal"].ToString();// ID por la que termina la validacioon
                                //string CampoID = rows["CampoID"].ToString();// Posicion del campo ID en QBE
                                string CampoTN = rows["CampoTN"].ToString();// Posicion del campo Tipo de Nomina en QBE
                                string CampoFC = rows["CampoFC"].ToString();// Posicion del campo Fecha Acumula en QBE

                                // Variables que engloban el proceso 
                                string Identificacion = rows["Identificacion"].ToString();//KNmAcumu // KNmPrima // KNmCopri
                                string Concepto = rows["Concepto"].ToString(); //KNmAcumu 1TB       //   2TB    //    3TB
                                string Cantidad = rows["Cantidad"].ToString();//KNmAcumu 4TB //
                                string Valor = rows["Valor"].ToString();//KNmAcumu 1TB // KNmPrima 1TB
                                string ValorTotal = rows["ValorTotal"].ToString();//KNmAcumu 3TB//                                
                                string ClaseNomina = rows["ClaseNomina"].ToString();//KNmAcumu 15TB//
                                string FechaInicio = rows["FechaInicio"].ToString();//KNmAcumu 2TB//
                                string FechaFinal = rows["FechaFinal"].ToString();//KNmAcumu 1TB//


                                string FechaPago = rows["FechaPago"].ToString();//KNmPrima 2TB
                                                                
                                string NroDiasTrabajados = rows["NroDiasTrabajados"].ToString(); //KNmPrima 5TB // KNmCopri 3TB
                                string NroDiasnoTrabajados = rows["NroDiasnoTrabajados"].ToString();//KNmPrima 1 TB // KNmCopri 1TB
                                string BaseLiq = rows["BaseLiq"].ToString(); //KNmPrima 1TB// KNmCopri 3TB

                     
                                string ValorRetefuente = rows["ValorRetefuente"].ToString();//KNmPrima 1TB
                                                                                                
                                string NroDiasDerecho = rows["NroDiasDerecho"].ToString();// KNmCopri 1D
                                string ConsolidadasAnterior = rows["ConsolidadasAnterior"].ToString();// KNmCopri 1TB
                                string Causadas = rows["Causadas"].ToString();// KNmCopri 1TB
                                string Pagadas = rows["Pagadas"].ToString();// KNmCopri 1TB
                                string Consolidada = rows["Consolidada"].ToString();// KNmCopri 1TB

                                //Variables del campo a validar
                                string CampoValidar1 = rows["CampoValidar1"].ToString();
                                string CampoValidar2 = rows["CampoValidar2"].ToString();
                                string CampoValidar3 = rows["CampoValidar3"].ToString();
                                string CampoValidar4 = rows["CampoValidar4"].ToString();
                                string CampoValidar5 = rows["CampoValidar5"].ToString();
                                string CampoValidar6 = rows["CampoValidar6"].ToString();
                                string CampoValidar7 = rows["CampoValidar7"].ToString();
                                string CampoValidar8 = rows["CampoValidar8"].ToString();
                                string CampoValidar9 = rows["CampoValidar9"].ToString();
                                string CampoValidar10 = rows["CampoValidar10"].ToString();
                                string CampoValidar11 = rows["CampoValidar11"].ToString();

                                //Variables cantidad de Tabs
                                string CampoTab1 = rows["CampoTab1"].ToString();
                                string CampoTab2 = rows["CampoTab2"].ToString();
                                string CampoTab3 = rows["CampoTab3"].ToString();
                                string CampoTab4 = rows["CampoTab4"].ToString();
                                string CampoTab5 = rows["CampoTab5"].ToString();
                                string CampoTab6 = rows["CampoTab6"].ToString();
                                string CampoTab7 = rows["CampoTab7"].ToString();
                                string CampoTab8 = rows["CampoTab8"].ToString();
                                string CampoTab9 = rows["CampoTab9"].ToString();
                                string CampoTab10 = rows["CampoTab10"].ToString();
                                string CampoTab11 = rows["CampoTab11"].ToString();


                                //Lista que contiene todos los campos
                                List<string> VarValorGlobalProceso = new List<string> { Concepto, Cantidad, Valor, ValorTotal, ClaseNomina, FechaInicio, FechaFinal, FechaPago, NroDiasTrabajados, NroDiasnoTrabajados, BaseLiq, ValorRetefuente, NroDiasDerecho, ConsolidadasAnterior, Causadas, Pagadas, Consolidada };

                                //Lista que contiene el nombre de los campos a validar
                                List<string> VarValorGlobalProcesoNom = new List<string> { nameof(Concepto), nameof(Cantidad), nameof(Valor), nameof(ValorTotal), nameof(ClaseNomina), nameof(FechaInicio), nameof(FechaFinal), nameof(FechaPago), nameof(NroDiasTrabajados), nameof(NroDiasnoTrabajados), nameof(BaseLiq), nameof(ValorRetefuente), nameof(NroDiasDerecho), nameof(ConsolidadasAnterior), nameof(Causadas), nameof(Pagadas), nameof(Consolidada) };

                                //Lista que contiene el nombre de los campos a validar
                                List<string> NomCamposValidarProceso = new List<string> { CampoValidar1, CampoValidar2, CampoValidar3, CampoValidar4, CampoValidar5, CampoValidar6, CampoValidar7, CampoValidar8, CampoValidar9, CampoValidar10, CampoValidar11 };

                                //Lista que contiene el número de tabs a pasar en los campos a validar
                                List<string> NumCamposTabProceso = new List<string> { CampoTab1, CampoTab2, CampoTab3, CampoTab4, CampoTab5, CampoTab6, CampoTab7, CampoTab8, CampoTab9, CampoTab10, CampoTab11 };

                                string file = "";
                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                ////Debugger.Launch();
                                if (cont == 0)
                                {
                                    string nomCapture = codProgram + motor + Hora();
                                    file = CrearDocumentoWordDinamico(nomCapture, Carpeta, motor);
                                }

                                try
                                {
                                    if (cont == 0)
                                    {
                                        ErrorValidacion.Clear();
                                        AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                        desktopSession = a.Start2(motor, "");
                                    }

                                    //Thread.Sleep(15000);
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
                                        
                                        //VERSION
                                        //FuncionesGlobales.GetVersion(desktopSession, file);
                                        
                                        ////Debugger.Launch();
                                        Thread.Sleep(500);
                                        if (cont == 0)
                                        {
                                            //KNmAcumu Campo:{,0} KNmPrima Campo:{0,0} KNmCesan:{0,0}
                                            errors = FuncionesKNmLinoe.QBEQry(desktopSession, TipoQbe, QbeFilter2, QbeFilter1,  QbeFilter, file, CampoFC, CampoTN);
                                            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                            Thread.Sleep(1000);

                                        }
                                        Thread.Sleep(20000);

                                        //Seleccionar Data a pasar para el proceso de validación
                                        NomCamposValidarProcesoP = FuncionesProcesosGlobales.ValidacionListaDatosProceso1(desktopSession, VarValorGlobalProceso, NomCamposValidarProceso, VarValorGlobalProcesoNom);
                                        VarValorGlobalProcesoP = FuncionesProcesosGlobales.ValidacionListaDatosProceso2(desktopSession, VarValorGlobalProceso, NomCamposValidarProceso, VarValorGlobalProcesoNom, NomCamposValidarProcesoP);
                                        NumCamposTabProcesoP = FuncionesProcesosGlobales.ValidacionListaDatosTab(desktopSession, NumCamposTabProceso);

                                        //Proceso de validación
                                        Thread.Sleep(2000);
                                        errors = FuncionesProcesosGlobales.ValidacionesProcesos(desktopSession, IdInicial, IdFinal, cont, NomCamposValidarProcesoP, VarValorGlobalProcesoP, NumCamposTabProcesoP, Identificacion);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }


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
                                                //break;
                                            }
                                        }
                                    }
                                    AbrirPrograma a = new AbrirPrograma();
                                    //a.Stop();
                                    //Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    //break;
                                }
                            }
                            else
                            {
                                errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Al validar variables, no tiene registros o no existe");
                            }
                            cont = cont + 1;
                        }
                        if (ErrorValidacion.Count > 0)
                        {
                            ////Debugger.Launch();
                            ReporteErrores(ErrorValidacion, codProgram, motor, Carpeta);
                            var separator = string.Format("{0}{0}", Environment.NewLine);
                            string errorMessageString = string.Join(separator, ErrorValidacion);
                            Assert.Inconclusive(string.Format("Los errores presentados en la prueba {2} son:{0}{1}", Environment.NewLine, errorMessageString, codProgram));
                        }
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
    }
}

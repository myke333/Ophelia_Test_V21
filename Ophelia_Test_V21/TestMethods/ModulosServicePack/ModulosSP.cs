using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data;
using Ophelia_Test_V21.AFuncionesGlobales;
using OpenQA.Selenium.Appium.Windows;
using System.Threading;
using System.Diagnostics;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesRolAuditor;
using System.Runtime.CompilerServices;
using System.IO;
using System.Configuration;
using System.Reflection;
using System.CodeDom;
using Microsoft.CSharp;
using System.CodeDom.Compiler;
using Microsoft.VisualBasic;
using System.Globalization;

namespace Ophelia_Test_V21.TestMethods.ModulosServicePack
{
    /// <summary>
    /// Descripción resumida de ModulosSP
    /// </summary>
    [TestClass]
    public class ModulosSP : FuncionesVitales
    {
        protected static WindowsDriver<WindowsElement> desktopSession;
        List<string> listaResultado;
        List<string> Celda = new List<string>();


        public ModulosSP()
        {
            //
            // TODO: Agregar aquí la lógica del constructor
            //
        }

        /*[TestCleanup]
        public void Limpiar()
        {
            AbrirPrograma a = new AbrirPrograma();
            if (desktopSession != null)
            {
                desktopSession.Close();
                desktopSession.Dispose();
            }
            a.Stop();
        }*/

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
        public void Modulos()
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

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosServicePack.ModulosSP.Modulos")
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
                                rows["CodProgram"].ToString().Length != 0 && rows["CodProgram"].ToString() != null
                                )
                            {

                                string user = rows["User"].ToString();
                                string motor = rows["Motor"].ToString();
                                string codProgram = rows["CodProgram"].ToString();

                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                //Debugger.Launch();
                                string nomCapture = codProgram + motor + Hora();
                                string file = CrearDocumentoWordDinamico(nomCapture, "LIBRERIAS_SP_107", motor);

                                try
                                {

                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();
                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "SP107");
                                    newFunctions_4.ScreenshotNuevo(nomCapture, true, file);
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
                                        //Cuando lee el programa
                                        newFunctions_4.ScreenshotNuevo(nomCapture, true, file);
                                        Thread.Sleep(1000);
                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                    }
                                    a.Stop();
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
                                errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Valide los datos de entrada");
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
        public void SmokeTest()
        {

            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            string motor = string.Empty;
            string codProgram = string.Empty;

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
            //Debugger.Launch();
            foreach (DataRow rowsi in OrderExecutionCase.Tables[0].Rows)
            {
                string plans = rowsi["plans"].ToString();
                string suite = rowsi["suite"].ToString();
                string CaseId = rowsi["CaseId"].ToString();
                string orders = rowsi["orders"].ToString();
                string states = rowsi["states"].ToString();
                string methodname = rowsi["methodname"].ToString();
                string CountDes = rowsi["CountDes"].ToString();

                if (methodname.Replace(" ", string.Empty) == "Ophelia_Test_V21.TestMethods.ModulosServicePack.ModulosSP.SmokeTest")
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
                                rows["TipoQbe"].ToString().Length != 0 && rows["TipoQbe"].ToString() != null &&
                                //rows["QbeFilter"].ToString().Length != 0 && rows["QbeFilter"].ToString() != null &&
                                rows["banExcel"].ToString().Length != 0 && rows["banExcel"].ToString() != null &&
                                rows["BanderaLupa"].ToString().Length != 0 && rows["BanderaLupa"].ToString() != null &&
                                rows["BanderaSum"].ToString().Length != 0 && rows["BanderaSum"].ToString() != null &&
                                rows["BanderaPreli"].ToString().Length != 0 && rows["BanderaPreli"].ToString() != null
                                )
                            {

                                string user = rows["User"].ToString();
                                motor = rows["Motor"].ToString();
                                codProgram = rows["CodProgram"].ToString();
                                string nomProgram = rows["NomProgram"].ToString();
                                string TipoQbe = rows["TipoQbe"].ToString();
                                string QbeFilter = rows["QbeFilter"].ToString();
                                string banExcel = rows["banExcel"].ToString();
                                string BanderaLupa = rows["BanderaLupa"].ToString();
                                string BanderaSum = rows["BanderaSum"].ToString();
                                string BanderaPreli = rows["BanderaPreli"].ToString();


                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                //Debugger.Launch();
                                string nomCapture = codProgram + motor + Hora();
                                string file = CrearDocumentoWordDinamico(nomCapture, "ST_SP_107", motor);

                                try
                                {
                                    ErrorValidacion.Clear();
                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
                                    newFunctions_4.ScreenshotNuevo(nomCapture, true, file);

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
                                        //APERTURA PROGRAMA
                                        newFunctions_4.ScreenshotNuevo("Abrir programa", true, file);
                                        Thread.Sleep(5000);

                                        //VALIDACION CODIGO DEL PROGRAMA
                                        //errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VALIDACION NOMBRE DEL PROGRAMA
                                        //errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //NOTAS
                                        newFunctions_4.openInnerNote(desktopSession, file);
                                        Thread.Sleep(1000);

                                        //LUPA
                                        errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //QBE
                                        errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //SUMATORIA
                                        errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //REPORTE PRELIMINAR
                                        string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        errors = PreliminarGlobal.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1, nomProgram, codProgram, TipoQbe, QbeFilter, motor);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);
                                        //Rectificar Pie Pagina PDF PRELIMINAR
                                        listaResultado = Textopdf(pdf1, codProgram, user);
                                        listaResultado.ForEach(e => ErrorValidacion.Add(e));
                                        Thread.Sleep(1000);

                                        //REPORTE DINAMICO
                                        string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
                                        errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);

                                        //EXPORTAR EXCEL
                                        string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        Thread.Sleep(1000);
                                        a.Stop();


                                    }
                                    else
                                    {
                                        ErrorValidacion.Add($"El programa no puede ser leido : {codProgram}");
                                    }
                                    if (ErrorValidacion.Count > 0)
                                    {
                                        ReporteErrores(ErrorValidacion, codProgram, motor, "ST_SP_107");
                                        var separator = string.Format("{0}{0}", Environment.NewLine);
                                        string errorMessageString = string.Join(separator, ErrorValidacion);
                                        Assert.Inconclusive(string.Format("Los errores presentados en la prueba {2} son:{0}{1}", Environment.NewLine, errorMessageString, codProgram));
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
                                    a.Stop();
                                    //Assert.Fail(CaseId + " ::::::" + e.ToString());
                                    //break;
                                }
                            }
                            else
                            {
                                errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: Valide los datos de entrada");
                            }
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

        [TestMethod]
        public void AbrirPrograma()
        {

            string codProgram = "KNmLinom";
            string user = "UNatalia";
            string motor = "SQL";

            string nomCapture = codProgram + motor + Hora();
            string file = CrearDocumentoWordDinamico(nomCapture, "SP107", motor);

            AbrirPrograma a = new AbrirPrograma(codProgram, user);
            desktopSession = a.Start2(motor, "");

            FuncionesGlobales.GetVersion(desktopSession, file);
            Thread.Sleep(1000);

        }


        private static bool CompileExecutable(string sourceName)
        {
            FileInfo sourceFile = new FileInfo(sourceName);
            CodeDomProvider provider = null;
            bool compileOk = false;

            // Select the code provider based on the input file extension. 
            if (sourceFile.Extension.ToUpper(CultureInfo.InvariantCulture) == ".CS")
            {
                provider = CodeDomProvider.CreateProvider("CSharp");
            }
            else if (sourceFile.Extension.ToUpper(CultureInfo.InvariantCulture) == ".VB")
            {
                provider = CodeDomProvider.CreateProvider("VisualBasic");
            }
            else
            {
                Console.WriteLine("Source file must have a .cs or .vb extension");
            }

            if (provider != null)
            {

                // Format the executable file name. 
                // Build the output assembly path using the current directory 
                // and <source>_cs.exe or <source>_vb.exe.

                String exeName = String.Format(@"{0}\{1}.exe",
                    System.Environment.CurrentDirectory,
                    sourceFile.Name.Replace(".", "_"));

                CompilerParameters cp = new CompilerParameters();

                // Generate an executable instead of  
                // a class library.
                cp.GenerateExecutable = true;

                // Specify the assembly file name to generate.
                cp.OutputAssembly = exeName;

                // Save the assembly as a physical file.
                cp.GenerateInMemory = false;

                // Set whether to treat all warnings as errors.
                cp.TreatWarningsAsErrors = false;

                // Invoke compilation of the source file.
                CompilerResults cr = provider.CompileAssemblyFromFile(cp,
                    sourceName);

                if (cr.Errors.Count > 0)
                {
                    // Display compilation errors.
                    Console.WriteLine("Errors building {0} into {1}",
                        sourceName, cr.PathToAssembly);
                    foreach (CompilerError ce in cr.Errors)
                    {
                        Console.WriteLine("  {0}", ce.ToString());
                        Console.WriteLine();
                    }
                }
                else
                {
                    // Display a successful compilation message.
                    Console.WriteLine("Source {0} built into {1} successfully.",
                        sourceName, cr.PathToAssembly);
                }

                // Return the results of the compilation. 
                if (cr.Errors.Count > 0)
                {
                    compileOk = false;
                }
                else
                {
                    compileOk = true;
                }
            }
            return compileOk;
        }

        [TestMethod]
        public void marco1()
        {
            //Debugger.Launch();
            string[] args = { "TestGraph.cs" };
            if (args.Length > 0)
            {
                //  First parameter is the source file name. 
                if (File.Exists(args[0]))
                {
                    CompileExecutable(args[0]);
                }
                else
                {
                    Console.WriteLine("Input source file not found - {0}",
                        args[0]);
                }
            }
            else
            {
                Console.WriteLine("Input source file not specified on command line!");
            }



            Process.Start("TestGraph_cs.exe");


            //Debugger.Launch();

            //    CodeCompileUnit compileUnit = new CodeCompileUnit();
            //CodeNamespace myNamespace = new CodeNamespace("MyNamespace");
            //myNamespace.Imports.Add(new CodeNamespaceImport("System"));
            //CodeTypeDeclaration myClass = new CodeTypeDeclaration("MyClass");
            //CodeEntryPointMethod start = new CodeEntryPointMethod();
            //CodeMethodInvokeExpression cs1 = new CodeMethodInvokeExpression(
            //new CodeTypeReferenceExpression("Console"),
            //"WriteLine", new CodePrimitiveExpression("Hello World!"));
            //compileUnit.Namespaces.Add(myNamespace);
            //myNamespace.Types.Add(myClass);
            //myClass.Members.Add(start);
            //start.Statements.Add(cs1);


            //CSharpCodeProvider provider = new CSharpCodeProvider();
            //using (StreamWriter sw = new StreamWriter("Hello World!", false))
            //{
            //    IndentedTextWriter tw = new IndentedTextWriter(sw,"");
            //    provider.GenerateCodeFromCompileUnit(compileUnit,
            //    tw, new CodeGeneratorOptions());
            //    tw.Close();
            //}

            //VBCodeProvider provider = new VBCodeProvider();
            //using (StreamWriter sw = new StreamWriter("Hello Jose!", false))
            //{
            //    IndentedTextWriter tw = new IndentedTextWriter(sw, "");
            //    provider.GenerateCodeFromCompileUnit(compileUnit,
            //    tw, new CodeGeneratorOptions());
            //    tw.Close();
            //}


            //VBCodeProvider providor = new VBCodeProvider();
            //using (stre sww = new StreamWriter("Hello Jose!", false))
            //{
            //    IndentedTextWriter tw = new IndentedTextWriter(sw, "");
            //    providor.GenerateCodeFromCompileUnit(compileUnit,
            //    tw, new CodeGeneratorOptions());
            //    tw.Close();
            //}


            //string path = @"C:\Marco1.dll";

            //try
            //{
            //    // Create the file, or overwrite if the file exists.
            //    using (FileStream fs = File.Create(path))
            //    {
            //        byte[] info = new UTF8Encoding(true).GetBytes("This is some text in the file.");
            //        // Add some information to the file.
            //        fs.Write(info, 0, info.Length);
            //    }

            //    // Open the stream and read it back.
            //    using (StreamReader sr = File.OpenText(path))
            //    {
            //        string s = "";
            //        while ((s = sr.ReadLine()) != null)
            //        {
            //            Console.WriteLine(s);
            //        }
            //    }
            //}

            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.ToString());
            //}

        }

        private CodeDomProvider GetCurrentProvider()
        {
            CodeDomProvider provider;
            string var = "CSharp";
            switch (var)
            {
                case "CSharp":
                    provider = CodeDomProvider.CreateProvider("CSharp");
                    break;
                case "Visual Basic":
                    provider = CodeDomProvider.CreateProvider("VisualBasic");
                    break;
                case "JScript":
                    provider = CodeDomProvider.CreateProvider("JScript");
                    break;
                default:
                    provider = CodeDomProvider.CreateProvider("CSharp");
                    break;
            }
            return provider;
        }

        [TestMethod]
        public void GenerateCode()
        {

            CodeDomProvider provider = CodeDomProvider.CreateProvider("CSharp");
            //CodeDomProvider provider = GetCurrentProvider();
            FuncionesGlobales.GenerateCode(provider, FuncionesGlobales.BuildHelloWorldGraph());
            
            // Build the source file name with the appropriate
            // language extension.
            String sourceFile;
            if (provider.FileExtension[0] == '.')
            {
                sourceFile = "TestGraph" + provider.FileExtension;
            }
            else
            {
                sourceFile = "TestGraph." + provider.FileExtension;
            }

            // Read in the generated source file and
            // display the source text.
            StreamReader sr = new StreamReader(sourceFile);
            string leer = sr.ReadToEnd();
            Console.WriteLine(leer);
            sr.Close();
        }

        [TestMethod]
        public void Marco()
        {

            CodeDomProvider provider = GetCurrentProvider();

            // Build the source file name with the appropriate
            // language extension.
            String sourceFile;
            if (provider.FileExtension[0] == '.')
            {
                sourceFile = "TestGraph" + provider.FileExtension;
            }
            else
            {
                sourceFile = "TestGraph." + provider.FileExtension;
            }

            // Compile the source file into an executable output file.
            CompilerResults cr = FuncionesGlobales.CompileCode(provider,
                                                            sourceFile,
                                                            "TestGraph.exe");
            


            //CodeDomProvider provider = GetCurrentProvider();
            //String sourceFile;
            //if (provider.FileExtension[0] == '.')
            //{
            //    sourceFile = "TestGraph" + provider.FileExtension;
            //}
            //else
            //{
            //    sourceFile = "TestGraph." + provider.FileExtension;
            //}

            //// Compile the source file into an executable output file.
            //CompilerResults cr = CodeDomExample.CompileCode(provider,
            //                                                sourceFile,
            //                                                "TestGraph.exe");

            //if (cr.Errors.Count > 0)
            //{
            //    // Display compilation errors.
            //    textBox1.Text = "Errors encountered while building " +
            //                    sourceFile + " into " + cr.PathToAssembly + ": \r\n\n";
            //    foreach (CompilerError ce in cr.Errors)
            //        textBox1.AppendText(ce.ToString() + "\r\n");
            //    run_button.Enabled = false;
            //}
            //else
            //{
            //    textBox1.Text = "Source " + sourceFile + " built into " +
            //                    cr.PathToAssembly + " with no errors.";
            //    run_button.Enabled = true;
            //}



            //Assert.IsTrue(File.Exists("C:\\Marco1.txt"));
            ////Debugger.Launch();

            //string line;
            //try
            //{
            //    //Pass the file path and file name to the StreamReader constructor
            //    StreamReader sr = new StreamReader("C:\\Marco1.txt");                
            //    //Read the first line of text
            //    line = sr.ReadLine();

            //    Assembly a = Assembly.ReflectionOnlyLoadFrom("C:/Marco1.txt");
            //    Type myType = a.GetType(line);
            //    MethodInfo myMethod = myType.GetMethod("metodo");
            //    object obj = Activator.CreateInstance(myType);
            //    myMethod.Invoke(obj, null);

            //    //Continue to read until you reach end of file
            //    while (line != null)
            //    {

            //        //write the lie to console window
            //        Console.WriteLine(line);
            //        //Read the next line
            //        line = sr.ReadLine();


            //        a = Assembly.Load(line);
            //        myType = a.GetType(line);
            //        myMethod = myType.GetMethod("metodo");
            //        obj = Activator.CreateInstance(myType);
            //        myMethod.Invoke(obj, null);

            //    }
            //    //close the file
            //    sr.Close();
            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine("Exception: " + e.Message);
            //}
            //finally
            //{
            //    Console.WriteLine("Executing finally block.");
            //}
        }



        [TestMethod]
        public void GlobalizarMarco()
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
                List<string> Variables_val = new List<string> { };
                if (methodname.Replace(" ", string.Empty) == (this.GetType().FullName + "." + this.testContextInstance.TestName))
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

                            int cont = 0;
                            string[,] Variables = new string[DataCase.Tables[0].Columns.Count, 2];
                            //////////GUARDA EL NOMBRE DE LAS VARIABLES EN UNA LISTA
                            foreach (DataColumn Nomvariable in DataCase.Tables[0].Columns)
                            {
                                Variables[cont, 0] = Nomvariable.ToString();
                                Variables[cont, 1] = rows[Variables[cont, 0]].ToString();
                                cont += 1;
                            }
                            bool Validacion = true;
                            ///////////VALIDACION DE DATOS
                            for (int i = 0; i < DataCase.Tables[0].Columns.Count; i++)
                            {
                                if (Variables[i, 0] != "QbeFilter")
                                {
                                if (!(rows[Variables[i, 0]].ToString().Length != 0 && rows[Variables[i, 0]].ToString() != null))
                                    {
                                        Validacion = false;
                                    }
                                }
                            }
                            /////////////ASIGNACION DE VARIABLES
                            if (Validacion == true)
                            {
                                string user = "", motor = "", codProgram = "", nomProgram = "", TipoQbe = "", QbeFilter = "",
                                 banExcel = "", BanderaLupa = "", BanderaSum = "", BanderaPreli = "";
                                //VARIABLES GENERALES
                                for (int i = 0; i < DataCase.Tables[0].Columns.Count; i++)
                                {
                                    switch (Variables[i, 0])
                                    {
                                        case "User":
                                            user = rows["User"].ToString();
                                            break;
                                        case "Motor":
                                            motor = rows["Motor"].ToString();
                                            break;
                                        case "CodProgram":
                                            codProgram = rows["CodProgram"].ToString();
                                            break;
                                        case "NomProgram":
                                            nomProgram = rows["NomProgram"].ToString();
                                            break;
                                        case "TipoQbe":
                                            TipoQbe = rows["TipoQbe"].ToString();
                                            break;
                                        case "QbeFilter":
                                            QbeFilter = rows["QbeFilter"].ToString();
                                            break;
                                        case "banExcel":
                                            banExcel = rows["banExcel"].ToString();
                                            break;
                                        case "BanderaLupa":
                                            BanderaLupa = rows["BanderaLupa"].ToString();
                                            break;
                                        case "BanderaSum":
                                            BanderaSum = rows["BanderaSum"].ToString();
                                            break;
                                        case "BanderaPreli":
                                            BanderaPreli = rows["BanderaPreli"].ToString();
                                            break;
                                        default:
                                            errorMessages.Add(methodname.Replace(" ", string.Empty) + "::::::" + "MSG: La Variable [" + Variables[i, 0] + "] no tiene registros");
                                            break;
                                    }
                                }
                                //CREAR EL DOCUMENTO DEL REPORTE DE LA PRUEBA
                                //Debugger.Launch();
                                string Nom_Carpeta = "MarcoGlobal";
                                string nomCapture = codProgram + motor + Hora();
                                string file = CrearDocumentoWordDinamico(nomCapture, Nom_Carpeta, motor);
                                try
                                {
                                    List<string> ErrorValidacion = new List<string>();
                                    List<string> errors = new List<string>();
                                    AbrirPrograma a = new AbrirPrograma(codProgram, user);
                                    desktopSession = a.Start2(motor, "");
                                    newFunctions_4.ScreenshotNuevo(nomCapture, true, file);
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
                                        ////////////////////////AQUI VA EL CRUD
                                        //APERTURA PROGRAMA
                                        newFunctions_4.ScreenshotNuevo("Abrir programa", true, file);
                                        Thread.Sleep(5000);
                                        //VALIDACION CODIGO DEL PROGRAMA
                                        errors = FuncionesGlobales.validarCodPrograma(desktopSession, codProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VALIDACION NOMBRE DEL PROGRAMA
                                        errors = FuncionesGlobales.validarDescripPrograma(desktopSession, nomProgram);
                                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        //VERSION
                                        FuncionesGlobales.GetVersion(desktopSession, file);
                                        Thread.Sleep(1000);

                                        ////NOTAS
                                        //newFunctions_4.openInnerNote(desktopSession, file);
                                        //Thread.Sleep(1000);
                                        
                                        ////LUPA
                                        //errors = newFunctions_3.LupaAud(desktopSession, BanderaLupa, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                                        ////QBE
                                        //errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(1000);

                                        ////SUMATORIA
                                        //errors = newFunctions_1.Sumatory(desktopSession, BanderaSum, file);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(1000);

                                        ////REPORTE DINAMICO
                                        //string pdf = @"C:\Reportes\ReportePDF1_" + "Dinamico" + "_" + Hora();
                                        //errors = newFunctions_5.ReporteDinamico(desktopSession, file, pdf);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(1000);

                                        ////REPORTE PRELIMINAR
                                        //string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                                        //errors = PreliminarGlobal.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1, nomProgram, codProgram, TipoQbe, QbeFilter, motor);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(1000);
                                        ////Rectificar Pie Pagina PDF PRELIMINAR
                                        //listaResultado = Textopdf(pdf1, codProgram, user);
                                        //listaResultado.ForEach(e => ErrorValidacion.Add(e));
                                        //Thread.Sleep(1000);

                                        ////EXPORTAR EXCEL
                                        //string ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                        //errors = newFunctions_2.ExpExcel(desktopSession, banExcel, file, codProgram, ruta);
                                        //if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                                        //Thread.Sleep(1000);

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
                                    a.Stop();
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
    }


}


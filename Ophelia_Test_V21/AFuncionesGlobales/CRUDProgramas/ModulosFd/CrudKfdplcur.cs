using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales;
using OpenQA.Selenium;
using System.Drawing;
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System.IO;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosFd
{
    class CrudKfdplcur : FuncionesVitales
    {
        static public WindowsDriver<WindowsElement> RootSession()
        {
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }

        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string parametro)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(parametro);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex

                var appCapabilities = new AppiumOptions();
                //appCapabilities.AddAdditionalCapability("app", "Root");
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }

   
        public static void ClickButtonExterno(WindowsDriver<WindowsElement> desktopSession, int tipo)
        {
            //1 -> Planeación Curso 2 -> Adicional 3 -> Evaluación curso 4 -> Informacion contractual 5-> Sesiones o Talleres 6 -> Cronograma
            List<string> errorMessages = new List<string>();
            if (tipo == 1)
            {
                desktopSession.FindElementByName("Planeación Cursos").Click();
            }
            else if (tipo == 2)
            {
                desktopSession.FindElementByName("Adicional").Click();
            }
            else if (tipo == 3)
            {
                desktopSession.FindElementByName("Evaluación Curso").Click();
            }
            else if (tipo == 4)
            {
                desktopSession.FindElementByName("Información Contractual").Click();
            }
            else if(tipo == 5)
            {
                desktopSession.FindElementByName("Sesiones o Talleres").Click();    
            }
            else if(tipo == 6)
            {
                desktopSession.FindElementByName("Cronograma").Click();
                
            }
        }

        public static void ClickButtonInterno(WindowsDriver<WindowsElement> desktopSession, int tipo)
        {
            //1 -> Definición Sesión 2 -> Logística 3 -> Inscritos  4 -> Documentos 5-> Sesiones o Talleres
            List<string> errorMessages = new List<string>();
            if (tipo == 1)
            {
                desktopSession.FindElementByName("Definición Sesión").Click();
            }
            else if (tipo == 2)
            {
                desktopSession.FindElementByName("Logística").Click();
            }
            else if (tipo == 3)
            {
                desktopSession.FindElementByName("Inscritos").Click();
            }
            else if (tipo == 4)
            {
                desktopSession.FindElementByName("Documentos").Click();
            } 
        }

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit"); 
            var ElementList1 = desktopSession.FindElementsByClassName("TDBComboBox");
            if (tipo == 1)
            {
                LupaDinamica(desktopSession, data);
                ElementList[4].SendKeys(data[7]);
                ElementList[3].SendKeys(data[8]);
                ElementList1[0].SendKeys(data[9]);
                newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);
                // 2 -> Adicional
                ClickButtonExterno(desktopSession, 2);
                var ElementList2 = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList2[2].SendKeys(data[10]);
                var ElementList3 = desktopSession.FindElementsByClassName("TDBMemo");
                ElementList3[0].SendKeys(data[11]);

            }
            else
            {
                ClickButtonExterno(desktopSession, 2);
                var ElementList3 = desktopSession.FindElementsByClassName("TDBMemo");
                ElementList3[0].Clear();
                ElementList3[0].SendKeys(data[12]);
            }
        }

        public static void CrudDetalle1(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 39, 30);
                desktopSession.Mouse.Click(null);

                for (int i = 0; i <= 6; i++)
                {
                    desktopSession.Keyboard.SendKeys(data[i]);
                    if (i < 6)
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        if (i == 4)
                        {
                            for (int j = 0; j < 4; j++)
                            {
                                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            }

                        }
                    } 
                } 
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 1236, 32);
                desktopSession.Mouse.DoubleClick(null);
                Ventanana(desktopSession);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(500);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 1367, 37);
                desktopSession.Mouse.Click(null);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 1367, 37);
                desktopSession.Mouse.DoubleClick(null);
                Ventanana(desktopSession);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 1429, 25);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(data[7]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //newFunctions_4.ScreenshotNuevo("Agregar Registro Detalle 1", true, file);*/
            }
            else
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Clear);
                desktopSession.Keyboard.SendKeys(data[8]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }


        public static void CrudDetalle2(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 28, 27);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //newFunctions_4.ScreenshotNuevo("Agregar Registro Detalle 1", true, file);
            }
        }

        public static void CrudDetalle3(WindowsDriver<WindowsElement> desktopSession, int tipo, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 35, 30);
                desktopSession.Mouse.Click(null);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(2000);
                WindowsDriver<WindowsElement> rootSession = null;
                rootSession = RootSession();
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }

        public static void CrudDetalle4(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                desktopSession.FindElementByName("Externo").Click();
                LupaDinamicaModificacion(desktopSession, data);
                var ElementList1 = desktopSession.FindElementsByClassName("TDBMemo");
                ElementList1[0].SendKeys(data[3]);
                
            }
            else
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(data[4]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }

        public static void CrudDetalle5(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 39, 33);
                desktopSession.Mouse.Click(null);
                for (int i = 0; i < 4; i++)
                {
                    desktopSession.Keyboard.SendKeys(data[i]);
                    if (i < 3)
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        if (i == 1)
                        {
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        }
                    }
                }
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                desktopSession.Keyboard.SendKeys(data[4]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }

        public static void CrudDetalle6(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 508, 30);
                desktopSession.Mouse.Click(null);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(2000);
                WindowsDriver<WindowsElement> rootSession = null;
                rootSession = RootSession();
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowUp);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                for (int i = 0; i < 9; i++)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }
                desktopSession.Keyboard.SendKeys(data[0]);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[1]);
            }
            else
            {
                desktopSession.Keyboard.SendKeys(data[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }

        public static void CrudDetalle7(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 32, 33);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }

        public static void CrudDetalle8(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 39, 33);
                desktopSession.Mouse.Click(null);
                for (int i = 0; i < 3; i++)
                {
                    desktopSession.Keyboard.SendKeys(data[i]);
                    if (i < 2)
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                }
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                desktopSession.Keyboard.SendKeys(data[3]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }

        public static void Ventanana(WindowsDriver<WindowsElement> desktopSession)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        }

        public static void Preliminar(WindowsDriver<WindowsElement> desktopSession, int tipo, string file)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> Session = null;
            Session = RootSession();
            Session = ReloadSession(Session, "TFrmFdManre");
            Thread.Sleep(500);
            if (tipo == 1)
            {
                Session.FindElementByName("Reporte Estándar").Click();
                newFunctions_4.ScreenshotNuevo("Opción Preliminar 1", true, file);
                AceptarPreliminar(Session);
                Thread.Sleep(10000);

            }
            else if (tipo == 2)
            {
                Session.FindElementByName("Reporte Resumido Plan de Capacitación").Click();
                newFunctions_4.ScreenshotNuevo("Opción Preliminar 2", true, file);
                AceptarPreliminar(Session);
                Thread.Sleep(3000);
            }
            else if (tipo == 3)
            {
                Session.FindElementByName("Reporte Detallado").Click();
                newFunctions_4.ScreenshotNuevo("Opción Preliminar 3", true, file);
                AceptarPreliminar(Session);
                Thread.Sleep(60000);
            }
            else if (tipo == 4)
            {
                Session.FindElementByName("Reporte Presupuesto por Curso").Click();
                newFunctions_4.ScreenshotNuevo("Opción Preliminar 4", true, file);
                AceptarPreliminar(Session);
                Thread.Sleep(2000);
            }
            else if (tipo == 5)
            {
                Session.FindElementByName("Reporte Presupuesto por Sesion").Click();
                newFunctions_4.ScreenshotNuevo("Opción Preliminar 5", true, file);
                AceptarPreliminar(Session);
                Thread.Sleep(2000);
            }
            else if (tipo == 6)
            {
                Session.FindElementByName("Exportar Datos de Inscritos a Excel ").Click();
                newFunctions_4.ScreenshotNuevo("Opción Preliminar 6", true, file);
                AceptarPreliminar(Session);
                Thread.Sleep(2000);
            }
            else if (tipo == 7)
            {
                Session.FindElementByName("Reporte de Instructores").Click();
                newFunctions_4.ScreenshotNuevo("Opción Preliminar 7", true, file);
                AceptarPreliminar(Session);
                Thread.Sleep(30000);
            }
            else if (tipo == 8)
            {
                Session.FindElementByName("Reporte Inasistencia").Click();
                newFunctions_4.ScreenshotNuevo("Opción Preliminar 8", true, file);
                AceptarPreliminar(Session);
                Thread.Sleep(4000);
            }

            Thread.Sleep(1000);
        }


        public static void AceptarPreliminar(WindowsDriver<WindowsElement> desktopSession)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> Session = null;
            Session = RootSession();
            Session = ReloadSession(Session, "TFrmFdManre");
            var ElementList = Session.FindElementByName("Reportes");
            Thread.Sleep(500);
            Thread.Sleep(2000);
            Session.Mouse.MouseMove(ElementList.Coordinates, 83, 270);
            Session.Mouse.Click(null);
        }

        public static List<string> VentanaAzul(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;

            bool ventana = newFunctions_1.coordinatesFinder(desktopSession, 27, 101, 206, screenWidth / 2, 347, 1);
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "IEFrame");
            var allFrame = rootSession.FindElementsByClassName("IEFrame");
            newFunctions_4.ScreenshotNuevo("Error Preliminar", true, file);
            if (ventana == false)
            {
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
                rootSession = RootSession();
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);
                errorMessages.Add("Existe error en alguna opción del preliminar");
            }

            return errorMessages;
        }

        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            int indexVal = 0;
            foreach (Point coord in coordenadas)
            {
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                if (contLupa == 1)
                {
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TCaptura");
                    Thread.Sleep(1000);
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    ////Debugger.Launch();
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    campo[1].Clear();
                    rootSession.Keyboard.SendKeys(data[indexVal]);
                    //Aqui se puede enviar un valor

                    var btn = rootSession.FindElementsByClassName("TBitBtn");
                    ////Debugger.Launch();
                    foreach (var boton in btn)
                    {
                        if (boton.Text == "Aceptar")
                        {
                            rootSession.Mouse.MouseMove(boton.Coordinates, 10, 10);
                            rootSession.Mouse.Click(null);
                        }
                    }
                    indexVal++;
                }
            }
        }

        public static void LupaDinamicaModificacion(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            int indexVal = 0;
            foreach (Point coord in coordenadas)
            {
                if (indexVal < 2)
                {
                    desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    if (contLupa == 1)
                    {
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TCaptura");
                        Thread.Sleep(1000);
                        var campo = rootSession.FindElementsByClassName("TEdit");
                        ////Debugger.Launch();
                        rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                        rootSession.Mouse.Click(null);
                        campo[1].Clear();
                        rootSession.Keyboard.SendKeys(data[indexVal]);
                        //Aqui se puede enviar un valor

                        var btn = rootSession.FindElementsByClassName("TBitBtn");
                        ////Debugger.Launch();
                        foreach (var boton in btn)
                        {
                            if (boton.Text == "Aceptar")
                            {
                                rootSession.Mouse.MouseMove(boton.Coordinates, 10, 10);
                                rootSession.Mouse.Click(null);
                            }
                        }
                        indexVal++;
                    }
                }
            }
        }

        public static Dictionary<string, Point> findname1(WindowsDriver<WindowsElement> desktopSession, int offset, int indiceBarra)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            List<string> errorMessages = new List<string>();
            Dictionary<string, Point> coordenadasBotones = new Dictionary<string, Point>();
            List<string> botones = new List<string>() { "Insertar", "Borrar", "Aplicar", "Cancelar" };
            Point coord = new Point();
            var ElementList = PruebaCRUD.NavClass(desktopSession);
            ////Debugger.Launch();
            int x = ElementList[indiceBarra].Size.Width / 2;
            int y = ElementList[indiceBarra].Size.Height / 2;
            int nameCounter = 0;
            WindowsElement boton;
            desktopSession.Mouse.MouseMove(ElementList[indiceBarra].Coordinates, x + 50, y + 50);
            x += 10;
            while (nameCounter < botones.Count && x < ElementList[indiceBarra].Size.Width)
            {
                desktopSession.Mouse.MouseMove(ElementList[indiceBarra].Coordinates, x, y);

                Thread.Sleep(300);

                boton = desktopSession.FindElementByName(botones[nameCounter]);

                if (boton != null)
                {
                    coord.X = x;
                    coord.Y = ElementList[indiceBarra].Size.Height / 2;
                    if (!coordenadasBotones.ContainsKey(botones[nameCounter]))
                    {
                        coordenadasBotones.Add(botones[nameCounter], coord);
                    }
                    if (nameCounter < botones.Count)
                    {
                        nameCounter++;
                    }
                }

                x += offset;
            }


            return coordenadasBotones;
        }



        public static List<string> ReportePreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, int tipo, string nomPrograma = null)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            if (bandera == "0" || bandera == "1")
            {
                var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
                {
                    Timeout = TimeSpan.FromSeconds(60),
                    PollingInterval = TimeSpan.FromSeconds(1)
                };
                wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
                WindowsDriver<WindowsElement> rootSession = null;
                WindowsElement codPanel = null;
                bool brk = false;
                bool IsDisplayedPreWin = false;
                wait.Until(driver =>
                {
                    List<Point> coordenadas = FuncionesGlobales.coordinatesFinder(desktopSession, 227, 117, 0);
                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
                    Thread.Sleep(1000);

                    try
                    {
                        desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(2000);
                        if (bandera == "0")
                        {
                            if (tipo == 1)
                            {
                                WindowsDriver<WindowsElement> rootSession1 = null;
                                rootSession1 = RootSession();
                                List<Point> coordenadasAceptar = CrudKfdplcur.coordinatesFinder(rootSession1, 66, 159, 6);
                                int coordX1 = coordenadasAceptar[0].X;
                                int coordY1 = coordenadasAceptar[0].Y;
                                var parentElement1 = rootSession1.FindElement(By.Name(rootSession1.Title));
                                rootSession1.Mouse.MouseMove(parentElement.Coordinates, coordX1, coordY1);
                                rootSession1.Mouse.Click(null);
                                //WindowsDriver<WindowsElement> rootSession1 = null;
                                //rootSession1 = RootSession();
                                ////rootSession1 = ReloadSession(rootSession1, "TFrmFdManre");
                                //var Preview = rootSession1.FindElementsByClassName("TFrmFdManre");
                                //Thread.Sleep(1000);
                                //rootSession.Mouse.MouseMove(Preview[0].Coordinates, 92, 270);
                                //rootSession.Mouse.Click(null);
                                //Thread.Sleep(2000);

                                //errors = GenerarReportes("Preliminar", file, pdf1, nomPrograma);
                                //if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                            }
                           

                        }
                    }
                    catch (Exception ex)
                    {
                        errorMessages.Add($"Error de Appium: {ex.ToString()}");
                    }
                    codPanel = parentElement;
                    return codPanel != null;
                });
            }
            return errorMessages;
        }


        public static List<string> GenerarReportes(string TipoReporte, string file, string ruta, string nomPrograma = null)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession;
            Thread.Sleep(4000);
            string ReportePdf = "";
            string ReporteW = "";
            string ReporteE = "";
            string Ruta1 = "";
            string Extension1 = "";
            string Ruta2 = "";
            string Extension2 = "";
            string Ruta3 = "";
            string Extension3 = "";
            bool Pdf = false;
            bool Word = false;
            bool Excel = false;

            // Revisar que se abrio el preview
            bool IsDisplayedReview = false;

            try
            {
                for (int i = 0; i < 20; i++)
                {
                    try
                    {
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                        var validar = rootSession.FindElementsByClassName("TfrxPreviewForm");
                        break;
                    }
                    catch
                    {
                        Thread.Sleep(1000);
                    }
                }
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxPreviewForm");

                int count = 0;
                while (count < 300)
                {
                    try
                    {
                        List<Point> carga = PruebaCRUD.CoordinatesFinder(rootSession, 224, 32, 64);
                        break;
                    }
                    catch
                    {
                        Thread.Sleep(1000);
                        count++;
                    }
                }

                List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(rootSession, 247, 223, 99);
                int CoordX = coordenadas[0].X;
                int CoordY = coordenadas[0].Y;

                var PanelElements = rootSession.FindElement(By.Name(rootSession.Title));
                bool IsDisplayedReport = false;

                // Clic en el icono de editar reporte
                rootSession.Mouse.MouseMove(PanelElements.Coordinates, CoordX, CoordY);
                rootSession.Mouse.Click(null);
                Thread.Sleep(8000);
                newFunctions_4.ScreenshotNuevo("La opción 'Editar Reporte' " + TipoReporte + " está activada", true, file);

                errorMessages.Add("El Preview del reporte " + TipoReporte + " tiene habilitado la opcion 'Editar Reporte'");

                // Cerrar la ventana de editar Reporte
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxDesignerForm");

                var CerrarEditar = rootSession.FindElementByName("Cerrar");
                CerrarEditar.Click();
                Thread.Sleep(2000);
                newFunctions_4.ScreenshotNuevo("Ventana 'Editar Reporte' " + TipoReporte + " Cerrada", true, file);
            }
            catch
            {

            }

            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TfrxPreviewForm");

            var ElementsPreview = rootSession.FindElementsByClassName("TToolBar");

            if (ElementsPreview.Count > 0)
            {
                IsDisplayedReview = true;
            }
            else
            {
                errorMessages.Add("No se puede visualizar la ventana Preview del Reporte " + TipoReporte);
            }

            string fecha = Hora();
            if (IsDisplayedReview == true)
            {
                ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                int cerrar = 0;
                try
                {
                    Thread.Sleep(2000);
                    List<Point> coordenadas1 = PruebaCRUD.CoordinatesFinder(rootSession, 224, 32, 64);
                    int CoordX = coordenadas1[0].X;
                    int CoordY = coordenadas1[0].Y;

                    var PanelElements = rootSession.FindElement(By.Name(rootSession.Title));

                    // Clic en el icono de exportar PDF
                    rootSession.Mouse.MouseMove(PanelElements.Coordinates, CoordX, CoordY);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TfrxPDFExportDialog");

                    var ExportPDF = rootSession.FindElementByName("Aceptar");

                    if (ExportPDF != null)
                    {
                        //Clic en Aceptar exportar PDF
                        ExportPDF.Click();
                        Thread.Sleep(2000);
                        rootSession = RootSession();

                        try
                        {
                            //ingresar la ruta
                            rootSession.Keyboard.SendKeys(ruta);
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            Thread.Sleep(2000);
                            //Abrir PDF
                            Process.Start(ruta + ".pdf");
                            Thread.Sleep(4000);
                            rootSession = RootSession();
                            rootSession = ReloadSession(rootSession, "AcrobatSDIWindow");
                            //var TraerPantalla = rootSession.FindElementsByName("AVL_AVView");
                            newFunctions_4.ScreenshotNuevo("Reporte " + TipoReporte + " exportado a PDF 1", true, file);
                            LimpiarProcesos();
                        }
                        catch
                        {
                            var Ok = rootSession.FindElementByName("OK");
                            Ok.Click();
                            Thread.Sleep(1000);
                            errorMessages.Add("No existe el botón exportar reporte a PDF");
                        }
                    }
                }
                catch
                {
                    errorMessages.Add("En el Preview  del Reporte " + TipoReporte + " NO Existe el Botón Exportar Reporte a PDF");
                    cerrar = 1;
                }

                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                Thread.Sleep(1000);

                //Clic en Disquet
                rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 57, 17);
                rootSession.Mouse.Click(null);
                Thread.Sleep(2000);

                //Reportes
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "#32768");
                var MenuItem = rootSession.FindElementsByXPath("/Menu[@Name=\"Contexto\"][@ClassName=\"#32768\"]/*");

                for (int i = 0; i < MenuItem.Count; i++)
                {
                    if (MenuItem[i].GetAttribute("Name").Contains("PDF"))
                    {
                        Pdf = true;
                        ReportePdf = MenuItem[i].GetAttribute("Name");
                        Ruta1 = @"C:\Reportes\Reporte_PDF_" + TipoReporte + "_" + fecha;
                        Extension1 = ".pdf";
                    }

                    if (MenuItem[i].GetAttribute("Name").Contains("Word"))
                    {
                        Word = true;
                        ReporteW = MenuItem[i].GetAttribute("Name");
                        Ruta2 = @"C:\Reportes\Reporte_Word_" + TipoReporte + "_" + fecha;
                        Extension2 = ".docx";
                    }

                    if (MenuItem[i].GetAttribute("Name").Contains("Excel") && MenuItem[i].GetAttribute("Name").Contains("OLE"))
                    {

                    }
                    else if (MenuItem[i].GetAttribute("Name").Contains("Excel"))
                    {
                        Excel = true;
                        ReporteE = MenuItem[i].GetAttribute("Name");
                        Ruta3 = @"C:\Reportes\Reporte_Excel_" + TipoReporte + "_" + fecha;
                        Extension3 = ".xlsx";
                    }
                }

                if (Pdf == true)
                {
                    try
                    {
                        ExportarReporte(ReportePdf, Ruta1, Extension1, TipoReporte, file);
                    }
                    catch
                    {
                        errorMessages.Add("El Proceso de exportar el reporte " + TipoReporte + " a PDF Falló");
                    }
                }

                if (Word == true)
                {
                    try
                    {
                        ExportarReporte(ReporteW, Ruta2, Extension2, TipoReporte, file);
                    }
                    catch
                    {
                        errorMessages.Add("El Proceso de exportar el reporte " + TipoReporte + " a Word Falló");
                    }
                }

                if (Excel == true)
                {
                    try
                    {
                        ExportarReporte(ReporteE, Ruta3, Extension3, TipoReporte, file, nomPrograma);
                    }
                    catch
                    {
                        errorMessages.Add("El Proceso de exportar el reporte " + TipoReporte + " a Excel Falló");
                    }
                }

                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                Thread.Sleep(1000);

                if (cerrar == 0)
                {
                    //Cerrar Preview
                    rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 600, 17);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    try
                    {
                        ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                        //Cerrar Preview
                        rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 543, 17);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(2000);
                        try
                        {
                            rootSession.Mouse.MouseMove(rootSession.FindElementByName("Cerrar").Coordinates);
                            rootSession.Mouse.Click(null);
                            errorMessages.Add("El Botón 'Cerrar' no Existe en el Preview del Reporte " + TipoReporte);
                        }
                        catch
                        {

                        }
                    }
                    catch
                    {

                    }
                }
                else
                {
                    //Cerrar Preview
                    rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 520, 17);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    try
                    {
                        ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                        //Cerrar Preview
                        rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 543, 17);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(2000);
                        try
                        {
                            rootSession.Mouse.MouseMove(rootSession.FindElementByName("Cerrar").Coordinates);
                            rootSession.Mouse.Click(null);
                            errorMessages.Add("El Botón 'Cerrar' no Existe en el Preview del Reporte " + TipoReporte);
                        }
                        catch
                        {

                        }
                    }
                    catch
                    {

                    }
                }
            }
            return errorMessages;
        }

        private static void ExportarReporte(string NomReporte, string Ruta, string Extension, string TipoReporte, string file, string nomPrograma = null)
        {
            WindowsDriver<WindowsElement> rootSession;

            if (NomReporte.Contains("Word") || NomReporte.Contains("Excel"))
            {
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                var ElementsPreview = rootSession.FindElementsByClassName("TToolBar");

                Thread.Sleep(1000);
                //Clic en Disquet
                rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 57, 17);
                rootSession.Mouse.Click(null);
                Thread.Sleep(1000);
            }

            rootSession = RootSession();
            //clic en Tipo Reporte
            var ClicReport = rootSession.FindElementByName(NomReporte);
            ClicReport.Click();
            Thread.Sleep(1000);
            rootSession = RootSession();
            var AcceptPDF = rootSession.FindElementByName("Aceptar");

            //clic en aceptar exportar Reporte
            AcceptPDF.Click();
            Thread.Sleep(2000);

            //ingresar la ruta
            rootSession.Keyboard.SendKeys(Ruta);
            Thread.Sleep(1000);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(3000);

            if (NomReporte.Contains("Excel"))
            {
                if (File.Exists(Ruta + Extension))
                {
                    if (nomPrograma == null)
                    {
                        //Abrir reporte
                        Process.Start(Ruta + Extension);
                        Thread.Sleep(4000);
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "XLMAIN");
                        newFunctions_4.ScreenshotNuevo("Reporte " + TipoReporte + " Exportado a Excel", true, file);
                        LimpiarProcesos();
                    }
                }
                else
                {
                    if (nomPrograma == null)
                    {
                        Extension = ".xls";
                        //Abrir reporte
                        Process.Start(Ruta + Extension);
                        Thread.Sleep(4000);
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "XLMAIN");
                        newFunctions_4.ScreenshotNuevo("Reporte " + TipoReporte + " Exportado a Excel", true, file);
                        LimpiarProcesos();
                    }
                }
            }
            else
            {
                //Abrir reporte
                Process.Start(Ruta + Extension);
                Thread.Sleep(4000);
                if (Extension == ".docx")
                {
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "OpusApp");
                    newFunctions_4.ScreenshotNuevo("Reporte " + TipoReporte + " Exportado a Word", true, file);
                }
                else if (Extension == ".pdf")
                {
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "AcrobatSDIWindow");
                    newFunctions_4.ScreenshotNuevo("Reporte " + TipoReporte + " Exportado a Pdf", true, file);
                }
                LimpiarProcesos();
            }
        }
        public static List<Point> coordinatesFinder(WindowsDriver<WindowsElement> session, int R, int G, int B)
        {
            Screenshot image = ((ITakesScreenshot)session).GetScreenshot();
            string name = FuncionesVitales.Hora();
            string path = "C:\\imagenesTest\\" + name + ".Png";
            Image imgSource;
            try
            {
                image.SaveAsFile(string.Format("C:\\imagenesTest\\{0}.Png", name), ScreenshotImageFormat.Png);
                Thread.Sleep(1000);
                imgSource = Image.FromFile(path);
            }
            catch (Exception e)
            {
                image.SaveAsFile(string.Format("C:\\imagenesTest\\{0}.Png", name), ScreenshotImageFormat.Png);
                Thread.Sleep(1000);
                imgSource = Image.FromFile(path);
            }

            Point puntos = new Point();
            int lastx = 0;
            int lasty = 0;
            List<Point> coordenadas = new List<Point>();
            Bitmap bmpSource = new Bitmap(imgSource);
            Bitmap bmpDest = new Bitmap(bmpSource.Width, bmpSource.Height);
            int romper = 0;
            for (int y = 0; y < bmpSource.Height; y++)
            {
                for (int x = 0; x < bmpSource.Width; x++)
                {
                    Color clrPixel = bmpSource.GetPixel(x, y);
                    if (clrPixel.R == R && clrPixel.G == G && clrPixel.B == B)
                    {
                        puntos.X = x;
                        puntos.Y = y;
                        if (x > lastx + 10 || y > lasty + 10)
                        {
                            coordenadas.Add(puntos);
                            lastx = x;
                            lasty = y;
                            romper++;
                            break;
                        }
                    }
                }
                if (romper != 0)
                {
                    break;
                }
            }
            return coordenadas;
        }
    }
}

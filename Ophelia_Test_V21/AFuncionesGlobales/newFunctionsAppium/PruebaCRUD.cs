
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

namespace Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium
{
    class PruebaCRUD : FuncionesVitales
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

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo)
        {
            List<string> errorMessages = new List<string>();
            
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            //Debug.WriteLine($"la cantidad de elementos es: {ElementList.Count}");
            int cont = 0;
            List<string> variables = new List<string>() { "N", "G", "PRUEBACRUD", "3939" };
            if (tipo == 1)
            {
                foreach (var elem in ElementList)
                {
                    if (elem.Enabled == true)
                    {
                        elem.SendKeys(variables[cont]);
                        cont += 1;
                    }
                    Thread.Sleep(500);
                }
            }
            else
            {
                ElementList[4].Clear();
                ElementList[4].SendKeys("Natalia");
            }
        }

        public static System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> NavClass(WindowsDriver<WindowsElement> desktopSession)
        {
            var ElementList = desktopSession.FindElementsByClassName("TKNavegador");
            Debug.WriteLine($"Navigators find: {ElementList.Count}");
            return ElementList;
        }

        public static Dictionary<string, Point> findname(WindowsDriver<WindowsElement> desktopSession, int offset, int indiceBarra)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            List<string> errorMessages = new List<string>();
            Dictionary<string, Point> coordenadasBotones = new Dictionary<string, Point>();
            List<string> botones = new List<string>() { "Insertar", "Borrar", "Editar", "Aplicar", "Cancelar"};
            Point coord = new Point();
            var ElementList = NavClass(desktopSession);
            ////Debugger.Launch();
            int x = ElementList[indiceBarra].Size.Width/2;
            int y = ElementList[indiceBarra].Size.Height / 2;
            int nameCounter = 0;
            WindowsElement boton;
            desktopSession.Mouse.MouseMove(ElementList[indiceBarra].Coordinates, x+20, y+50);
            while (nameCounter < botones.Count && x < ElementList[indiceBarra].Size.Width) {
                desktopSession.Mouse.MouseMove(ElementList[indiceBarra].Coordinates, x, y);
                
                Thread.Sleep(300);
                
                boton = desktopSession.FindElementByName(botones[nameCounter]);

                if (boton != null) {
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

       
        public static void CRUDHabda(WindowsDriver<WindowsElement> desktopSession)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBLookupComboBox");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList3 = desktopSession.FindElementsByClassName("TDBRichEdit");
            Debug.WriteLine($"la cantidad de elementos es: {ElementList.Count}");
            Debug.WriteLine($"la cantidad de elementos es: {ElementList2.Count}");
            Debug.WriteLine($"la cantidad de elementos es: {ElementList3.Count}");

            ElementList[0].SendKeys("H");
            //new Actions(desktopSession).MoveToElement(ElementList[1], 0, 0).MoveByOffset(ElementList2[2].Size.Height / 2, ElementList2[2].Size.Width - 5).ClickAndHold();
            ElementList[1].SendKeys("V");
            ElementList3[0].SendKeys("PRUEBACRUD");
            
        }

        public static void CRUDAusen(WindowsDriver<WindowsElement> desktopSession, int tipo,List<string> crudVars, string file)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            if (tipo == 1)
            {
                foreach (Point coord in coordenadas)
                {
                    Thread.Sleep(2000);
                    //Actions noteClicks = new Actions(desktopSession);
                    //noteClicks.MoveToElement(parentElement).MoveByOffset(coord.X + 10, coord.Y).Click().Perform();

                    desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                    desktopSession.Mouse.Click(null);
                    if (contLupa == 1)
                    {
                        string QbeFilter = crudVars[0];
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TKQBE");
                        if (rootSession != null)
                        {
                            IsDisplayedQbe = true;
                        }
                        else
                        {
                            errorMessages.Add($"No puede encontrar la ventana de QBE");
                        }
                        if (IsDisplayedQbe)
                        {
                            var elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                            if (!string.IsNullOrEmpty(QbeFilter))
                            {
                                elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/Pane[contains(@ClassName, 'TStringGrid')]/*"));
                                elements[0].SendKeys(QbeFilter);
                            }
                            elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                            //new Actions(rootSession).MoveToElement(elements[8], 0, 0).MoveByOffset(10, 10).DoubleClick().Perform();
                            rootSession.Mouse.MouseMove(elements[8].Coordinates, 20, 20);
                            rootSession.Mouse.Click(null);
                            Thread.Sleep(500);

                            rootSession = RootSession();
                            rootSession = ReloadSession(rootSession, "TCaptura");
                            var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                            ////Debugger.Launch();
                            rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                            rootSession.Mouse.Click(null);
                            rootSession.Keyboard.SendKeys(crudVars[0]);
                            //Aqui se puede enviar un valor

                            var btn = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Pane[@ClassName=\"TPanel\"]/Button[@ClassName=\"TBitBtn\"][@Name=\"Aceptar\"]");
                            rootSession.Mouse.MouseMove(btn[0].Coordinates, 10, 10);
                            rootSession.Mouse.Click(null);

                        }
                    }
                    else if (contLupa == 2)
                    {
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TCaptura");
                        var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\"LISTA DE CONCEPTOS\"]/Edit[@ClassName=\"TEdit\"]");
                        ////Debugger.Launch();
                        rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                        rootSession.Mouse.Click(null);
                        rootSession.Keyboard.SendKeys(crudVars[1]);
                        //Aqui se puede enviar un valor

                        var btn = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\"LISTA DE CONCEPTOS\"]/Pane[@ClassName=\"TPanel\"]/Button[@ClassName=\"TBitBtn\"][@Name=\"Aceptar\"]");
                        rootSession.Mouse.MouseMove(btn[0].Coordinates, 10, 10);
                        rootSession.Mouse.Click(null);
                    }
                    else
                    {
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TCaptura");
                        var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\"MOTIVOS DE AUSENCIA\"]/Edit[@ClassName=\"TEdit\"]");
                        ////Debugger.Launch();
                        rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                        rootSession.Mouse.Click(null);
                        rootSession.Keyboard.SendKeys(crudVars[2]);
                        //Aqui se puede enviar un valor

                        var btn = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\"MOTIVOS DE AUSENCIA\"]/Pane[@ClassName=\"TPanel\"]/Button[@ClassName=\"TBitBtn\"][@Name=\"Aceptar\"]");
                        rootSession.Mouse.MouseMove(btn[0].Coordinates, 10, 10);
                        rootSession.Mouse.Click(null);
                    }
                    contLupa++;
                }
                var campo2 = desktopSession.FindElementByClassName("TDBMemo");
                campo2.Clear();
                campo2.SendKeys(crudVars[2]);
                Screenshot("Agregar Registro", true, file,desktopSession);
            }
            else
            {
                var campo = desktopSession.FindElementByClassName("TDBMemo");
                campo.Clear();
                campo.SendKeys(crudVars[3]);
                Screenshot("Editar Registro", true, file,desktopSession);
            }
        }

        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 0;
            foreach (Point coord in coordenadas)
            {
                Thread.Sleep(2000);
                //Actions noteClicks = new Actions(desktopSession);
                //noteClicks.MoveToElement(parentElement).MoveByOffset(coord.X + 10, coord.Y).Click().Perform();

                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                desktopSession.Mouse.Click(null);
                if (contLupa == 1)
                {
                    try
                    {
                        rootSession = RootSession();
                        rootSession = PruebaCRUD.ReloadSession(rootSession, "TKQBE");
                    }
                    catch
                    {

                    }
                    if (rootSession != null)
                    {
                        IsDisplayedQbe = true;
                    }
                    else
                    {
                        errorMessages.Add($"No puede encontrar la ventana de QBE");
                    }
                    if (IsDisplayedQbe)
                    {
                        var elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        if (!string.IsNullOrEmpty(data[indexVal]))
                        {
                            elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/Pane[contains(@ClassName, 'TStringGrid')]/*"));
                            elements[0].SendKeys(data[indexVal]);
                        }
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        new Actions(rootSession).MoveToElement(elements[8], 0, 0).MoveByOffset(20, 10).DoubleClick().Perform();
                        Thread.Sleep(500);
                        IsDisplayedQbe = false;

                    }
                    Thread.Sleep(1000);
                    rootSession = RootSession();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                    //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    ////Debugger.Launch();
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(1500);
                    rootSession.Keyboard.SendKeys(data[indexVal]);
                    //Aqui se puede enviar un valor

                    var btn = rootSession.FindElementsByClassName("TBitBtn");
                    //Debugger.Launch();
                    foreach (var boton in btn)
                    {
                        if (boton.Text == "Aceptar")
                        {
                            rootSession.Mouse.MouseMove(boton.Coordinates, 10, 10);
                            rootSession.Mouse.Click(null);
                            Thread.Sleep(15000);
                        }
                    }
                    indexVal++;
                }
            }

        }

        public static void LupaDinamica2(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 0;
            foreach (Point coord in coordenadas)
            {
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                desktopSession.Mouse.Click(null);
                if (contLupa == 1)
                {
                    Thread.Sleep(2000);
                    rootSession = RootSession();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                    //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
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

        public static List<Point> CoordinatesFinder(WindowsDriver<WindowsElement> session, int R, int G, int B)
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
            coordenadas = PixelSeeker(bmpSource, R, G, B, new Point(20, 20));

            //if (coordenadas.Count == 0) coordenadas = PixelSeeker(bmpSource, R, G, B, new Point(20, 20));

            imgSource.Dispose();

            ////Debugger.Launch();
            return coordenadas;
            
        }
        public static List<Point> PixelSeeker(Bitmap image, int R, int G, int B, Point offset)
        {

            Point punto = new Point();
            int lastx = 0;
            int lasty = 0;
            List<Point> coordinates = new List<Point>();

            for (int y = 0; y < image.Height; y++)
            {
                for (int x = 0; x < image.Width; x++)
                {
                    Color clrPixel = image.GetPixel(x, y);
                    if (clrPixel.R == R && clrPixel.G == G && clrPixel.B == B)
                    {
                        punto.X = x;
                        punto.Y = y;
                        if (x > lastx + offset.X || y > lasty + offset.Y)
                        {
                            coordinates.Add(punto);
                            lastx = x;
                            lasty = y;
                        }
                    }
                }
            }


            image.Dispose();
            return coordinates;

        }

        public static void CRUDRango(WindowsDriver<WindowsElement> desktopSession, int tipo)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            Debug.WriteLine($"la cantidad de elementos es: {ElementList.Count}");
            int cont = 0;
            List<string> variables = new List<string>() { "12345", "PRUEBACRUD", "3939","Inicial"};
            if (tipo == 1)
            {
                foreach (var elem in ElementList)
                {
                    if (elem.Enabled == true)
                    {
                        elem.SendKeys(variables[cont]);
                        cont += 1;
                    }
                    Thread.Sleep(500);
                }
                var checkbox = desktopSession.FindElementsByClassName("TDBCheckBox");
                ////Debugger.Launch();
                foreach(var check in checkbox)
                {
                    //Debug.WriteLine($"nombre atributos: {check.GetAttribute("Name")}");
                    if(check.GetAttribute("Name")==variables[cont])
                    {
                       check.Click();
                       cont+= 1;
                    }
                }

            }
            else
            {
                ElementList[3].Clear();
                ElementList[3].SendKeys("EDITAR PRUEBA");
            }

        }
        public static void CRUDDetalleRango(WindowsDriver<WindowsElement> desktopSession, int tipo)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
            //Debug.WriteLine($"la cantidad de elementos es: {ElementList.Count}");
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 10, 10);
            desktopSession.Mouse.Click(null);
            int cont = 0;
            List<string> variables = new List<string>() { "123", "45", "1010", "50" };
            if (tipo == 1)
            {
                for(int i=0; i < variables.Count; i++)
                {
                    //Keyboard.SendKeys(variables[cont]);
                    cont += 1;
                    //Keyboard.SendKeys("{RIGHT}");
                    Thread.Sleep(500);
                }
                
            }
            else
            {
                ElementList[1].Clear();
                ElementList[1].SendKeys("10");
            }

        }

        public static void cerrarBorrar(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            string navigator = AbrirPrograma.SelectNavigator();
            if (navigator == "IE")
            {
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "IEFrame");
                var allFrame = rootSession.FindElementsByClassName("IEFrame");
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
                rootSession = RootSession();
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                rootSession = RootSession();
                //rootSession = ReloadSession(rootSession, "Chrome_WidgetWin_1");
                var allFrame = rootSession.FindElementsByClassName("#32769");
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
                rootSession = RootSession();
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }

        }

        public static void EliminarRegistro(WindowsDriver<WindowsElement> desktopSession, WindowsElement ElementList, Dictionary<string, Point> botonesNavegador, string file)
        {
            desktopSession.Mouse.MouseMove(ElementList.Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);

            // New
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = PruebaCRUD.RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            //
            //OLD
            //PruebaCRUD.cerrarBorrar(desktopSession);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
            Thread.Sleep(2000);

        }
        public static bool ValEditOra(WindowsDriver<WindowsElement> desktopSession) 
        {
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var ventana = rootSession.FindElementsByClassName("TReconcileErrorForm");
            
            /*Thread.Sleep(1000);
            if (ventana == null)
            {
                rootSession = RootSession();
                var barraTareas = rootSession.FindElementsByClassName("Shell_TrayWnd");
                rootSession.FindElementByName("Host de Windows Presentation Foundation - 1 ventana de ejecución").Click();
                Thread.Sleep(1000);
                rootSession = RootSession();
                ventana = rootSession.FindElementsByClassName("TReconcileErrorForm");
            }
            */

            bool ValEdit = false;

            if (ventana.Count > 0) {  rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter); ValEdit = true; }
            return ValEdit;
        }
        //SE USA PARA CUANDO HAY DETALLES SON EL BOTON DE EDITAR
        public static Dictionary<string, Point> findname1(WindowsDriver<WindowsElement> desktopSession, int offset, int indiceBarra)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            List<string> errorMessages = new List<string>();
            Dictionary<string, Point> coordenadasBotones = new Dictionary<string, Point>();
            List<string> botones = new List<string>() { "Insertar", "Borrar", "Aplicar", "Cancelar" };
            Point coord = new Point();
            var ElementList = NavClass(desktopSession);
            ////Debugger.Launch();
            int x = ElementList[indiceBarra].Size.Width / 2;
            int y = ElementList[indiceBarra].Size.Height / 2;
            int nameCounter = 0;
            WindowsElement boton;
            desktopSession.Mouse.MouseMove(ElementList[indiceBarra].Coordinates, x + 50, y + 50);
            x += 10;
            while (nameCounter < botones.Count && x < ElementList[indiceBarra].Size.Width)
            {
                desktopSession.Mouse.MouseMove(ElementList[indiceBarra].Coordinates, x , y);

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

        public static List<string> QBEQry(WindowsDriver<WindowsElement> desktopSession, string bandera, string QbeFilter, string file=null,string campo=null)
        {
            List<string> errorMessages = new List<string>();
            if (bandera == "0")
            {
                var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
                {
                    Timeout = TimeSpan.FromSeconds(120),
                    PollingInterval = TimeSpan.FromSeconds(1)
                };
                wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
                wait.Until(driver =>
                {
                    WindowsDriver<WindowsElement> rootSession = null;
                    List<Point> coordenadas = FuncionesGlobales.coordinatesFinder(desktopSession, 244, 202, 32);

                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));

                    bool IsDisplayedQbe = false;


                    desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);
                    ////Debugger.Launch();
                    newFunctions_4.ScreenshotNuevo("QBE", true, file);
                    rootSession = RootSession();
                    rootSession = FuncionesGlobales.ReloadQbeSession(rootSession, "TKQBE");
                    if (rootSession != null)
                    {
                        IsDisplayedQbe = true;
                    }
                    else
                    {
                        errorMessages.Add($"No puede encontrar la ventana de QBE");
                    }

                    if (IsDisplayedQbe)
                    {
                        var elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        if (!string.IsNullOrEmpty(QbeFilter))
                        {
                            elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/Pane[contains(@ClassName, 'TStringGrid')]/*"));
                            if (campo !=null)
                            {
                                elements[0].Click();
                                int m = Convert.ToInt32(campo);
                                for (int i = 1; i <= m; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                rootSession.Keyboard.SendKeys(QbeFilter);
                            }
                            else
                            {
                                elements[0].SendKeys(QbeFilter);
                            }
                            
                        }
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        rootSession.Mouse.MouseMove(elements[8].Coordinates, 20, 20);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(500);
                       
                        if (string.IsNullOrEmpty(QbeFilter))
                        {
                            rootSession = FuncionesGlobales.ReloadQbeSession(rootSession, "TKQBE");
                            elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("//Button[contains(@Name, 'Sí')]"));
                            rootSession.Mouse.MouseMove(elements[0].Coordinates, 20, 20);
                            rootSession.Mouse.Click(null);
                        }
                        Thread.Sleep(500);
                        newFunctions_4.ScreenshotNuevo("Registros QBE", true, file);
                    }

                    return rootSession != null;
                });
            }
            else
            {
                var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
                {
                    Timeout = TimeSpan.FromSeconds(120),
                    PollingInterval = TimeSpan.FromSeconds(1)
                };
                wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
                wait.Until(driver =>
                {
                    WindowsDriver<WindowsElement> rootSession = null;
                    List<Point> coordenadas = FuncionesGlobales.coordinatesFinder(desktopSession, 224, 209, 6);

                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));

                    bool IsDisplayedQbe = false;


                    desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    newFunctions_4.ScreenshotNuevo("QBE", true, file);
                    rootSession = RootSession();
                    rootSession = FuncionesGlobales.ReloadQbeSession(rootSession, "TKQBE");
                    if (rootSession != null)
                    {
                        IsDisplayedQbe = true;
                    }
                    else
                    {
                        errorMessages.Add($"No puede encontrar la ventana de QBE");
                    }

                    if (IsDisplayedQbe)
                    {
                        var elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        if (!string.IsNullOrEmpty(QbeFilter))
                        {
                            elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/Pane[contains(@ClassName, 'TStringGrid')]/*"));
                            if (campo != null)
                            {
                                elements[0].Click();
                                int m = Convert.ToInt32(campo);
                                for (int i = 1; i <= m; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                rootSession.Keyboard.SendKeys(QbeFilter);
                            }
                            else
                            {
                                elements[0].SendKeys(QbeFilter);
                            }

                        }
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        rootSession.Mouse.MouseMove(elements[8].Coordinates, 20, 20);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(500);

                        if (string.IsNullOrEmpty(QbeFilter))
                        {
                            rootSession = FuncionesGlobales.ReloadQbeSession(rootSession, "TKQBE");
                            elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("//Button[contains(@Name, 'Sí')]"));
                            rootSession.Mouse.MouseMove(elements[0].Coordinates, 20, 20);
                            rootSession.Mouse.Click(null);
                        }
                        Thread.Sleep(500);
                        newFunctions_4.ScreenshotNuevo("Registros QBE", true, file);
                    }

                    return rootSession != null;
                });
            }
            return errorMessages;
        }

        public static List<string> Sumatory(WindowsDriver<WindowsElement> desktopSession, string BanderaSum, string file,string mover=null)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            int Height = Screen.PrimaryScreen.Bounds.Height;

            if (BanderaSum == "1")
            {
                var allWindowHandles = rootSession.WindowHandles;
                WebDriverWait wait = new WebDriverWait(rootSession, TimeSpan.FromSeconds(1));

                bool validar = newFunctions_1.coordinatesFinder(desktopSession, 240, 240, 240, 1550, 330, 2);
                if (validar == true)
                {
                    newFunctions_1.lapiz(desktopSession);
                }

                int counterTry = 0;
                ////Debugger.Launch();
                var fieldInt = rootSession.FindElementsByClassName("TKactusDBgrid");

                new Actions(rootSession).MoveToElement(fieldInt[0], 0, 0).MoveByOffset(50, 40).ClickAndHold().Click().Perform();
                Thread.Sleep(2000);
                new Actions(rootSession).Click().Perform();
                int contscreen = 0;
                if(mover!=null)
                {
                    int m = Convert.ToInt32(mover);
                    for (int i = 1; i <= m; i++)
                    {
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                    }                    
                }
                while (counterTry < 1)
                {
                    try
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Alt + "H" + OpenQA.Selenium.Keys.Alt);

                        for (int i = 1; i <= 4; i++)
                        {
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                            Thread.Sleep(100);
                        }

                        //PRIMER SCREEN DEL MENU SUMATORIA
                        if (contscreen == 0) { newFunctions_4.ScreenshotNuevo("Menú Sumatoria", true, file); ; contscreen = 1; }

                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(2000);

                        bool ventana = newFunctions_1.coordinatesFinder(desktopSession, 27, 101, 206, screenWidth / 2, 347, 1);
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "IEFrame");
                        var allFrame = rootSession.FindElementsByClassName("IEFrame");
                        new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
                        rootSession = RootSession();

                        if (ventana == false)
                        {
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            Thread.Sleep(2000);
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                        }
                        else
                        {
                            newFunctions_4.ScreenshotNuevo("Valor Sumatoria", true, file);
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            counterTry += 1;
                            break;
                        }
                    }
                    catch (Exception)
                    {
                        errorMessages.Add("Se ha encontrado una Excepción y se sale el try");
                        break;
                    }
                }
            }
            return errorMessages;
        }

        public static List<string> VentanaAzul(WindowsDriver<WindowsElement> desktopSession)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            //Debugger.Launch();
            bool ventana = newFunctions_1.coordinatesFinder(desktopSession, 27, 101, 206, screenWidth / 2, 347, 1);
            rootSession = RootSession();
            string navigator = AbrirPrograma.SelectNavigator();
            if (navigator == "IE")
            {
                rootSession = ReloadSession(rootSession, "IEFrame");
                var allFrame = rootSession.FindElementsByClassName("IEFrame");
                Thread.Sleep(1000);
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
                rootSession = RootSession();
                if (ventana == false)
                {
                    //new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
                    //rootSession = RootSession();
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(2000);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    errorMessages.Add("No hay registros para exportar o el generar excel produce un error");
                    Thread.Sleep(2000);
                }
            }
            else
            {
                //rootSession = ReloadSession(rootSession, "Chrome_WidgetWin_1");
                var allFrame = rootSession.FindElementsByClassName("#32769");
                Thread.Sleep(1000);
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
                rootSession = RootSession();

                if (ventana == false)
                {
                    //new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
                    //rootSession = RootSession();
                    Thread.Sleep(1000);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(2000);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    errorMessages.Add("No hay registros para exportar o el generar excel produce un error");
                    Thread.Sleep(2000);
                }
            }
            return errorMessages;
        }

    }
}

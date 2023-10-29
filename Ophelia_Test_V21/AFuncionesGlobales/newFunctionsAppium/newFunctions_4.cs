using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales;
using System.Drawing.Imaging;


namespace Ophelia_Test_V21.AFuncionesGlobales.newFolder
{
    class newFunctions_4 : FuncionesVitales
    {

        static public WindowsDriver<WindowsElement> RootSessionNew()
        {
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }


        static public WindowsDriver<WindowsElement> ReloadXSession(WindowsDriver<WindowsElement> session, string clase)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(clase);
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


        //Funcion de Notas
        public static void openInnerNote(WindowsDriver<WindowsElement> desktopSession, string file)
        {

            string closeNote = "//Window[contains(@ClassName,TFrmNotas)]/TitleBar[contains(@AutomationId,TitleBar)]/Button[contains(@Name,Cerrar)][contains(@AutomationId,Close)]";
            string editFieldxpath = "//Pane[contains(@ClassName,TGroupBox)]/Edit[contains(@ClassName,TDBEdit)]";
            ////Debugger.Launch();
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            ////////////////////////////////////////////////

            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByXPath(editFieldxpath);
            WebDriverWait close;
            WindowsElement closeButtons;
            List<string> names = new List<string>();

            List<Point> coordenadas = coordinatesFinder(desktopSession);
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            bool IsNoteDisplayed = false;

            if (coordenadas.Count > 0)
            {

                foreach (Point coord in coordenadas)
                {
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                    desktopSession.Mouse.Click(null);
                    
                    try
                    {
                        rootSession = RootSessionNew();
                        rootSession = ReloadXSession(rootSession, "TFrmNotas");
                        Thread.Sleep(3000);
                        
                        if (rootSession == null)
                        {
                            rootSession = RootSessionNew();
                            rootSession = ReloadXSession(rootSession, "TFrmNotas");
                            Thread.Sleep(3000);
                        }
                        
                        close = new WebDriverWait(rootSession, TimeSpan.FromSeconds(3));
                        closeButtons = rootSession.FindElementByXPath(closeNote);
                        close.Until(res => closeButtons.Displayed);
                        if (closeButtons.Displayed && closeButtons.GetAttribute("Name") == "Cerrar" || closeButtons.Displayed && closeButtons.GetAttribute("Name") == "Close")
                        {
                            ScreenshotNuevo("Nota", true, file);
                            Thread.Sleep(500);
                            cerrarNotaSuperior(rootSession, closeButtons);
                        }
                        string navigator = AbrirPrograma.SelectNavigator();
                        if (navigator == "IE")
                        {
                            //Jose
                            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
                            bool ventana = newFunctions_1.coordinatesFinder(desktopSession, 27, 101, 206, screenWidth / 2, 347, 1);
                            rootSession = RootSessionNew();
                            rootSession = ReloadXSession(rootSession, "IEFrame");
                            var allFrame = rootSession.FindElementsByClassName("IEFrame");
                            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
                            rootSession = RootSessionNew();
                        }


                    }

                    catch (Exception e)
                    {
                        rootSession = RootSessionNew();                        
                        Thread.Sleep(2000);
                        close = new WebDriverWait(rootSession, TimeSpan.FromSeconds(3));
                        ScreenshotNuevo("Error en la Nota", true, file);
                        Thread.Sleep(500);
                        Console.WriteLine("Error en la nota: La Propiedad Dataset no ha sido asignada.");
                        desktopSession.Close();
                        desktopSession.Dispose();
                        //PruebaCRUD.VentanaAzul(desktopSession);

                        //closeButtons = rootSession.FindElementByXPath(closeNote);                        
                        //close.Until(res => closeButtons.Displayed);

                        //if (closeButtons.Displayed && closeButtons.GetAttribute("Name") == "Cerrar" || closeButtons.Displayed && closeButtons.GetAttribute("Name") == "Close")
                        //{
                        //    ScreenshotNuevo("Nota", true, file);
                        //    Thread.Sleep(500);
                        //    cerrarNotaSuperior(rootSession, closeButtons);

                        //}                       
                        //else
                        //{                            
                        //    ScreenshotNuevo("Error en la Nota", true, file);
                        //    Thread.Sleep(500);
                        //    PruebaCRUD.VentanaAzul(desktopSession);
                        //    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        //}
                    }

                }

            }

            string pathImg = "C:\\imagenesTest\\Programa.Png";
            if (File.Exists(pathImg))
            {
                File.Delete(pathImg);
            }

        }



        public static bool editErrorHandler(WindowsDriver<WindowsElement> desktopSession, string file) {

            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            ////////////////////////////////////////////////

            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSessionNew();
            Thread.Sleep(2000);
            WebDriverWait error = new WebDriverWait(rootSession, TimeSpan.FromSeconds(3));
            WindowsElement errorWindow = null;
            try {
                errorWindow = rootSession.FindElementByClassName("TReconcileErrorForm");
                error.Until(res => errorWindow.Displayed);
            }
            catch(Exception){ 
            
            }

            if (errorWindow != null)
            {
                ScreenshotNuevo("Error controlado al editar", true, file);
                List<string> errors = new List<string>();
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                errors = newFunctions_3.LupaAud(desktopSession, "1", file);
            }

            return errorWindow != null ? true : false;
        }


        public static void ScreenshotNuevo(string maestro, bool bandera, string file)
        {
            try

            {

                //Creating a new Bitmap object

                int screenWidth = Screen.PrimaryScreen.Bounds.Width;
                int screenHeight = Screen.PrimaryScreen.Bounds.Height;

                string name = maestro;
                string path = @"C:\Reportes\" + maestro + ".bmp";

                Bitmap captureBitmap = new Bitmap(screenWidth, screenHeight, PixelFormat.Format32bppArgb);

                Rectangle captureRectangle = Screen.AllScreens[0].Bounds;

                Graphics captureGraphics = Graphics.FromImage(captureBitmap);

                captureGraphics.CopyFromScreen(captureRectangle.Left, captureRectangle.Top, 0, 0, captureRectangle.Size);

                captureBitmap = (new Bitmap(captureBitmap, new Size(1600, 900)));

                captureBitmap.Save(string.Format("C:\\Reportes\\{0}.bmp", name), ImageFormat.Bmp);

                InsertAPicture(file, path, maestro, bandera);

            }

            catch (Exception ex)

            {

                MessageBox.Show(ex.Message);

            }
        }

        public static string ScreenshotInicio(string codProgram)
        {
            //------------- Add Daniel-----------------------------------------------------            
            string mensaje = string.Format("Error al Abrir programa;{0}", codProgram);
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            int screenHeight = Screen.PrimaryScreen.Bounds.Height;
            string name = mensaje;
            //string path = @"C:\Reportes\" + mensaje + ".bmp";

            Bitmap captureBitmap = new Bitmap(screenWidth, screenHeight, PixelFormat.Format32bppArgb);
            Rectangle captureRectangle = Screen.AllScreens[0].Bounds;
            Graphics captureGraphics = Graphics.FromImage(captureBitmap);
            captureGraphics.CopyFromScreen(captureRectangle.Left, captureRectangle.Top, 0, 0, captureRectangle.Size);

            captureBitmap = (new Bitmap(captureBitmap, new Size(1600, 900)));
            captureBitmap.Save(string.Format("C:\\Reportes\\{0}.png", name), ImageFormat.Png);
            //--------------------------------------------------------------------------------

            string path = @"C:\Reportes\" + mensaje + ".png";

            return path;
        }



        //Beta de Navegacion de Botonera
        public static void findNavigationButtons(WindowsDriver<WindowsElement> desktopSession)
        {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TKNavegador");
            editFields[0].ClearCache();
            desktopSession.Mouse.MouseMove(editFields[0].Coordinates);
            Thread.Sleep(4000);

            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> SelectedButtons = desktopSession.FindElementsByName("Insertar");
        }


        //Funciones necesarias para crud
        public static List<Point> coordinatesFinder(WindowsDriver<WindowsElement> session)
        {
            Screenshot image = ((ITakesScreenshot)session).GetScreenshot();
            string name = "Programa";
            string path = "C:\\imagenesTest\\" + name + ".Png";

            Image imgSource;
            image.SaveAsFile(string.Format("C:\\imagenesTest\\{0}.Png", name), ScreenshotImageFormat.Png);
            Thread.Sleep(1000);
            imgSource = Image.FromFile(path);

            List<Point> coordenadas = new List<Point>();
            Bitmap bmpSource = new Bitmap(imgSource);
            Bitmap bmpDest = new Bitmap(bmpSource.Width, bmpSource.Height);

            coordenadas = PixelSeeker(bmpSource, 250, 224, 103, new Point(20, 20), coordenadas);

            coordenadas = PixelSeeker(bmpSource, 251, 245, 199, new Point(20, 20), coordenadas);

            imgSource.Dispose();

            return coordenadas;


        }


        public static List<Point> PixelSeeker(Bitmap image, int R, int G, int B, Point offset, List<Point> coordinates)
        {

            Point punto = new Point();
            int lastx = 0;
            int lasty = 0;

            for (int x = 0; x < image.Width; x++)
            {
                for (int y = 0; y < image.Height; y++)
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



            return coordinates;

        }


        private static void abrirNotaSuperior(WindowsDriver<WindowsElement> session, WindowsElement noteButton, int x, int y)
        {
            bool isShowingNote = false;

            Thread.Sleep(1500);
            noteButton.ClearCache();
            noteButton.DisableCache();
            session.Mouse.MouseMove(noteButton.Coordinates, x, y);
            session.Mouse.Click(null);


        }


        private static void cerrarNotaSuperior(WindowsDriver<WindowsElement> session, WindowsElement closeB)
        {

            string name = closeB.GetAttribute("Name");

            if (closeB.Displayed && name == "Cerrar" || closeB.Displayed && name == "Close")
            {

                closeB.ClearCache();
                closeB.DisableCache();
                closeB.Click();
            }


        }


        public static void findText(WindowsDriver<WindowsElement> desktopSession, List<string> val, string className)
        {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName(className);

            int i = 0;

            foreach (var elem in editFields)
            {

                elem.ClearCache();
                desktopSession.Mouse.MouseMove(elem.Coordinates);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(val[i]);
                i++;

                Thread.Sleep(5000);
            }

        }



      
        public static void findCheckBoxes(WindowsDriver<WindowsElement> desktopSession, string checkkBoxName)
        {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            ////////////////////////////////////////////////
            WebDriverWait printScreen;
            WindowsElement printerField;
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editCheckBoxes = desktopSession.FindElementsByClassName("TCheckBox");

            foreach (var elem in editCheckBoxes)
            {
                if (elem.Text == checkkBoxName)
                {
                    elem.Click();
                }
            }


        }

       

        public static List<string> BotonAceptar(WindowsDriver<WindowsElement> desktopSession, string ruta, string file)
        {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            ////////////////////////////////////////////////
            WebDriverWait printScreen;
            WindowsElement printerField;
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> botones = desktopSession.FindElementsByClassName("TBitBtn");
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();

            foreach (var elem in botones)
            {
                if (elem.Text == "Aceptar")
                {
                    elem.Click();
                    Thread.Sleep(2000);
                    try
                    {
                        errors = newFunctions_5.GenerarReportes("Aceptar", file, ruta);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                    }
                    catch {
                        Thread.Sleep(3000);
                        ScreenshotNuevo("Error al generar preliminar", true, file);
                        rootSession = PruebaCRUD.RootSession();
                        Thread.Sleep(10000);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(2000);
                    }

                  
                }
            }

            return errorMessages;

        }


        public static List<string> ReportePreliminarIfnotOptions(WindowsDriver<WindowsElement> desktopSession, string ruta, string file, string buttonName, bool isDateNeccessary, List<string> dates = null) {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WebDriverWait printScreen;
            WindowsElement printerField;
            WindowsDriver<WindowsElement> rootSession = null;



            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            ////////////////////////////////////////////////


            Thread.Sleep(2000);
            List<Point> coordenadas = FuncionesGlobales.coordinatesFinder(desktopSession, 227, 117, 0);
            int coordX = coordenadas[0].X;
            int coordY = coordenadas[0].Y;
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            Thread.Sleep(1000);
            try
            {
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                desktopSession.Mouse.Click(null);
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }

            rootSession = RootSessionNew();
            Thread.Sleep(5000);
            WebDriverWait reporteWindow = new WebDriverWait(rootSession, TimeSpan.FromSeconds(4));


            //ventana de reporte preliminar: "/Pane[@ClassName=\"#32769\"][@Name=\"Escritorio 1\"]/Window[@ClassName=\"TFrmBiProxExa\"][@Name=\"Reportes\"]"
            //radio butons: /RadioButton[@ClassName=\"TGroupButton\"][@Name=\"Próximo Examen\"]"
            //fechasfields: "/Pane[@ClassName=\"#32769\"][@Name=\"Escritorio 1\"]/Window[@ClassName=\"TFrmBiProxExa\"][@Name=\"Reportes\"]/Edit[@ClassName=\"TEdit\"]"
            //boton Aceptar: "/Pane[@ClassName=\"#32769\"][@Name=\"Escritorio 1\"]/Window[@ClassName=\"TFrmBiProxExa\"][@Name=\"Reportes\"]/Button[@ClassName=\"TBitBtn\"][@Name=\"Aceptar\"]"

            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> radioButton = rootSession.FindElementsByClassName("TGroupButton");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> aceptButton = rootSession.FindElementsByName("Aceptar");
            reporteWindow.Until(res => aceptButton[0].Displayed);

            //foreach (var but in radioButton) {
            //    rootSession.Mouse.MouseMove(but.Coordinates);
            //    Thread.Sleep(2000);
            //}

            foreach (var elem in radioButton)
            {
                if (elem.Text == buttonName && isDateNeccessary == false)
                {
                    rootSession.Mouse.MouseMove(elem.Coordinates);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    ScreenshotNuevo(buttonName, true, file);
                    rootSession.Mouse.MouseMove(aceptButton[0].Coordinates);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(3000);
                    try
                    {
                        errors = newFunctions_5.GenerarReportes("Preliminar", file, ruta);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                    }
                    catch
                    {
                        newFunctions_4.ScreenshotNuevo("Error en Preliminar", true, file);
                        Thread.Sleep(1000);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    }
                }
                else
                {
                    if (elem.Text == buttonName && isDateNeccessary == true && dates != null)
                    {
                        rootSession.Mouse.MouseMove(elem.Coordinates);
                        rootSession.Mouse.Click(null);
                        System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> dateField = rootSession.FindElementsByClassName("TEdit");
                        int indexDates = dates.Count - 1;
                        foreach (var field in dateField)
                        {

                            rootSession.Mouse.MouseMove(field.Coordinates);
                            rootSession.Mouse.Click(null);
                            field.Clear();
                            rootSession.Keyboard.SendKeys(dates[indexDates]);
                            Thread.Sleep(1000);
                            indexDates--;
                        }

                        rootSession.Mouse.MouseMove(aceptButton[0].Coordinates);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(3000);
                        try
                        {
                            errors = newFunctions_5.GenerarReportes("Preliminar", file, ruta);
                            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        }
                        catch
                        {
                            newFunctions_4.ScreenshotNuevo("Error en Preliminar", true, file);
                            Thread.Sleep(1000);
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        }

                    }
                }
            }
            return errorMessages;
        }


        public static List<string> ReportePreliminarIfnot(WindowsDriver<WindowsElement> desktopSession, string ruta, string file) {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WebDriverWait printScreen;
            WindowsElement printerField;
            WindowsDriver<WindowsElement> rootSession = null;


            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            Thread.Sleep(2000);
            List<Point> coordenadas = FuncionesGlobales.coordinatesFinder(desktopSession, 227, 117, 0);
            int coordX = coordenadas[0].X;
            int coordY = coordenadas[0].Y;
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            Thread.Sleep(1000);
            try
            {
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                desktopSession.Mouse.Click(null);
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }

            rootSession = RootSessionNew();
            Thread.Sleep(5000);
            WebDriverWait reporteWindow = new WebDriverWait(rootSession, TimeSpan.FromSeconds(4));


            try
            {
                errors = newFunctions_5.GenerarReportes("Preliminar", file, ruta);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            catch {
                newFunctions_4.ScreenshotNuevo("Error en Preliminar", true, file);
                Thread.Sleep(1000);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }

            return errorMessages;

        }


        public static List<string> ReportePreliminarOpciones(WindowsDriver<WindowsElement> desktopSession, string ruta, string buttonName, string file, bool isDateNeccessary, List<string> dates = null)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WebDriverWait printScreen;
            WindowsElement printerField;
            WindowsDriver<WindowsElement> rootSession = null;

          
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            ////////////////////////////////////////////////
                
               
            Thread.Sleep(2000);
            List<Point> coordenadas = FuncionesGlobales.coordinatesFinder(desktopSession, 227, 117, 0);
            int coordX = coordenadas[0].X;
            int coordY = coordenadas[0].Y;
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            Thread.Sleep(1000);
            try
            {
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                desktopSession.Mouse.Click(null);
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }

            rootSession = RootSessionNew();
            Thread.Sleep(5000);
            WebDriverWait reporteWindow = new WebDriverWait(rootSession, TimeSpan.FromSeconds(4));
            

            //ventana de reporte preliminar: "/Pane[@ClassName=\"#32769\"][@Name=\"Escritorio 1\"]/Window[@ClassName=\"TFrmBiProxExa\"][@Name=\"Reportes\"]"
            //radio butons: /RadioButton[@ClassName=\"TGroupButton\"][@Name=\"Próximo Examen\"]"
            //fechasfields: "/Pane[@ClassName=\"#32769\"][@Name=\"Escritorio 1\"]/Window[@ClassName=\"TFrmBiProxExa\"][@Name=\"Reportes\"]/Edit[@ClassName=\"TEdit\"]"
            //boton Aceptar: "/Pane[@ClassName=\"#32769\"][@Name=\"Escritorio 1\"]/Window[@ClassName=\"TFrmBiProxExa\"][@Name=\"Reportes\"]/Button[@ClassName=\"TBitBtn\"][@Name=\"Aceptar\"]"

            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> radioButton = rootSession.FindElementsByClassName("TGroupButton");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> aceptButton = rootSession.FindElementsByName("Aceptar");
            reporteWindow.Until(res => aceptButton[0].Displayed);

            //foreach (var but in radioButton) {
            //    rootSession.Mouse.MouseMove(but.Coordinates);
            //    Thread.Sleep(2000);
            //}

            foreach (var elem in radioButton)
            {
                if (elem.Text == buttonName && isDateNeccessary == false)
                {
                    rootSession.Mouse.MouseMove(elem.Coordinates);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    ScreenshotNuevo(buttonName, true, file);
                    rootSession.Mouse.MouseMove(aceptButton[0].Coordinates);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(3000);
                    errors = newFunctions_5.GenerarReportes("Preliminar", file, ruta);
                    if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                }
                else {
                    if (elem.Text == buttonName && isDateNeccessary == true && dates != null) {
                        rootSession.Mouse.MouseMove(elem.Coordinates);
                        rootSession.Mouse.Click(null);
                        System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> dateField = rootSession.FindElementsByClassName("TEdit");
                        int indexDates = dates.Count - 1;
                        foreach (var field in dateField) {
                            
                            rootSession.Mouse.MouseMove(field.Coordinates);
                            rootSession.Mouse.Click(null);
                            field.Clear();
                            rootSession.Keyboard.SendKeys(dates[indexDates]);
                            Thread.Sleep(1000);
                            indexDates--;
                        }

                        rootSession.Mouse.MouseMove(aceptButton[0].Coordinates);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(3000);
                        errors = newFunctions_5.GenerarReportes("Preliminar", file, ruta);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        

                    }
                }
            }
            return errorMessages;
        }

        public static void changeWindow(WindowsDriver<WindowsElement> desktopSession, string name, int index = 0)
        {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByName(name);

            desktopSession.Mouse.MouseMove(editFields[index].Coordinates, editFields[index].Size.Width / 2, editFields[index].Size.Height / 2);
            Thread.Sleep(3000);
            desktopSession.Mouse.Click(null);
        }

        public static List<Point> CoordinatesFinder2(WindowsDriver<WindowsElement> session, int R, int G, int B, Point offset)
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
            coordenadas = PruebaCRUD.PixelSeeker(bmpSource, R, G, B, offset);

            //if (coordenadas.Count == 0) coordenadas = PixelSeeker(bmpSource, R, G, B, new Point(20, 20));

            imgSource.Dispose();

            return coordenadas;

        }


        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data, Point offset, List<int> lupasOffIndex = null) {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = CoordinatesFinder2(desktopSession, 82, 150, 75, new Point(20,20));
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSessionNew();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 0;

            if (lupasOffIndex != null)
            {
                List<Point> backupList = new List<Point>();
                for (int i = 0; i < coordenadas.Count; i++)
                {
                    if (!lupasOffIndex.Contains(i)) backupList.Add(coordenadas[i]);
                }
                coordenadas = backupList;
            }

            for (int coordIndex = 0; coordIndex < coordenadas.Count; coordIndex++)
            {
                Thread.Sleep(2000);
                if (coordIndex == 1) Thread.Sleep(3000);

                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordenadas[coordIndex].X + offset.X, coordenadas[coordIndex].Y + offset.Y);
                desktopSession.Mouse.Click(null);
                if (contLupa == 1)
                {
                    try
                    {
                        rootSession = RootSessionNew();
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

                    rootSession = RootSessionNew();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                    //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    rootSession.Keyboard.SendKeys(data[indexVal]);
                    //Aqui se puede enviar un valor

                    var btn = rootSession.FindElementsByClassName("TBitBtn");
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



        public static void LupaDinamicaDiscriminatoria2(WindowsDriver<WindowsElement> desktopSession, List<string> data, Point offset, List<int> lupasOffIndex = null)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSessionNew();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 0;

            if (lupasOffIndex != null)
            {
                List<Point> backupList = new List<Point>();
                for (int i = 0; i < coordenadas.Count; i++)
                {
                    if (!lupasOffIndex.Contains(i)) backupList.Add(coordenadas[i]);
                }
                coordenadas = backupList;
            }

            for (int coordIndex = 0; coordIndex < coordenadas.Count; coordIndex++)
            {
                Thread.Sleep(2000);
                if (coordIndex == 1) Thread.Sleep(3000);

                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordenadas[coordIndex].X + offset.X, coordenadas[coordIndex].Y + offset.Y);
                desktopSession.Mouse.Click(null);
                if (contLupa == 1)
                {
                    try
                    {
                        rootSession = RootSessionNew();
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

                    rootSession = RootSessionNew();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                    //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    rootSession.Keyboard.SendKeys(data[indexVal]);
                    //Aqui se puede enviar un valor

                    var btn = rootSession.FindElementsByClassName("TBitBtn");
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



        public static void LupaDinamicaDiscriminatoria(WindowsDriver<WindowsElement> desktopSession, List<string> data, List<int> lupasOffIndex = null)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSessionNew();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 0;

            if (lupasOffIndex != null) {
                List<Point> backupList = new List<Point>();
                for (int i = 0; i < coordenadas.Count; i++) {
                    if (!lupasOffIndex.Contains(i)) backupList.Add(coordenadas[i]);
                }
                coordenadas = backupList;
            }

            for(int coordIndex = 0; coordIndex < coordenadas.Count; coordIndex++)
            {
                Thread.Sleep(2000);
                if (coordIndex == 1) Thread.Sleep(3000);
                
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordenadas[coordIndex].X + 10, coordenadas[coordIndex].Y);
                desktopSession.Mouse.Click(null);
                if (contLupa == 1)
                {
                    try
                    {
                        rootSession = RootSessionNew();
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

                    rootSession = RootSessionNew();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                    //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    rootSession.Keyboard.SendKeys(data[indexVal]);
                    //Aqui se puede enviar un valor

                    var btn = rootSession.FindElementsByClassName("TBitBtn");
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


        public static List<string> validarCodPrograma(WindowsDriver<WindowsElement> desktopSession, string programa)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSessionNew();
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(rootSession)
            {
                Timeout = TimeSpan.FromSeconds(60),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            var ElementList = desktopSession.FindElementsByClassName("TPanel");
            var obtenerLetrasProg = programa.Substring(0, 3).ToLower();

            foreach (var elem in ElementList)
            {
                if (elem.Text.Length > 0)
                {
                    if(elem.Text.Length > 3)
                    {
                        var obtenerLetrasTpanel = elem.Text.Substring(0, 3).ToLower();
                        if (obtenerLetrasProg == obtenerLetrasTpanel)
                        {
                            if (elem.Text == programa)
                            {
                                Debug.WriteLine("Codigo del programa coincide");
                            }
                            else
                            { 
                                errorMessages.Add($"Codigo de programa es diferente al esperado. Esperado: {programa}, Encontrado: {elem.Text}");
                            }
                            break;
                        }
                    }

                }
            }
            return errorMessages;
        }


        public static List<string> validarDescripPrograma(WindowsDriver<WindowsElement> desktopSession, string descripcionDada)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSessionNew();
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(60),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            var ElementList = desktopSession.FindElementsByClassName("Internet Explorer_Server");
            string Texto = ElementList[0].Text.Remove(0,54);

            //Encontrar posicion inicio
            int posicion0 = Texto.IndexOf("&") + 1;
            int posicion1 = Texto.IndexOf("&.");
 
            ////Obtener nombre
            string descripcionObte = Texto.Substring(posicion0, posicion1 - posicion0);

            //Limpiando la descripcion del programa obtenido
            string descripcionObteMostrar = descripcionObte.Replace("%20", " ");
            descripcionObte = descripcionObte.Replace("%20", "");
            descripcionObte = descripcionObte.ToLower();

            //Limpiando la descripcion del programa dado
            string descripcionDadaMostrar = descripcionDada;
            descripcionDada = descripcionDada.Replace(" ", "");
            descripcionDada = descripcionDada.ToLower();

            if (descripcionDada == descripcionObte)
            {
                Debug.WriteLine("Descripcion del programa coincide");
            }
            else
            {
                errorMessages.Add($"Descripción del programa es diferente al esperado. Esperado: {descripcionDadaMostrar}, Encontrado: {descripcionObteMostrar}");
            }
            return errorMessages;
        }        
    }
}

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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosCo
{
    class CrudKcoacmej:FuncionesVitales
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
        //ClickWindow
        static public void ClickWindow(WindowsDriver<WindowsElement> desktopSession)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            //TPanel
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 61, 20);
            Thread.Sleep(2000);
            desktopSession.Mouse.Click(null);
        }
        //Click
        static public void Click(WindowsDriver<WindowsElement> desktopSession)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("FrameGrabHandle");
            //TPanel
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 480, 21);
            Thread.Sleep(2000);
            desktopSession.Mouse.Click(null);
        }
        //Click
        public static List<Point> coordinatesFinder(WindowsDriver<WindowsElement> session)
        {
            Screenshot image = ((ITakesScreenshot)session).GetScreenshot();
            string name = "ProgramaExcel";
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
                    if (clrPixel.R == 151 && clrPixel.G == 179 && clrPixel.B == 143)
                    {
                        puntos.X = x;
                        puntos.Y = y;
                        if (x > lastx + 10 || y > lasty + 10)
                        {
                            coordenadas.Add(puntos);
                            lastx = x;
                            lasty = y;
                            romper = 1;
                            break;
                        }
                    }
                }
                if (romper == 1)
                {
                    break;
                }
            }
            return coordenadas;
        }
        static public WindowsDriver<WindowsElement> ReloadSession1(WindowsDriver<WindowsElement> session, string parametro)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(parametro);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }
        public static List<string> generarExcel(WindowsDriver<WindowsElement> Session, string file, string codProgram, string ruta)
        {
            List<string> errorMessages = new List<string>();
            try
            {
                Session = RootSession();
                Session = ReloadSession1(Session, "TFrmExcelex");
                var element = Session.FindElementsByXPath("//Window[contains(@ClassName, 'TFrmExcelex')]");
                if (element[0] != null)
                {
                    bool IsDisplayedExcel = false;
                    string Fecha = Hora();
                    Thread.Sleep(500);
                    if (element[0] != null)
                    {
                        IsDisplayedExcel = true;
                    }
                    else
                    {
                        errorMessages.Add($"No se puede visualizar la ventana Exportar a excel");
                    }

                    if (IsDisplayedExcel)
                    {
                        Session = RootSession();
                        Session = ReloadSession1(Session, "TFrmExcelex");
                        var ReportElements = Session.FindElementsByXPath("//Window[contains(@ClassName, 'TFrmExcelex')]");
                        Session.Mouse.MouseMove(ReportElements[0].Coordinates, 222, 207);
                        Session.Mouse.Click(null);
                        newFunctions_4.ScreenshotNuevo("Exportar datos Excel", true, file);
                        var buttonAccept = Session.FindElementByName("Aceptar");
                        buttonAccept.Click();

                        Thread.Sleep(5000);
                        try
                        {
                            Session = RootSession();
                            Session = ReloadSession1(Session, "XLMAIN");
                            if (Session != null)
                            {
                                int count = 0;
                                while (count < 240)
                                {
                                    try
                                    {
                                        //Session = RootSession();
                                        //Session = ReloadSession1(Session, "XLMAIN");
                                        Thread.Sleep(500);
                                        var saveExcel1 = Session.FindElementsByName("Maximizar");
                                        if (saveExcel1.Count > 0)
                                        {
                                            break;
                                        }
                                    }
                                    catch
                                    {
                                        Thread.Sleep(1000);
                                        count++;
                                    }
                                }
                                var saveExcel = Session.FindElementsByName("Maximizar");
                                var cantidad = saveExcel.Count;
                                if (cantidad == 2)
                                {
                                    saveExcel[1].Click();
                                }
                                else
                                {
                                    saveExcel[0].Click();
                                }
                                //Cambio Jose
                                Session.FindElementByName("Pestaña Archivo").Click();
                                //Fin Cambio Jose

                                //Session.FindElementByName("Guardar").Click();
                                Session.FindElementByName("Guardar").Click();
                                Session.FindElementByName("Examinar").Click();
                                Thread.Sleep(1000);
                                Session.Keyboard.SendKeys(ruta);
                                Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                Thread.Sleep(1000);
                                newFunctions_4.ScreenshotNuevo("Reporte Excel Guardado", true, file);
                                LimpiarProcesos();

                            }
                            else
                            {
                                //Debugger.Launch();
                                Thread.Sleep(1000);
                                Session = RootSession();
                                Session = ReloadSession1(Session, "Shell_TrayWnd");
                                Thread.Sleep(1000);
                                var venExcel = Session.FindElementByName("Excel - 1 ventana de ejecución");
                                if (venExcel != null)
                                {
                                    venExcel.Click();
                                    Thread.Sleep(4000);
                                    int count = 0;
                                    while (count < 240)
                                    {
                                        try
                                        {
                                            Session = RootSession();
                                            Session = ReloadSession1(Session, "XLMAIN");
                                            Thread.Sleep(500);
                                            var saveExcel1 = Session.FindElementsByName("Maximizar");
                                            if (saveExcel1.Count > 0)
                                            {
                                                break;
                                            }
                                        }
                                        catch
                                        {
                                            Thread.Sleep(1000);
                                            count++;
                                        }
                                    }
                                    var saveExcel = Session.FindElementsByName("Maximizar");
                                    var cantidad = saveExcel.Count;
                                    if (cantidad == 2)
                                    {
                                        saveExcel[1].Click();
                                    }
                                    else
                                    {
                                        saveExcel[0].Click();
                                    }
                                    //Cambio Jose
                                    Session.FindElementByName("Pestaña Archivo").Click();
                                    //Fin Cambio Jose

                                    //Session.FindElementByName("Guardar").Click();
                                    Session.FindElementByName("Guardar").Click();
                                    Session.FindElementByName("Examinar").Click();
                                    Thread.Sleep(1000);
                                    Session.Keyboard.SendKeys(ruta);
                                    Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                    Thread.Sleep(1000);
                                    newFunctions_4.ScreenshotNuevo("Reporte Excel Guardado", true, file);
                                    LimpiarProcesos();
                                }
                            }

                        }
                        catch
                        {
                            Thread.Sleep(1000);
                            /*string navigator = AbrirPrograma.SelectNavigator();
                            if (navigator == "Edge") 
                            {
                                newFunctions_4.ScreenshotNuevo("Opción de Excel no trae registros", true, file);
                                PruebaCRUD.VentanaAzul(Session);
                                Thread.Sleep(2000);
                            }*/
                        }

                        if (Session == null)
                        {
                            Thread.Sleep(1000);
                            newFunctions_4.ScreenshotNuevo("Opción de Excel no trae registros", true, file);
                            Thread.Sleep(1000);
                            Session = RootSession();
                            PruebaCRUD.VentanaAzul(Session);
                            Thread.Sleep(500);
                        }

                    }
                }
            }
            catch
            {
                //Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                errorMessages.Add($"Error al intentar generar el excel");
                newFunctions_4.ScreenshotNuevo("Error en el reporte excel", true, file);
                Thread.Sleep(500);
                //rootSession = RootSession();
                //rootSession = ReloadSession(rootSession, "Chrome_WidgetWin_1");
                //rootSession.Close();
                //rootSession.Dispose();
                //ext = 1;
                //break;
                string varNavigator = null;
                varNavigator = "msedge";
                Process[] processes = Process.GetProcessesByName(varNavigator);

                if (processes.Length > 0)
                {
                    for (int j = 0; j < processes.Length; j++)
                    {
                        processes[j].Kill();
                    }
                }
                return errorMessages;
            }

            return errorMessages;
        }
        public static List<string> ExporExcel(WindowsDriver<WindowsElement> desktopSession, string banExcel, string file, string codProgram, string ruta)
            {
                List<string> errorMessages = new List<string>();
                var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
                {
                    Timeout = TimeSpan.FromSeconds(10),
                    PollingInterval = TimeSpan.FromSeconds(1)
                };
                wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
                //Debugger.Launch();
                try
                {
                    wait.Until(driver =>
                    {
                        WindowsDriver<WindowsElement> rootSession = null;
                        List<Point> coordenadas = coordinatesFinder(desktopSession);
                        int coordX = coordenadas[0].X;
                        int coordY = coordenadas[0].Y;
                        var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));

                        string navigator = AbrirPrograma.SelectNavigator();
                        //EDEGE
                        var ExcelOption1 = desktopSession.FindElementsByClassName("TKNavegador");
                        if (navigator == "IE")
                        {
                            //IE
                            string ClickButtonExcel = "//Pane[contains(@ClassName, 'TKNavegador')]/..";
                            ExcelOption1 = desktopSession.FindElementsByXPath(ClickButtonExcel);
                        }

                        if (ExcelOption1.Count > 0)
                        {
                            desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                            desktopSession.Mouse.Click(null);
                            newFunctions_4.ScreenshotNuevo("Opciones Exportar Excel", true, file);
                            bool optionExcel = false;
                            int cont = 0;
                            string clase = "";
                            if (banExcel == "0")
                            {
                                Thread.Sleep(1000);
                                try
                                {
                                    generarExcel(rootSession, file, codProgram, ruta);
                                    rootSession = RootSession();
                                }
                                catch
                                {
                                    PruebaCRUD.VentanaAzul(desktopSession);
                                    rootSession = RootSession();
                                }
                            }
                            else if (banExcel == "1")
                            {
                                rootSession = RootSession();
                                rootSession = ReloadSession1(rootSession, "TForm");
                            }
                            else if (banExcel == "2")
                            {
                                rootSession = RootSession();
                                rootSession = ReloadSession1(rootSession, "TFrmMenuExcel");
                            }
                            else if (banExcel == "3")
                            {

                                rootSession = RootSession();
                                rootSession = ReloadSession1(rootSession, "TFrmMenExpo");
                            }
                            else if (banExcel == "4")
                            {

                                rootSession = RootSession();
                                rootSession = ReloadSession1(rootSession, "TFrmTipExpo");
                            }

                            if (banExcel == "1" || banExcel == "2" || banExcel == "3" || banExcel == "4")
                            {
                                ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                var contButton = rootSession.FindElementsByClassName("TGroupButton");
                                int cont2 = contButton.Count() - 1;
                                rootSession.FindElementsByClassName("TGroupButton");
                                rootSession.FindElementsByClassName("TGroupButton")[cont2].Click();
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                try
                                {
                                    generarExcel(rootSession, file, codProgram, ruta);
                                    
                                }
                                catch
                                {
                                    Thread.Sleep(2000);
                                    PruebaCRUD.VentanaAzul(desktopSession);
                                    Thread.Sleep(2000);
                                }
                                int i;
                                
                                //Debugger.Launch();
                                for (i = cont2 - 1; i > -1; i--)
                                {
                                    i = cont2 - 1;
                                    ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                    desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);                                    
                                    desktopSession.Mouse.Click(null);
                                    if (banExcel == "1")
                                    {
                                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                    }
                                    else
                                    {                                        
                                        rootSession.FindElementsByClassName("TGroupButton")[cont2 /*- 1*/].Click();
                                    }
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                    Thread.Sleep(1000);
                                    newFunctions_4.ScreenshotNuevo("Opción Excel", true, file);
                                    Thread.Sleep(1000);
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                    Thread.Sleep(1000);
                                    try
                                    {
                                        generarExcel(rootSession, file, codProgram, ruta);
                                    }
                                    catch
                                    {
                                        Thread.Sleep(7000);
                                        PruebaCRUD.VentanaAzul(desktopSession);
                                        Thread.Sleep(2000);
                                    }
                                    cont2 = cont2 - 1;
                                    break;

                                }
                            }

                            if (banExcel == "5")
                            {
                                rootSession = RootSession();
                                rootSession = ReloadSession1(rootSession, "TForm");
                                Thread.Sleep(1000);
                                rootSession.FindElementByName("Exportación estándar").Click();
                                Thread.Sleep(1000);
                                newFunctions_4.ScreenshotNuevo("Opción Excel", true, file);
                                Thread.Sleep(1000);
                                rootSession.FindElementByName("Aceptar").Click();
                                Thread.Sleep(4000);
                                rootSession = RootSession();
                                rootSession = ReloadSession1(rootSession, "XLMAIN");
                                Thread.Sleep(3000);
                                //Debugger.Launch();
                                if (rootSession == null)
                                {
                                    Thread.Sleep(1000);
                                    newFunctions_4.ScreenshotNuevo("No hay registros para exportar o el generar excel", true, file);
                                    Thread.Sleep(2000);
                                    PruebaCRUD.VentanaAzul(desktopSession);
                                    Thread.Sleep(2000);
                                }
                                else
                                {
                                    generarExcel(rootSession, file, codProgram, ruta);
                                }
                                for (int i = 1; i < 16; i++)
                                {
                                    //Clic Excel
                                    ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                    desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                                    desktopSession.Mouse.Click(null);
                                    Thread.Sleep(1000);
                                    rootSession = RootSession();
                                    rootSession = ReloadSession1(rootSession, "TForm");
                                    Thread.Sleep(1000);
                                    rootSession.FindElementByName("Exportación estándar").Click();
                                    Thread.Sleep(1000);
                                    for (int j = 0; j < i; j++)
                                    {
                                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                    }
                                    Thread.Sleep(1000);
                                    newFunctions_4.ScreenshotNuevo("Opción Excel", true, file);
                                    Thread.Sleep(1000);
                                    rootSession.FindElementByName("Aceptar").Click();
                                    Thread.Sleep(4000);
                                    rootSession = RootSession();
                                    rootSession = ReloadSession1(rootSession, "XLMAIN");
                                    if (rootSession == null)
                                    {
                                        Thread.Sleep(1000);
                                        newFunctions_4.ScreenshotNuevo("No hay registros para exportar o el generar excel", true, file);
                                        Thread.Sleep(2000);
                                        PruebaCRUD.VentanaAzul(desktopSession);
                                        Thread.Sleep(2000);
                                    }
                                    else
                                    {
                                        generarExcel(rootSession, file, codProgram, ruta);
                                    }

                                }
                            }

                        }
                        return rootSession != null;
                    });
                }

                catch
                {
                    Debug.WriteLine("No se encuentra la opcion de excel para este programa o no hay registros para exportar");

                }
                string pathImg = "C:\\imagenesTest\\Programa.Png";
               //if (File.Exists(pathImg))
                //{
                    //File.Delete(pathImg);
                //}
                return errorMessages;
            }



        
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
           

            var ElementList = desktopSession.FindElementsByClassName("TDBCmbCodigoBox");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBComboBox");
            Debug.WriteLine($"Cantidad de elementos Lista 1 {ElementList.Count}");
            Debug.WriteLine($"Cantidad de elementos Lista 2 {ElementList.Count}");

            if (tipo == 1)
            {
                LupaDinamica(desktopSession, data[0],1);
                ElementList[0].Click();
                ElementList[0].SendKeys(data[2]);
                ElementList[0].SendKeys(OpenQA.Selenium.Keys.Enter);
                LupaDinamica(desktopSession, data[1],2);
                newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);
            }
            else
            {
                ElementList2[0].Click();
                ElementList2[0].SendKeys(data[3]);
                ElementList2[0].SendKeys(OpenQA.Selenium.Keys.Enter);
                newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);
            }
        }

        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, string data,int ban)
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
                if (contLupa == ban)
                {
                    desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                    desktopSession.Mouse.Click(null);
                    rootSession = RootSession();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                    //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    ////Debugger.Launch();
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    rootSession.Keyboard.SendKeys(data);
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
                }
                contLupa++;
            }

        }


    }
}

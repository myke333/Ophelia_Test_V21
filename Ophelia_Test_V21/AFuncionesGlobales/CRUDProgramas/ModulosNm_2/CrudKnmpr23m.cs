using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales;
using OpenQA.Selenium;
using System.Drawing;
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmpr23m : FuncionesVitales
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

        public static List<string> ReportePreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
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

                    rootSession = RootSession();
                    //rootSession = ReloadSession(rootSession, "IEFrame");
                    //var allFrame = rootSession.FindElementsByClassName("IEFrame");
                    //new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, (allFrame[0].Size.Height / 2) + 30).DoubleClick().Click().Perform();
                    //rootSession = RootSession();

                    var btn1 = rootSession.FindElementsByClassName("TScrollBox");

                    var btn = rootSession.FindElementsByClassName("TGroupButton");
                    //Debugger.Launch();
                    //foreach (var boton in btn)
                    //{
                    //    if (boton.Text == "Aceptar")
                    //    {
                    //        boton.Click();
                    //    }
                    //}
                    desktopSession.Mouse.MouseMove(btn1[0].Coordinates, 763, 225);
                    Thread.Sleep(2000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    desktopSession.Mouse.Click(null);

                    Thread.Sleep(2000);
                    if (bandera == "0")
                    {
                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                    }
                }
                catch (Exception ex)
                {
                    errorMessages.Add($"Error de Appium: {ex.ToString()}");
                }
                codPanel = parentElement;
                return codPanel != null;
            });
            return errorMessages;
        }

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
        {
            List<string> errorMessages = new List<string>();
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            if (tipo == 1)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 77, 42);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudVariables[0]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudVariables[0]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudVariables[1]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudVariables[2]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);

            }

            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 77, 42);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudVariables[3]);
                Thread.Sleep(1000);



            }
        }
        //public static void ReportePreliminar(WindowsDriver<WindowsElement> desktopSession, string ban, string file, string pdf1)
        //{
        //    List<string> errorMessages = new List<string>();
        //    WindowsDriver<WindowsElement> Session = null;
        //    Session = RootSession();
        //    Session = ReloadSession(Session, "TFrmActividad");
        //    Thread.Sleep(500);
        //    if (tipo == 1)
        //    {
        //        var ElementList = Session.FindElementByName("Activos");
        //        Thread.Sleep(2000);
        //        Session.Mouse.MouseMove(ElementList.Coordinates, 8, 12);
        //        Session.Mouse.Click(null);
        //    }
        //    else if (tipo == 2)
        //    {
        //        var ElementList = Session.FindElementByName("Inactivos");
        //        Thread.Sleep(2000);
        //        Session.Mouse.MouseMove(ElementList.Coordinates, 11, 13);
        //        Session.Mouse.Click(null);
        //    }
        //    else if (tipo == 3)
        //    {
        //        var ElementList = Session.FindElementByName("Pensionados");
        //        Thread.Sleep(2000);
        //        Session.Mouse.MouseMove(ElementList.Coordinates, 6, 23);
        //        Session.Mouse.Click(null);
        //    }
        //    else if (tipo == 4)
        //    {
        //        var ElementList = Session.FindElementByName("Retirados");
        //        Thread.Sleep(2000);
        //        Session.Mouse.MouseMove(ElementList.Coordinates, 9, 13);
        //        Session.Mouse.Click(null);
        //    }
        //    else
        //    {
        //        var ElementList = Session.FindElementByName("Todos");
        //        Thread.Sleep(2000);
        //        Session.Mouse.MouseMove(ElementList.Coordinates, 8, 14);
        //        Session.Mouse.Click(null);
        //    }
        //    Thread.Sleep(1000);
        //}
    }
}

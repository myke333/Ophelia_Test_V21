using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Drawing;
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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_7
{
    class CrudKnmmcpar: FuncionesVitales
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
        public static void cerrarBorrar(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "IEFrame");
            var allFrame = rootSession.FindElementsByClassName("IEFrame");
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
            rootSession = RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        }
        public static void EliminarRegistro(WindowsDriver<WindowsElement> desktopSession, WindowsElement ElementList, Dictionary<string, Point> botonesNavegador, string file)
        {
            desktopSession.Mouse.MouseMove(ElementList.Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Delete record", true, file);
            cerrarBorrar(desktopSession);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
            Thread.Sleep(2000);
        }

        public static void OPCIONES(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist = desktopSession.FindElementsByName("Generales");
            var Elementlist1 = desktopSession.FindElementsByName("Temporalidad");
            var Elementlist2 = desktopSession.FindElementsByName("Subtipo de comisiones");
            var Elementlist3 = desktopSession.FindElementsByName("Variables");
            switch (tipo)
            {
                case ("Generales"):
                    desktopSession.Mouse.MouseMove(Elementlist[0].Coordinates, 25, 4);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    break;
                case ("Temporalidad"):
                    desktopSession.Mouse.MouseMove(Elementlist1[0].Coordinates, 42, 4);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    break;
                case ("SubComisiones"):
                    desktopSession.Mouse.MouseMove(Elementlist2[0].Coordinates, 54, 4);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    break;
                case ("Variables"):
                    desktopSession.Mouse.MouseMove(Elementlist3[0].Coordinates, 28, 4);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    break;
            }
        }

        public static void CRUD_DETALLE1(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist = desktopSession.FindElementsByName("Generales");
            var Elementlist1 = desktopSession.FindElementsByName("Abajo");
            var Elementlist2 = desktopSession.FindElementsByName("Arriba");
           
            switch (tipo)
            {
                case ("0"):
                    desktopSession.Mouse.MouseMove(Elementlist[0].Coordinates, 25, 4);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    Elementlist1[0].Click();
                    break;
                case ("1"):
                    Elementlist2[0].Click();
                    break;
            }
        }

        public static void CRUD_DETALLE2(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist = desktopSession.FindElementsByClassName("TKactusDBgrid");
            switch (tipo)
            {
                case ("0"):
                    desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 254, 30);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(data[0]);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 434, 30);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(data[1]);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 556, 30);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(data[2]);
                    Thread.Sleep(1000);
                    break;
                case ("1"):
                    desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 254, 30);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Delete);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(data[3]);
                    Thread.Sleep(1000);
                    break;
            }
        }

        public static void CRUD_DETALLE3(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist = desktopSession.FindElementsByClassName("TKactusDBgrid");
            switch (tipo)
            {
                case ("0"):
                    desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 319, 30);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(data[0]);
                    Thread.Sleep(1000);
                    break;
                case ("1"):
                    desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 319, 30);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Delete);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(data[3]);
                    Thread.Sleep(1000);
                    break;
            }
        }

        public static void CRUD_DETALLE4(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist = desktopSession.FindElementsByClassName("TKactusDBgrid");
            var Elementlist1 = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
            switch (tipo)
            {
                case ("0"):
                    desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 37, 30);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(data[4]);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 164, 30);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(data[0]);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 362, 30);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(data[5]);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 656, 30);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(data[6]);
                    Thread.Sleep(1000);
                    //Thread.Sleep(100000000);
                    break;
                case ("1"):
                    desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 164, 30);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Delete);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(data[3]);
                    Thread.Sleep(1000);
                    break;
            }
        }
    }
}

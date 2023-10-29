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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmprvig : FuncionesVitales
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

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 69, 27);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[1]);
            }
            else
            {
                desktopSession.Keyboard.SendKeys(data[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }

      
        public static void ClickButton(WindowsDriver<WindowsElement> desktopSession, int tipo, string file)
        {
            List<string> errorMessages = new List<string>();
            // 1 - Parámatros de Vigencia 2 - Cuentas de Correo 3 - Cuerpo del Correo
            if (tipo == 1)
            {
                desktopSession.FindElementByName("Pámetros de Vigencia").Click();
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Parámetros de Vigencia", true, file);
            }
            else if (tipo == 2)
            {
                desktopSession.FindElementByName("Cuentas de Correo").Click();
            }
            else
            {
                desktopSession.FindElementByName("Cuerpo del Correo").Click();
            }   
        }

        public static void ClickButtonInterno1(WindowsDriver<WindowsElement> desktopSession, int tipo, string file)
        {
            List<string> errorMessages = new List<string>();
            // 1 - Días de Prueba 2 - "Días de Prorroga
            if (tipo == 1)
            {
                desktopSession.FindElementByName("Días de Prueba").Click();
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Cuentas de Correo - Días de Prueba", true, file);
            }
            else if (tipo == 2)
            {
                desktopSession.FindElementByName("Días de Prorroga").Click();
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Cuentas de Correo - Días de Prorroga", true, file);
            }
        }

        public static void ClickButtonInterno2(WindowsDriver<WindowsElement> desktopSession, int tipo, string file)
        {
            List<string> errorMessages = new List<string>();
            // 1 - Días de Prueba 2 - "Días de Prorroga
            if (tipo == 1)
            {
                desktopSession.FindElementByName("Días de Prueba").Click();
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Cuerpo del Correo - Días de Prueba", true, file);
            }
            else if (tipo == 2)
            {
                desktopSession.FindElementByName("Días de Prorroga").Click();
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Cuerpo del Correo - Días de Prorroga", true, file);
            }
        }
    }
}

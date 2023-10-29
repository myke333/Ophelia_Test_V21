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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmrango:FuncionesVitales
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
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            int cont = 0;

            if (tipo == 1)
            {
                ElementList[4].SendKeys(data[0]);
                ElementList[3].SendKeys(data[1]);
                ElementList[2].SendKeys(data[2]);
                newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);
            }
            else
            {
                ElementList[3].Clear();
                ElementList[3].SendKeys(data[3]);
                newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);
            }

        }

        public static void CrudDetalle(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var ElementList = rootSession.FindElementsByClassName("TDBGridInplaceEdit");
            new Actions(rootSession).MoveToElement(ElementList[1], 0, 0).MoveByOffset(50, 40).ClickAndHold().Click().Perform();
            Thread.Sleep(2000);
            new Actions(rootSession).Click().Perform();

            //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 15, 7);
            //desktopSession.Mouse.Click(null);
            Thread.Sleep(500);

            rootSession.Keyboard.SendKeys(data[0]);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);

            rootSession.Keyboard.SendKeys(data[1]);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);

            rootSession.Keyboard.SendKeys(data[2]);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);

            rootSession.Keyboard.SendKeys(data[3]);
        }
    }
}

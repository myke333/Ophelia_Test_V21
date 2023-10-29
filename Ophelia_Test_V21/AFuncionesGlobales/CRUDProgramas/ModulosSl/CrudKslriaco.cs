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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSl
{
    class CrudKslriaco : FuncionesVitales
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

        static public void Click(WindowsDriver<WindowsElement> desktopSession, int bandera, string file)
        {

            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 275, 146);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Click pestaña Experiencia Formal", true, file);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 338, 208);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Se pasan los datos al otro lado", true, file);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 338, 289);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Se devuelven los datos al otro lado", true, file);
            }
        }
    }
}

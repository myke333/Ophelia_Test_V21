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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_9
{
    class CrudKnmmemgr : FuncionesVitales
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

                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }
        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            //for (int i = 0; i < btnTDVI.Count; i++)
            //{
            //    btnTDVI[i].SendKeys(i.ToString());
            //}


            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 54, 56);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(2000);
    
            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 54, 56);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            }

        }

    }
}

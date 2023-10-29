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
    class CrudKsltrada : FuncionesVitales
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

        static public void Proceso(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {

            //var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            //var btnTDVI2 = desktopSession.FindElementsByClassName("#32770");
            //Debugger.Launch();
            //for (int i = 0; i < btnTDVI.Count; i++)
            //{
            //    btnTDVI[i].SendKeys(i.ToString());
            //}
            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            switch (bandera)
            {
                case 0:
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 386, 14);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    break;

                case 1:
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 247, 179);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    break;

                case 2:
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 652, 611);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(4000);
                    break;

                case 3:
                    
                                
                    for (int a = 0; a < 3; a++)
                    {
                        Thread.Sleep(3000);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(3000);
                    }
                    break;

                case 4:
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 908, 611);
                    desktopSession.Mouse.Click(null);
                    break;
            }

          
        }

        
    }
}

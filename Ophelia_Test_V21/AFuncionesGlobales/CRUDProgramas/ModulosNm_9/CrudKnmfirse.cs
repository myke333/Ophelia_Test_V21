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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_9
{
    class CrudKnmfirse : FuncionesVitales
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
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist = desktopSession.FindElementsByClassName("TDBEdit");            
            switch (tipo)
            {
                case ("0"):                    
                    Elementlist[6].SendKeys(data[0]);               
                    Elementlist[6].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[4].SendKeys(data[1]);
                    break;
                case ("1"):
                    Elementlist[4].Clear();
                    Elementlist[4].SendKeys(data[2]);
                    break;
            }
        }

        static public void CrudPrincipal(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudVariables)
        {
        
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            if (bandera==0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 58, 40);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudVariables[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVariables[1]);
            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 523, 40);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(crudVariables[2]);
            }


        }


    }
}

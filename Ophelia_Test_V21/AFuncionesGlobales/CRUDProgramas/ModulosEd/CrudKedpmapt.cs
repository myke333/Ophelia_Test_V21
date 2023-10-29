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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosEd
{
    class CrudKedpmapt : FuncionesVitales
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
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
        {
            List<string> errorMessages = new List<string>();
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementButton = desktopSession.FindElementsByClassName("TDBCmbCodigoBox");
            var ElementEdit = desktopSession.FindElementsByClassName("TDBGrid");
            var ElementEditText = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
            
            if (tipo == 1)
            {

                
                   
                    ElementList[3].SendKeys(crudVariables[0]);
                    ElementList[2].SendKeys(crudVariables[1]);
                    ElementButton[1].SendKeys(crudVariables[2]);
                    ElementButton[0].SendKeys(crudVariables[3]);
                

            }
            if (tipo == 2)
            {

                
                    desktopSession.Mouse.MouseMove(ElementEdit[0].Coordinates, 27, 27);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(crudVariables[5]);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                    desktopSession.Keyboard.SendKeys(crudVariables[6]);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                    desktopSession.Keyboard.SendKeys(crudVariables[7]);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                    desktopSession.Keyboard.SendKeys(crudVariables[8]);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                    desktopSession.Keyboard.SendKeys(crudVariables[9]);
                

            }
            if (tipo == 3)
            {
                desktopSession.Mouse.MouseMove(ElementEdit[0].Coordinates, 27, 27);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);


                desktopSession.Keyboard.SendKeys(crudVariables[10]);
            }

            if(tipo == 4)
            {
                ElementList[3].Clear();

                ElementList[3].SendKeys(crudVariables[4]);

            }
        }
    }
}

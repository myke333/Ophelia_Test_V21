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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKgnwayud:FuncionesVitales
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
           
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit"); 
            var ElementList1 = desktopSession.FindElementsByClassName("TDBMemo"); 
            var ElementList2 = desktopSession.FindElementsByClassName("TDBLookupComboBox");
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(data[2]);
                ElementList1[0].SendKeys(data[1]);
                ElementList[2].SendKeys(data[0]);
                newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);
            }
            else
            {
                ElementList1[0].Clear();
                ElementList1[0].SendKeys(data[3]);
                newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);
            }
        }

        public static void CrudDetalle(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");         
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 134, 30);
                desktopSession.Mouse.Click(null);
                 for (int i = 0; i < 3; i++)
                 {
                    desktopSession.Keyboard.SendKeys(data[i]);
                    Thread.Sleep(500);
                    if (i == 2) { desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter); break; } 
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }
                newFunctions_4.ScreenshotNuevo("Agregar Registro Detalle", true, file);
            }
        }

    }
}

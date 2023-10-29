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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosAc
{
    class CrudKacpaorg:FuncionesVitales
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

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data,string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBCmbCodigoBox");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBEdit");
            if(tipo==1)
            {
                newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);
                ElementList[0].SendKeys(data[0]);
                ElementList2[3].SendKeys(data[1]);
                ElementList2[2].SendKeys(data[2]);
            }
            else
            {
                newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);
                ElementList2[2].Clear();
                ElementList2[2].SendKeys(data[3]);
            }    
        }
        public static void crudDetalle(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data,string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if(tipo==1)
            {
                newFunctions_4.ScreenshotNuevo("Agregar Detalle", true, file);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 27, 27);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[1]);
            }
            else
            {
                newFunctions_4.ScreenshotNuevo("Editar Detalle", true, file);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 27, 27);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[2]);
            }
            

        }



    }
}

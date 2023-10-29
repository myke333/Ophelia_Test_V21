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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmparme:FuncionesVitales
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

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo,string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TGroupButton");

            int contbtn = 1;
            if(tipo==1)
            {
                //ElementList[4].Click();
                foreach (var elem in ElementList)
                {
                    if (contbtn == 13)
                    {
                        elem.Click();
                    }
                    else if (contbtn == 22)
                    {
                        elem.Click();
                    }
                    else if (contbtn == 30)
                    {
                        elem.Click();
                    }
                    else if (contbtn == 38)
                    {
                        elem.Click();
                    }
                    contbtn++;
                    Thread.Sleep(200);
                }
                newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);
            }
            else
            {
                ElementList[13].Click();
                //ElementList[4].Click();
                newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);
            }
                
        }

        public static void primerCLick(WindowsDriver<WindowsElement> desktopSession)
        {
            var ElementList = desktopSession.FindElementsByClassName("TGroupButton");
            ElementList[11].Click();
        }







    }
}

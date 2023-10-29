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
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using OpenQA.Selenium;
using System.Drawing;
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;


namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKgnusuar:FuncionesVitales
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
                    //int cont = 0;
                    //int a = 0;
                    //foreach (var item in Elementlist)
                    //{
                    //    if (cont == 3 && a == 0)
                    //    {
                    //        a = 1;
                    //        cont = 2;
                    //    }
                    //    item.SendKeys(Convert.ToString(cont));
                    //    cont = cont + 1;
                    //}
                    Elementlist[13].SendKeys(data[0]);
                    Elementlist[12].SendKeys(data[1]);
                    Elementlist[8].SendKeys(data[2]);
                    Elementlist[8].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
                case ("1"):
                    Elementlist[12].Clear();
                    Elementlist[12].SendKeys(data[3]);
                    break;
            }
        }
    }
}

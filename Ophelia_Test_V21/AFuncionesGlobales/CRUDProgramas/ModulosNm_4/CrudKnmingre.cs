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
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4
{
    class CrudKnmingre : FuncionesVitales 
    {

        public static WindowsDriver<WindowsElement> RootSession()
        {
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }

        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string Clase)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(Clase);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }


        public static void CRUDKNmIngre(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByName("Concepto de los aportes y Datos a cargo del trabajador o pensionado");
            var ElementList3 = desktopSession.FindElementsByClassName("TPageControl");
            List<string> data = new List<string>() { crudVars[0] };
            List<int> Dis = new List<int>() { 1,2 };
            if (tipo == 1)
            {
                newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, data, Dis);
                ElementList[8].SendKeys(crudVars[1]);
                ElementList[7].SendKeys(crudVars[2]);
                ElementList[10].SendKeys(crudVars[3]);
                ElementList[9].SendKeys(crudVars[4]);
                ElementList[17].SendKeys(crudVars[5]);
                ElementList[19].SendKeys(crudVars[6]);
                ElementList[16].SendKeys(crudVars[7]);
                ElementList2[0].Click();
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList[17].SendKeys(crudVars[8]);
                Thread.Sleep(2000);
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 644, 166);
                //desktopSession.Mouse.Click(null);
                //for (int a=0;a<10;a++)
                //{
                //    desktopSession.Keyboard.SendKeys(crudVars[5]);
                //    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //}
                //Thread.Sleep(2000);
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 228, 172);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 653, 170);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                //for (int a = 0; a < 12; a++)
                //{
                //    desktopSession.Keyboard.SendKeys(crudVars[5]);
                //    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //}
            }
            else
            {
                ElementList[17].Clear();
                ElementList[17].SendKeys(crudVars[9]);
            }
        }
    }
}

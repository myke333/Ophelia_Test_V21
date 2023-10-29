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


namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_8
{
    class CrudKnmcaban : FuncionesVitales
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
  
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, List<String> dataLupa, string file)
        {
            var Elementlist = desktopSession.FindElementsByClassName("TDBEdit");
            var Elementlist1 = desktopSession.FindElementsByClassName("Edit");
            var Elementlist2 = desktopSession.FindElementsByName("Específicos");
            var Elementlist4 = desktopSession.FindElementsByName("Permisos por Usuario");
            switch (tipo)
            {
                case ("0"):
                    Elementlist[4].SendKeys(data[0]);
                    Elementlist[3].SendKeys(data[2]);
                    Elementlist1[2].SendKeys(data[1]);
                    Thread.Sleep(200);
                    PruebaCRUD.LupaDinamica2(desktopSession, dataLupa);      
                    break;

                case ("1"):
                    Elementlist2[0].Click();
                    var Elementlist3 = desktopSession.FindElementsByClassName("TDBEdit");
                    Elementlist3[4].SendKeys(data[3]);
                    break;

                case ("2"):
                    Elementlist4[0].Click();
                    break;

            }
        }
        public static void CRUDdetelles1(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
            switch (tipo)
            {
                case ("0"):
                    Elementlist[1].SendKeys(data[0]);
                    Elementlist[1].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;

                case ("1"):
                    var winTDBGrid = desktopSession.FindElementsByClassName("TDBGrid");
                    desktopSession.Mouse.MouseMove(winTDBGrid[1].Coordinates, 64, 25);
                    desktopSession.Mouse.Click(null);
                    var Elementlist1 = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
                    Elementlist1[1].SendKeys(data[1]);
                    Elementlist1[1].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
            }
        }
        public static void CRUDdetelles2(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist2 = desktopSession.FindElementsByName("Específicos");
            switch (tipo)
            {
                case ("0"):
                    
                    Elementlist2[0].Click();
                    break;

                case ("1"):
                    var winTDBGrid = desktopSession.FindElementsByClassName("TDBGrid");
                    desktopSession.Mouse.MouseMove(winTDBGrid[0].Coordinates, 100,25);
                    desktopSession.Mouse.Click(null);

                    Thread.Sleep(2000);
                    //Debugger.Launch();
                    var Elementlist1 = desktopSession.FindElementsByXPath("//Window[contains(@ClassName, 'TDBGrid']/Edit[@ClassName, 'TDBGridInplaceEdit')]");
                    //Session.FindElementsByXPath("//Window[contains(@ClassName, 'TFrmExcelex')]");
                    Elementlist1[0].SendKeys("0");
                    ////Thread.Sleep(2000);
                    //Elementlist1[0].SendKeys("0");
                    //Elementlist1[0].SendKeys(OpenQA.Selenium.Keys.Tab);
                    //Thread.Sleep(2000);
                    break;
            }
        }
    }
}

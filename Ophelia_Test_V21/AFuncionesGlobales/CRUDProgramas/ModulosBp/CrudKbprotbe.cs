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
using System.Drawing.Imaging;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBp
{
    class CrudKbprotbe : FuncionesVitales
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


        public static List<string> QBEQry(WindowsDriver<WindowsElement> desktopSession, string bandera, string QbeFilter, string file)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            bool IsDisplayedQbe = false;
            if (bandera == "0")
            {
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "IEFrame");
                var allFrame = rootSession.FindElementsByClassName("IEFrame");
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 60, (allFrame[0].Size.Height / 2) + 60).DoubleClick().Click().Perform();
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("QBE", true, file);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TKQBE");
                if (rootSession != null)
                {
                    IsDisplayedQbe = true;
                }
                else
                {
                    errorMessages.Add($"No puede encontrar la ventana de QBE");
                }

                if (IsDisplayedQbe)
                {
                    var elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                    if (QbeFilter != "0")
                    {
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/Pane[contains(@ClassName, 'TStringGrid')]/*"));
                        elements[0].SendKeys(QbeFilter);
                    }
                    elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                    rootSession.Mouse.MouseMove(elements[8].Coordinates, 20, 20);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(500);

                    if (QbeFilter == "0")
                    {
                        rootSession = ReloadSession(rootSession, "TKQBE");
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("//Button[contains(@Name, 'Sí')]"));
                        rootSession.Mouse.MouseMove(elements[0].Coordinates, 20, 20);
                        rootSession.Mouse.Click(null);
                    }
                    Thread.Sleep(500);
                    newFunctions_4.ScreenshotNuevo("Registros QBE", true, file);
                }


            }
            return errorMessages;
        }
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
        {


            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
           


            if (tipo == 1)
            {
                string navigator = AbrirPrograma.SelectNavigator();
                if (navigator == "IE")
                {
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "IEFrame");
                }
                    
                var ElementList = rootSession.FindElementsByClassName("TEdit");
                ElementList[1].SendKeys(crudVariables[0]);
                ElementList[0].SendKeys(crudVariables[1]);

            }
        }
    }
}

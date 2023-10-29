using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4
{
    class CrudKnmremba : PruebaCRUD
    {
        public static void CRUDKnmremba(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TEdit");
            ElementList[1].SendKeys(data[0]);
            ElementList[0].SendKeys(data[1]);
        }

        public static List<string> PreliminarKnmremba(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            WindowsDriver<WindowsElement> RootSession()
            {
                var appCapabilities = new AppiumOptions();
                appCapabilities.AddAdditionalCapability("app", "Root");
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
                return ApplicationSession;
            }
            WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string parametro)
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
            try
            {
                Thread.Sleep(1000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "IEFrame");
                var allFrame = rootSession.FindElementsByClassName("IEFrame");
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) - 70, (allFrame[0].Size.Height / 2) + 40).DoubleClick().Click().Perform();

                Thread.Sleep(3000);
                //Ventana 2
                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TfrmNmRembaFirma");
                //--Click Agregar bmp
                rootSession.FindElementByName("...").Click();
                //rootSession.FindElementByName("...").Click();
                Thread.Sleep(3000);
                rootSession = PruebaCRUD.RootSession();
                Thread.Sleep(1000);
                rootSession.Keyboard.SendKeys("Vive.bmp");
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                rootSession.FindElementByName("Aceptar").Click();
                Thread.Sleep(5000);
                newFunctions_5.GenerarReportes("Preliminar", file, pdf1);                
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;            
        }
    }
}

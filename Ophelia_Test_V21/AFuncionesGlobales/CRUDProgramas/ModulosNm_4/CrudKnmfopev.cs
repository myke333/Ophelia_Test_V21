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
    class CrudKnmfopev : FuncionesVitales
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
        public static void CRUDKNmFopev(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == 1)
            {
                List<string> data = new List<string>() { crudVars[0], crudVars[1], crudVars[2], crudVars[3] };
                PruebaCRUD.LupaDinamica(desktopSession, data);

                ElementList[14].SendKeys(crudVars[4]);
                ElementList[13].SendKeys(crudVars[5]);
                ElementList[12].SendKeys(crudVars[6]);
                ElementList[15].SendKeys(crudVars[7]);
                ElementList[22].SendKeys(crudVars[8]);
                ElementList[23].SendKeys(crudVars[9]);
            }
            else
            {
                ElementList[23].Clear();
                ElementList[23].SendKeys(crudVars[10]);
            }
        }

        public static List<string> PreliKNmFopev(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1, string nomPrograma)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            Thread.Sleep(1000);
            WindowsDriver<WindowsElement> rootSession = null;
            Thread.Sleep(1000);
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "IEFrame");
            var allFrame = rootSession.FindElementsByClassName("IEFrame");
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) - 70, (allFrame[0].Size.Height / 2) + 40).DoubleClick().Click().Perform();
            Thread.Sleep(1000);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);


            Thread.Sleep(3000);
            //Ventana 2
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "TfrmNmFopevFirma");
            //--Click Agregar bmp
            rootSession.FindElementByName("...").Click();
            Thread.Sleep(1000);
            rootSession = PruebaCRUD.RootSession();
            Thread.Sleep(1000);
            rootSession.Keyboard.SendKeys("Vive.bmp");
            Thread.Sleep(500);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(1000);
            rootSession.FindElementByName("Aceptar").Click();
            Thread.Sleep(5000);
            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomPrograma);
            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            return errorMessages;
        }

        static public void reportePreliminar(WindowsDriver<WindowsElement> desktopSession, int bandera, string file)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TGridPanel");
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 130, 17);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            //RootSession = ReloadSession(RootSession, "TFrmNmFopev");

            //var btnTDVI2 = RootSession.FindElementsByClassName("TFrmNmFopev");

            if (bandera == 0)
            {
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                newFunctions_4.ScreenshotNuevo("Primera opcion", true, file);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                RootSession.Keyboard.SendKeys("a");
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                newFunctions_4.ScreenshotNuevo("Insertar ruta", true, file);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                newFunctions_4.ScreenshotNuevo("aceptar primera opcion", true, file);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                newFunctions_4.ScreenshotNuevo("Mensaje de error", true, file);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            }
            else
            {
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }


        }



    }
}

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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_6
{
    class CrudKnmcamji : FuncionesVitales
    {
        public static void CRUDKNmCamji(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file)
        {
            List<string> data = new List<string>() { crudVars[0], crudVars[0] };
            var Aceptar = desktopSession.FindElementsByName("Aceptar");

            Thread.Sleep(500);
            PruebaCRUD.LupaDinamica2(desktopSession, data);
            Thread.Sleep(500);
            WindowsDriver<WindowsElement> rootSession = null;
            Aceptar = desktopSession.FindElementsByName("Aceptar");
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Digitar Datos", true, file);
            Thread.Sleep(500);
            Aceptar[0].Click();
            Thread.Sleep(500);
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Proceso", true, file);
            var allFrame = rootSession.FindElementsByClassName("IEFrame");
            Thread.Sleep(500);
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 130, (allFrame[0].Size.Height / 2) + 90).Click().Perform();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Proceso Finalizado", true, file); 
            Thread.Sleep(500);
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 100, (allFrame[0].Size.Height / 2) + 90).DoubleClick().Perform();
        }
    }
}

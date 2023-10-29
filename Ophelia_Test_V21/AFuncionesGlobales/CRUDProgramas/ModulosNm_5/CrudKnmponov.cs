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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmponov : FuncionesVitales
    {
        public static void CRUDKNmPonov(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TEdit");
            var Aceptar = desktopSession.FindElementsByName("Aceptar");
            WindowsDriver<WindowsElement> rootSession = null;
            List<string> data = new List<string>() { crudVars[0] };
            if (tipo == 1)
            {
                PruebaCRUD.LupaDinamica2(desktopSession, data);
                ElementList[0].SendKeys(crudVars[1]);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Digitar Datos", true, file);
                Thread.Sleep(500);
                
            }
            else
            {
                Aceptar[0].Click();
                Thread.Sleep(500);
                Thread.Sleep(500);
                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Proceso Finalizado", true, file);
                var allFrame = rootSession.FindElementsByClassName("IEFrame");
                Thread.Sleep(500);
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2), (allFrame[0].Size.Height / 2) + 40).DoubleClick().Perform();
            }
        }
    }
}

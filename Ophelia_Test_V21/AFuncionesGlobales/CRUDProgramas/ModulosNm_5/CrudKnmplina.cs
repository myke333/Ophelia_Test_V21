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
    class CrudKnmplina : FuncionesVitales
    {
        public static void CRUDKNmPlina(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file, string motor)
        {
            var Element = desktopSession.FindElementsByClassName("TEdit");
            var Aceptar = desktopSession.FindElementsByName("Aceptar");
            WindowsDriver<WindowsElement> rootSession = null;            
            Thread.Sleep(500);
            Element[2].SendKeys(crudVars[0]);
            Element[4].SendKeys(crudVars[1]);
            Element[3].SendKeys(crudVars[2]);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Digitar Datos", true, file);
            Thread.Sleep(500);
            Aceptar[0].Click();
            Thread.Sleep(500);
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "#32770");
            Thread.Sleep(500);
            var Element4 = rootSession.FindElementsByClassName("Edit");
            var Element5 = rootSession.FindElementsByClassName("Button");
            string Ruta = @"C:\Kactus\Garh\Nm\wrk";
            Element4[0].SendKeys(Ruta);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Sleccionar carpeta", true, file);
            Thread.Sleep(500);
            Element5[0].Click();
            Thread.Sleep(500);
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Proceso Finalizado", true, file);
            var allFrame = rootSession.FindElementsByClassName("IEFrame");
            if(motor=="SQL")
            {
                Thread.Sleep(500);
            }
            else
            {
                Thread.Sleep(130000);
            }
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 90, (allFrame[0].Size.Height / 2) + 80).DoubleClick().Perform();
        }
    }
}
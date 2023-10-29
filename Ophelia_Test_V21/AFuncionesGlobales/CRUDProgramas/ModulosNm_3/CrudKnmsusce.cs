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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmsusce : FuncionesVitales
    {
        public static void CRUDKNmSusce(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TGroupButton");
            ElementList2[2].Click();
            //for (int i = 0; i < ElementList.Count; i++)
            //{
            //    ElementList[i].Click();
            //    Thread.Sleep(500);
            //}
            List<string> data = new List<string>() { crudVars[0], crudVars[1] };
            PruebaCRUD.LupaDinamica(desktopSession, data);
            ElementList[0].SendKeys(crudVars[2]);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Agregar Datos", true, file);
            Thread.Sleep(500);
            var Aceptar = desktopSession.FindElementsByName("Aceptar");
            Aceptar[0].Click();

            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "TFrmNmFinMv");
            ElementList = rootSession.FindElementsByClassName("TEdit");
            Aceptar = rootSession.FindElementsByName("Aceptar");
            ElementList[3].SendKeys(crudVars[3]);
            ElementList[1].SendKeys(crudVars[4]);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Agregar Datos Finalización", true, file);
            Thread.Sleep(500);
            Aceptar[0].Click();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Proceso de Reporte", true, file);
            Thread.Sleep(500);
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
            var allFrame = rootSession.FindElementsByClassName("IEFrame");
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 130, (allFrame[0].Size.Height / 2) + 90).Click().Perform();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Proceso Terminado", true, file);
            Thread.Sleep(500);
            rootSession.Mouse.Click(null);
        }
    }
}

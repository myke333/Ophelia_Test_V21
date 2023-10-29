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
    class CrudKnmcinct : FuncionesVitales
    {
        public static void CRUDKNmCinct(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {

            if (tipo == 1)
            {
                var ElementList = desktopSession.FindElementsByClassName("TEdit");
                var ElementList2 = desktopSession.FindElementsByClassName("TMaskEdit");
                ElementList[1].SendKeys(crudVars[0]);
                ElementList2[0].SendKeys(crudVars[1]);
                ElementList[0].SendKeys(crudVars[2]);
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Digitar Datos", true, file);
            }
            else
            {
                var Aceptar = desktopSession.FindElementsByName("Aceptar");
                Aceptar[0].Click();
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Aceptar", true, file);
            }
        }
    }
}

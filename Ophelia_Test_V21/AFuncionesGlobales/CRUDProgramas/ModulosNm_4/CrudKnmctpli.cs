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
    class CrudKnmctpli : FuncionesVitales
    {

        public static void CRUDKNmCtpli(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TEdit");
            var ElementList2 = desktopSession.FindElementsByName("Ver Errores");
            var ElementList3 = desktopSession.FindElementsByName("Actualiza Calendario Turnos");

            ElementList[2].SendKeys(crudVars[0]);
            ElementList2[0].Click();
            ElementList3[0].Click();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Agregar Fecha", true, file);

        }
        public static void AceptarKNmCtpli(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList4 = desktopSession.FindElementsByName("Aceptar");
            WindowsDriver<WindowsElement> rootSession = null;

            Thread.Sleep(500);
            ElementList4[0].Click();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Botón Aceptar", true, file);
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
            Thread.Sleep(500);
            var allFrame = rootSession.FindElementsByClassName("IEFrame");
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 170, (allFrame[0].Size.Height / 2) + 90).Click().Perform();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Estado Proceso", true, file);
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 90, (allFrame[0].Size.Height / 2) + 85).Click().Click().Perform();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Proceso Finalizado", true, file);

        }
    }
}

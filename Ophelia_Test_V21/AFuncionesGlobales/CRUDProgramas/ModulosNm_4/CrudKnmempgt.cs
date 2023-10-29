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
    class CrudKnmempgt : FuncionesVitales
    {
        public static void CRUDKNmEmpgt(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file)
        {
            var Element = desktopSession.FindElementsByClassName("TEdit");
            WindowsDriver<WindowsElement> rootSession = null;           

            Thread.Sleep(500);
            Element[6].SendKeys(crudVars[0]);
            Thread.Sleep(500);
            List<string> data = new List<string>() { crudVars[1] };
            PruebaCRUD.LupaDinamica2(desktopSession, data);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Digitar Datos", true, file);

            var Element2 = desktopSession.FindElementsByClassName("TScrollBox");
            desktopSession.Mouse.MouseMove(Element2[0].Coordinates, 75, 135);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "TKQBE");
            Thread.Sleep(500);
            var Element3 = rootSession.FindElementsByClassName("TInplaceEdit");
            Element3[0].SendKeys(crudVars[2]);
            var Element4 = rootSession.FindElementsByClassName("TPanel");
            rootSession.Mouse.MouseMove(Element4[5].Coordinates, 40,10);
            rootSession.Mouse.Click(null);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Cargar Empleados", true, file);
            Thread.Sleep(500);
        }
    }    
}

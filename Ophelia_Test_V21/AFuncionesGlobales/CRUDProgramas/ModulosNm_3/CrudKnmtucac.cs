using System;
using System.Collections.Generic;
using System.Diagnostics;
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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmtucac : PruebaCRUD
    {
        public static void CRUDKnmtucac(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();

            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            LupaDinamica2(desktopSession, data);
        }
        public static void BotonAplicar(WindowsDriver<WindowsElement> desktopSession)
        {
            List<string> errorMessages = new List<string>();

            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var Boton = desktopSession.FindElementByName("Aplicar");
            Boton.Click();
            cerrarBorrar(desktopSession);
            desktopSession.Keyboard.SendKeys(Keys.Enter);
        }
    }
}

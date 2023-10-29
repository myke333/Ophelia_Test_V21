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
    class CrudKnmbasli : FuncionesVitales
    {
        public static void CRUDKNmBasli(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            List<string> data = new List<string>() { crudVars[0], crudVars[1], crudVars[2] };
            if (tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, data);
                Thread.Sleep(500);
                ElementList[7].SendKeys(crudVars[5]);
                Thread.Sleep(500);
                ElementList[9].SendKeys(crudVars[4]);
                Thread.Sleep(500);
                ElementList[8].SendKeys(crudVars[3]);
            }
            else
            {
                ElementList[7].Clear();
                ElementList[7].SendKeys(crudVars[6]);
            }
        }
        public static List<string> PreliminarKNmBasli(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1, string nomPrograma = null)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            try
            {
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "TOKBDMenuRep");
            var ElementList = rootSession.FindElementsByName("SI");
            Thread.Sleep(500);
            ElementList[0].Click();
            Thread.Sleep(2000);
            newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomPrograma);
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;
        }
    }
}

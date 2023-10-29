using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmdefsp
    {
        public static WindowsDriver<WindowsElement> RootSession()
        {
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }

        public static void CrudKNmdefsp(WindowsDriver<WindowsElement> desktopSession, int Tipo, string nombre, string nomfis, string edit)
        {
            if (Tipo == 1)
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList[3].SendKeys(nombre);
                ElementList[2].SendKeys(nomfis);
            }
            else if (Tipo == 2)
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList[2].Clear();
                ElementList[2].SendKeys(edit);
            }
        }
        public static void CrudKNmdefspDetalle1(string usuario)
        {
            WindowsDriver<WindowsElement> rootSession = RootSession();
            var ElementList = rootSession.FindElementsByClassName("TDBGrid");
            rootSession.Mouse.MouseMove(ElementList[0].Coordinates, 66, 26);
            rootSession.Mouse.Click(null);

            ElementList[0].SendKeys(usuario);
            //rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
        }
        public static void CrudKNmdefspDetalle2(string codigo)
        {
            WindowsDriver<WindowsElement> rootSession = RootSession();
            var ElementList = rootSession.FindElementsByClassName("TDBGrid");
            rootSession.Mouse.MouseMove(ElementList[0].Coordinates, 66, 26);
            rootSession.Mouse.Click(null);

            ElementList[0].SendKeys(codigo);
            //rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
        }
    }
}

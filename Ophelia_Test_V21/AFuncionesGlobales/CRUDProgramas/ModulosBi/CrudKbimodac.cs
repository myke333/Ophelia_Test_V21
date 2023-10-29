using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbimodac
    {
        public static void CrudKBimodacCabecera(WindowsDriver<WindowsElement> rootSession, int Tipo, string Modalidad, string NomMod, string Peso, string CodDane, string Editar)
        {
            var ElementList = rootSession.FindElementsByClassName("TDBEdit");

            if (Tipo == 1)
            {
                ElementList[5].SendKeys(Modalidad);
                Thread.Sleep(500);
                ElementList[4].SendKeys(Peso);
                Thread.Sleep(500);
                ElementList[3].SendKeys(NomMod);
                Thread.Sleep(500);
                ElementList[2].SendKeys(CodDane);
            }
            else if (Tipo == 2)
            {
                ElementList[3].Clear();
                ElementList[3].SendKeys(Editar);
            }
        }

        public static void CrudKBimodacDetalle(WindowsDriver<WindowsElement> rootSession, int Tipo, string Profesion, string NomProfe, string Registro, string Req, string Espec, string Editar)
        {
            var ElementList = rootSession.FindElementsByClassName("TDBGrid");
            rootSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
            rootSession.Mouse.Click(null);

            if (Tipo == 1)
            {
                ElementList[0].SendKeys(Profesion);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList[0].SendKeys(NomProfe);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList[0].SendKeys(Registro);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList[0].SendKeys(Req);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList[0].SendKeys(Espec);

                rootSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
                rootSession.Mouse.Click(null);
            }
            else if (Tipo == 2)
            {
                ElementList[0].Clear();
                ElementList[0].SendKeys(Editar);
            }
        }
    }
}

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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsoplaem : FuncionesVitales
    {
        public static void CRUDKSoPlaem(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == 1)
            {
                ElementList[7].SendKeys(crudVars[0]);
                ElementList[9].SendKeys(crudVars[1]);
                ElementList[5].SendKeys(crudVars[2]);
                ElementList[4].SendKeys(crudVars[3]);
            }
            else
            {
                ElementList[7].Clear();
                ElementList[7].SendKeys(crudVars[4]);
            }
        }
        public static void CRUDKDetalle1SoPlaem(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 45, 35);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            if (ElementList2.Count == 0)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            }
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVars[0]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVars[1]);
            }
        }
        public static void CRUDKDetalle2SoPlaem(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 45, 35);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[1].FindElementsByClassName("TDBGridInplaceEdit");
            Thread.Sleep(500);
            if (ElementList2.Count == 0)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                ElementList2 = ElementList[1].FindElementsByClassName("TDBGridInplaceEdit");
            }
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVars[0]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVars[1]);
            }
        }
        public static void CRUDKDetalle3SoPlaem(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            WindowsDriver<WindowsElement> rootSession = null;
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 45, 35);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            Thread.Sleep(500);            
            if (tipo == 1)
            {
                if (ElementList2.Count == 0)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                }
                ElementList2[0].SendKeys(crudVars[0]);
                for (int i = 1; i <= 8; i++)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    if (i == 4)
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(500);
                        rootSession = PruebaCRUD.RootSession();
                        rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                        Thread.Sleep(500);
                        var ElementList3 = rootSession.FindElementsByClassName("TEdit");
                        var Aceptar = rootSession.FindElementsByName("Aceptar");
                        Thread.Sleep(500);
                        ElementList3[0].SendKeys(crudVars[4]);
                        ElementList3[1].SendKeys(crudVars[4]);
                        Thread.Sleep(500);
                        Aceptar[0].Click(); 
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                    desktopSession.Keyboard.SendKeys(crudVars[i]);
                   
                }                
            }
            else
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                if (ElementList2.Count == 0)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                }
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVars[9]);
            }
        }
        public static void CRUDKDetalle4SoPlaem(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 45, 35);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            Thread.Sleep(500);
            if (ElementList2.Count == 0)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            }
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVars[0]);
                for (int i = 1; i <= 2; i++)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(crudVars[i]);
                }
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVars[3]);
            }
        }
        public static void CRUDKDetalle5SoPlaem(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file, string motor)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            WindowsDriver<WindowsElement> rootSession = null;
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 45, 35);
            desktopSession.Mouse.Click(null);            
            if (tipo == 1)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                var ElementList2 = ElementList[1].FindElementsByClassName("TDBGridInplaceEdit");
                Thread.Sleep(500);
                if (ElementList2.Count == 0)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    ElementList2 = ElementList[1].FindElementsByClassName("TDBGridInplaceEdit");
                }
                ElementList2[0].SendKeys(crudVars[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                if (motor == "ORA")
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Enter);
                    rootSession = PruebaCRUD.RootSession();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                    Thread.Sleep(500);
                    var ElementList3 = rootSession.FindElementsByClassName("TEdit");
                    var Aceptar = rootSession.FindElementsByName("Aceptar");
                    Thread.Sleep(500);
                    ElementList3[0].SendKeys(crudVars[1]);
                    ElementList3[1].SendKeys(crudVars[1]);
                    Thread.Sleep(500);
                    Aceptar[0].Click();
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }
                else
                {
                    desktopSession.Keyboard.SendKeys(crudVars[1]);
                }
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVars[2]);

            }
            else
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                var ElementList2 = ElementList[1].FindElementsByClassName("TDBGridInplaceEdit");
                Thread.Sleep(500);
                if (ElementList2.Count == 0)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    ElementList2 = ElementList[1].FindElementsByClassName("TDBGridInplaceEdit");
                }
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVars[3]);
            }
        }
        public static void CRUDKDetalle6SoPlaem(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 45, 35);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            Thread.Sleep(500);            
            if (tipo == 1)
            {
                if (ElementList2.Count == 0)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                }
                ElementList2[0].SendKeys(crudVars[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVars[1]);

            }
            else
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                if (ElementList2.Count == 0)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                }
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVars[2]);
            }
        }
        public static void CRUDKDetalle7SoPlaem(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 45, 35);
            desktopSession.Mouse.Click(null);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            var ElementList2 = ElementList[1].FindElementsByClassName("TDBGridInplaceEdit");
            Thread.Sleep(500);
            if (ElementList2.Count == 0)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                ElementList2 = ElementList[1].FindElementsByClassName("TDBGridInplaceEdit");
            }
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVars[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVars[1]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVars[2]);
            }

        }
        public static void CRUDKDetalle8SoPlaem(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 45, 35);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            Thread.Sleep(500);
            if (ElementList2.Count == 0)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            }
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVars[0]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVars[1]);
            }
        }
        public static void EliminarRegistro(WindowsDriver<WindowsElement> desktopSession, WindowsElement ElementList, Dictionary<string, Point> botonesNavegador, string file)
        {
            desktopSession.Mouse.MouseMove(ElementList.Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);           
            newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
            try
            {
                WindowsDriver<WindowsElement> rootSession = null;
                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TMessageForm");
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            catch
            {
                PruebaCRUD.cerrarBorrar(desktopSession);
            }
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
            Thread.Sleep(2000);
        }
        public static void PasarPestaña(WindowsDriver<WindowsElement> desktopSession, int Cantidad, int tipo, int numeral)
        {
            var TPageControl = desktopSession.FindElementsByClassName("TPageControl");
            desktopSession.Mouse.MouseMove(TPageControl[numeral].Coordinates, 30, 13);
            desktopSession.Mouse.Click(null);
            desktopSession.Mouse.Click(null);
            for (int i = 1; i <= Cantidad; i++)
            {
                if (tipo == 1)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                }
                else
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Left);
                }
            }
        }
    }
}

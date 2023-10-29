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
    class CrudKsoparam : FuncionesVitales
    {
        public static void CRUDKSoParam(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file, string motor)
        {
            if (tipo == 1)
            {
                ////PESTAÑA 1 GENERALES
                var ElementList = desktopSession.FindElementsByClassName("TDBLookupComboBox");
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 245, 10);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                ////PASAR A PESTAÑA 2
                var TPageControl = desktopSession.FindElementsByClassName("TPageControl");
                desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 93, 13);
                desktopSession.Mouse.Click(null);
                var ElementList2 = desktopSession.FindElementsByClassName("TDBEdit");
                var ElementList3 = desktopSession.FindElementsByClassName("TDBCheckBox");
                ElementList2[2].SendKeys(crudVars[0]);
                
                if (motor=="SQL")
                {                       
                    ElementList3[6].Click();
                    ElementList3[8].Click();
                    ElementList3[9].Click();
                }
                if (motor == "ORA")
                {
                    ElementList3[2].Click();
                    ElementList3[3].Click();
                    ElementList3[4].Click();
                    ElementList3[6].Click();
                    ElementList3[8].Click();
                }

                ////PASAR A PESTAÑA 3
                PasarPestaña(desktopSession, 1, 1);              

                ElementList3 = desktopSession.FindElementsByClassName("TDBCheckBox");
                if (motor == "SQL")
                {
                    ElementList3[15].Click();
                    ElementList3[5].Click();
                    ElementList3[7].Click();
                    ElementList3[8].Click();
                    ElementList3[9].Click();
                }
                if (motor == "ORA")
                {
                    ElementList3[0].Click();
                    ElementList3[4].Click();
                    ElementList3[6].Click();
                    ElementList3[7].Click();
                    ElementList3[8].Click();
                    ElementList3[9].Click();
                    ElementList3[11].Click();
                    ElementList3[12].Click();
                    ElementList3[14].Click();
                }

                if (motor=="SQL")
                {

                ////PASAR A PESTAÑA 4
                PasarPestaña(desktopSession, 2, 1);

                ElementList2 = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList2[5].SendKeys(crudVars[1]);
                ElementList2[4].SendKeys(crudVars[2]);
                ElementList2[3].SendKeys(crudVars[3]);
                ElementList2[2].SendKeys(crudVars[4]);

                ////PASAR A PESTAÑA 5
                PasarPestaña(desktopSession, 3, 1);

                ElementList3 = desktopSession.FindElementsByClassName("TDBCheckBox");
                ElementList3[5].Click();
                ElementList3[4].Click();
                ElementList2 = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList2[5].SendKeys(crudVars[5]);
                ElementList2[4].SendKeys(crudVars[6]);
                ElementList2[3].SendKeys(crudVars[7]);
                var ElementList4 = desktopSession.FindElementsByClassName("TDBCmbCodigoBox");
                desktopSession.Mouse.MouseMove(ElementList4[0].Coordinates, 50, 10);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                }

                ////PASAR A PESTAÑA 7
                PasarPestaña(desktopSession, 5, 1);

                var ElementList5 = desktopSession.FindElementsByClassName("TDBComboBox");
                if (motor == "SQL")
                {
                    ElementList5[0].SendKeys(crudVars[8]);
                }
                if (motor == "ORA")
                {
                    ElementList5[0].SendKeys(crudVars[1]);
                }

                ////PASAR A PESTAÑA 8
                PasarPestaña(desktopSession, 6, 1);

                List<string> data = new List<string>() { };
                if (motor == "SQL")
                {
                    data = new List<string>() { crudVars[9], crudVars[10], crudVars[11] };
                }
                if (motor == "ORA")
                {
                    data = new List<string>() { crudVars[2], crudVars[3], crudVars[4] };
                }
                Thread.Sleep(500);
                PruebaCRUD.LupaDinamica(desktopSession, data);

                ////PASAR A PESTAÑA 9
                PasarPestaña(desktopSession, 7, 1);

                ElementList3 = desktopSession.FindElementsByClassName("TDBCheckBox");
                Thread.Sleep(500);
                if (motor=="SQL")
                {
                    ElementList3[4].Click();
                    ElementList3[2].Click();
                    ElementList3[1].Click();
                }
                if (motor == "ORA")
                {
                    ElementList3[1].Click();
                    ElementList3[2].Click();
                    ElementList3[3].Click();
                    ElementList3[4].Click();

                    ElementList3[1].Click();
                    ElementList3[1].Click();

                    ElementList2 = desktopSession.FindElementsByClassName("TDBEdit");
                    ElementList2[3].SendKeys(crudVars[5]);
                }

                ////PASAR A PESTAÑA 10
                PasarPestaña(desktopSession, 8, 1);

                ElementList3 = desktopSession.FindElementsByClassName("TDBCheckBox");
                
                if (motor=="ORA")
                {
                    ElementList3[0].Click();
                    ElementList2 = desktopSession.FindElementsByClassName("TDBEdit");
                    ElementList2[2].SendKeys(" ");

                }
                for (int i = 2; i < ElementList3.Count; i++)
                {
                    ElementList3[i].Click();
                }
            }
            else
            {

                var ElementList2 = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList2[3].Clear();
                if (motor == "SQL")
                {
                    ElementList2[3].SendKeys(crudVars[12]);
                }
                if (motor == "ORA")
                {
                    ElementList2[3].SendKeys(crudVars[6]);
                }
            }
        }

        public static void CRUDDetalle1KSoParam(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet1, string file, string motor)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            if (tipo == 1)
            {
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 60, 35);
                desktopSession.Mouse.Click(null);
                if (motor=="ORA")
                {
                    Dictionary<string, Point> botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 1);
                    var NavClass = PruebaCRUD.NavClass(desktopSession);

                    for (int i = 0; i < 2; i++)
                    {                        
                        desktopSession.Mouse.MouseMove(NavClass[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(200);
                        ElementList = desktopSession.FindElementsByClassName("TDBGrid");
                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 60, 35);
                        desktopSession.Mouse.Click(null);
                        ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                        if (ElementList2.Count == 0)
                        {
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                        }
                        ElementList2[0].SendKeys(crudVarsdet1[i]);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                        desktopSession.Mouse.MouseMove(ElementList2[0].Coordinates, 85, 10);
                        desktopSession.Mouse.Click(null);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        desktopSession.Mouse.MouseMove(NavClass[1].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
                        desktopSession.Mouse.Click(null);
                    }
                    desktopSession.Mouse.MouseMove(NavClass[1].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
                    desktopSession.Mouse.Click(null);
                    ElementList = desktopSession.FindElementsByClassName("TDBGrid");
                    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 60, 35);
                    desktopSession.Mouse.Click(null);
                    ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                    if (ElementList2.Count == 0)
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                    }
                }
                if (motor == "ORA")
                {
                    ElementList2[0].SendKeys(crudVarsdet1[2]);
                }
                if (motor == "SQL")
                {
                    ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                    ElementList2[0].SendKeys(crudVarsdet1[0]);
                }
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                desktopSession.Mouse.MouseMove(ElementList2[0].Coordinates, 85, 10);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                if (motor == "ORA")
                {
                    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 60, 75);
                    desktopSession.Mouse.Click(null);
                }
                if (motor == "SQL")
                {
                    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 60, 35);
                    desktopSession.Mouse.Click(null);
                }
                ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                if (ElementList2.Count == 0)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                }
                ElementList2[0].Clear();
                if (motor == "ORA")
                {
                    ElementList2[0].SendKeys(crudVarsdet1[3]);
                }
                if (motor == "SQL")
                {
                    ElementList2[0].SendKeys(crudVarsdet1[1]);
                }
            }
        }

        public static void PasarPestaña(WindowsDriver<WindowsElement> desktopSession, int Cantidad, int tipo)
        {
            var TPageControl = desktopSession.FindElementsByClassName("TPageControl");
            desktopSession.Mouse.MouseMove(TPageControl[0].Coordinates, 130, 13);
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

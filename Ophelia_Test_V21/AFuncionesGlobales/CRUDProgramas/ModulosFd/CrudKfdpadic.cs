﻿using OpenQA.Selenium.Appium;
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
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosFd
{
    class CrudKfdpadic : FuncionesVitales
    {
        static public WindowsDriver<WindowsElement> RootSession()
        {

            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }
        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string parametro)
        {
            try
            {

                var ApplicationWindow = session.FindElementByClassName(parametro);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                //appCapabilities.AddAdditionalCapability("app", "Root");
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }
        public static void LupaDinamicaDiscriminatoria(WindowsDriver<WindowsElement> desktopSession, List<string> data,int pos, List<int> lupasOffIndex = null)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = pos;

            if (lupasOffIndex != null)
            {
                List<Point> backupList = new List<Point>();
                for (int i = 0; i < coordenadas.Count; i++)
                {
                    if (!lupasOffIndex.Contains(i)) backupList.Add(coordenadas[i]);
                }
                coordenadas = backupList;
            }

            for (int coordIndex = 0; coordIndex < coordenadas.Count; coordIndex++)
            {
                Thread.Sleep(2000);
                if (coordIndex == 1) Thread.Sleep(3000);

                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordenadas[coordIndex].X + 10, coordenadas[coordIndex].Y);
                desktopSession.Mouse.Click(null);
                if (contLupa == 1)
                {
                    try
                    {
                        rootSession = RootSession();
                        rootSession = PruebaCRUD.ReloadSession(rootSession, "TKQBE");
                    }
                    catch
                    {

                    }
                    if (rootSession != null)
                    {
                        IsDisplayedQbe = true;
                    }
                    else
                    {
                        errorMessages.Add($"No puede encontrar la ventana de QBE");
                    }
                    if (IsDisplayedQbe)
                    {
                        var elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        if (!string.IsNullOrEmpty(data[indexVal]))
                        {
                            elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/Pane[contains(@ClassName, 'TStringGrid')]/*"));
                            elements[0].SendKeys(data[indexVal]);
                        }
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        new Actions(rootSession).MoveToElement(elements[8], 0, 0).MoveByOffset(20, 10).DoubleClick().Perform();
                        Thread.Sleep(500);
                        IsDisplayedQbe = false;

                    }

                    rootSession = RootSession();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                    //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    rootSession.Keyboard.SendKeys(data[indexVal]);
                    //Aqui se puede enviar un valor

                    var btn = rootSession.FindElementsByClassName("TBitBtn");
                    foreach (var boton in btn)
                    {
                        if (boton.Text == "Aceptar")
                        {
                            rootSession.Mouse.MouseMove(boton.Coordinates, 10, 10);
                            rootSession.Mouse.Click(null);
                        }
                    }
                    indexVal++;
                }
            }

        }
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
        {
            List<string> errorMessages = new List<string>();
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            WindowsDriver<WindowsElement> rootSession = null;
            List<int> lupasOff = new List<int>() { 1,2 };
            rootSession = RootSession();
        
            var ElementButton = desktopSession.FindElementsByClassName("TGroupButton");
            var ElementPage = desktopSession.FindElementsByClassName("TPageControl");
            

            if (tipo == 1)
            {

                {
                   
                    foreach (var elem in ElementButton)
                    {
                        if (elem.Text == crudVariables[0])
                        {
                            elem.Click();
                        }
                        Thread.Sleep(500);
                        if (elem.Text == crudVariables[1])
                        {
                            elem.Click();
                        }
                        Thread.Sleep(500);
                    }

                    desktopSession.Mouse.MouseMove(ElementPage[0].Coordinates, 24, 8);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                    Thread.Sleep(6000);
                    var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                   
                    ElementList[7].SendKeys(crudVariables[2]);
                   
                   

                }

            }

            else
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList[7].Clear();
                ElementList[7].SendKeys(crudVariables[3]);
               
            }
        }

    }
}
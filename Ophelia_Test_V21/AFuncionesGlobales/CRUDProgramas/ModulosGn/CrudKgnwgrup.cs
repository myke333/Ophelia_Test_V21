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
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKgnwgrup:FuncionesVitales
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

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TGroupButton");

            if (tipo == 1)
            {
                ElementList[3].SendKeys(data[0]);
                ElementList[2].SendKeys(data[1]);
                foreach(var elem in ElementList2)
                {
                    if(elem.Text== "Asistente")
                    {
                        elem.Click();
                    }
                }
                newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);
            }
            else
            {
                ElementList[2].Clear();
                ElementList[2].SendKeys(data[2]);

                newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);
            }
        }
        public static void CrudDetalle(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            bool IsDisplayedQbe = false;
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                for (int i = 0; i <2; i++)
                {
                    if (i == 0)
                    {
                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 40, 29);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(500);
                        desktopSession.Keyboard.SendKeys(data[0]);
                    }
                    else if (i == 1)
                    {
                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 140, 29);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(500);

                        desktopSession.Mouse.Click(null);
                        desktopSession.Mouse.DoubleClick(null);
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
                           Debug.WriteLine($"No puede encontrar la ventana de QBE");
                        }
                        if (IsDisplayedQbe)
                        {
                            var elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                            if (!string.IsNullOrEmpty(data[1]))
                            {
                                elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/Pane[contains(@ClassName, 'TStringGrid')]/*"));
                                elements[0].SendKeys(data[1]);
                            }
                            elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                            new Actions(rootSession).MoveToElement(elements[8], 0, 0).MoveByOffset(20, 10).DoubleClick().Perform();
                            Thread.Sleep(500);
                            IsDisplayedQbe = false;

                        }
                        
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TCaptura");
                        Thread.Sleep(1000);
                        var campo = rootSession.FindElementsByClassName("TEdit");
                        ////Debugger.Launch();
                        rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                        rootSession.Mouse.Click(null);
                        campo[1].Clear();
                        rootSession.Keyboard.SendKeys(data[1]);
                        //Aqui se puede enviar un valor

                        var btn = rootSession.FindElementsByClassName("TBitBtn");
                        ////Debugger.Launch();
                        foreach (var boton in btn)
                        {
                            if (boton.Text == "Aceptar")
                            {
                                rootSession.Mouse.MouseMove(boton.Coordinates, 10, 10);
                                rootSession.Mouse.Click(null);
                            }
                        }
                    }
                }
                newFunctions_4.ScreenshotNuevo("Agregar Detalle", true, file);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            }
            else
            {
                ElementList[5].Clear();
                ElementList[5].SendKeys(data[4]);
            }
        }

    }
}

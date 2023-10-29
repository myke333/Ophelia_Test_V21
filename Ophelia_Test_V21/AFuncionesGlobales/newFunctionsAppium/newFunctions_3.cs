using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using Ophelia_Test_V21;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;


namespace Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium
{
    class newFunctions_3 : FuncionesVitales
    {
        static public WindowsDriver<WindowsElement> RootSession()
        {

            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }

        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string className)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(className);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                //appCapabilities.AddAdditionalCapability("app", "Root");
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                ////Debugger.Launch();
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }

        public static List<string> LupaAud(WindowsDriver<WindowsElement> desktopSession, string BanderaLupa, string file)
        {

            List<string> errorMessages = new List<string>();

            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            wait.Until(driver =>
            {
                WindowsDriver<WindowsElement> rootSession = null;

                if (BanderaLupa == "0" || BanderaLupa != "0")
                {
                    string navigator = AbrirPrograma.SelectNavigator();
                    if (navigator == "Edge")
                    {
                        try
                        {
                            rootSession = RootSession();
                            /*string xpath_LeftClickButtonLupa = "//Pane[contains(@ClassName, 'TKNavegador')]/..";
                            ReadOnlyCollection<WindowsElement> xpath_LeftClickButtonLupaA = desktopSession.FindElementsByXPath(xpath_LeftClickButtonLupa);*/
                            var xpath_LeftClickButtonLupaA = desktopSession.FindElementsByClassName("TKNavegador");
                            if (xpath_LeftClickButtonLupaA.Count > 0)
                            {
                                desktopSession.Mouse.MouseMove(xpath_LeftClickButtonLupaA[0].Coordinates, 325, 50);//470,50
                                desktopSession.Mouse.Click(null);
                            }
                            else
                            {
                                //var xpath_LeftClickButtonLupa = "//Pane[contains(@ClassName, 'TPanel')]";///Pane[@ClassName, 'TGridPanel'][@Name, 'GridPanel1']
                                //xpath_LeftClickButtonLupaA = desktopSession.FindElementsByXPath(xpath_LeftClickButtonLupa); 
                                var identificarLup = desktopSession.FindElementsByName("ToolBar1");
                                if (identificarLup.Count > 0)
                                {
                                    xpath_LeftClickButtonLupaA = desktopSession.FindElementsByClassName("TGridPanel");
                                    var cantidad = xpath_LeftClickButtonLupaA.Count;
                                    desktopSession.Mouse.MouseMove(xpath_LeftClickButtonLupaA[0].Coordinates, 470, 50);
                                    desktopSession.Mouse.Click(null);
                                }
                            }
                            //string xpath_LeftClickButton_14_15 = "//Pane[contains(@ClassName,'TPanel')]/Pane[contains(@ClassName,'TPanel')]/Button[@ClassName=\'TBitBtn\']";
                            //var name = rootSession.FindElementByName("Seleccionar X Fecha Auditoria").GetAttribute("Name");
                            var name = desktopSession.FindElementByName("Seleccionar X Fecha Auditoria");
                            //Debugger.Launch();
                            if (xpath_LeftClickButtonLupaA.Count > 0 || name != null /*|| name == "Seleccionar X Fecha Auditoria"*/)
                            {
                                Thread.Sleep(500);
                                newFunctions_4.ScreenshotNuevo("Lupa Auditoria", true, file);
                                //ReadOnlyCollection<WindowsElement> winElem_LeftClickButton_14_15 = desktopSession.FindElementsByXPath(xpath_LeftClickButton_14_15);
                                var winElem_LeftClickButton_14_15 = desktopSession.FindElementsByClassName("TBitBtn");
                                desktopSession.Mouse.MouseMove(winElem_LeftClickButton_14_15[0].Coordinates, 10, 15);
                                desktopSession.Mouse.Click(null);
                                desktopSession.Mouse.Click(null);
                                Thread.Sleep(500);
                                newFunctions_4.ScreenshotNuevo("Consultar Lupa Auditoria", true, file);
                            }
                        }
                        catch
                        {
                            Debug.WriteLine("No se encuentra la opcion Lupa de auditoria");
                        }
                    } else if(navigator == "IE")
                    {
                        try
                        {
                            rootSession = RootSession();
                            string xpath_LeftClickButtonLupa = "//Pane[contains(@ClassName, 'TKNavegador')]/..";
                            ReadOnlyCollection<WindowsElement> xpath_LeftClickButtonLupaA = desktopSession.FindElementsByXPath(xpath_LeftClickButtonLupa);
                            if (xpath_LeftClickButtonLupaA.Count > 0)
                            {
                                desktopSession.Mouse.MouseMove(xpath_LeftClickButtonLupaA[0].Coordinates, 470, 50);
                                desktopSession.Mouse.Click(null);
                            }
                            else
                            {
                                xpath_LeftClickButtonLupa = "//Pane[contains(@ClassName, 'TPanel')]";///Pane[@ClassName, 'TGridPanel'][@Name, 'GridPanel1']
                                xpath_LeftClickButtonLupaA = desktopSession.FindElementsByXPath(xpath_LeftClickButtonLupa);
                                desktopSession.Mouse.MouseMove(xpath_LeftClickButtonLupaA[0].Coordinates, 470, 50);
                                desktopSession.Mouse.Click(null);
                            }
                            string xpath_LeftClickButton_14_15 = "//Pane[contains(@ClassName,'TPanel')]/Pane[contains(@ClassName,'TPanel')]/Button[@ClassName=\'TBitBtn\']";
                            var name = rootSession.FindElementByName("Seleccionar X Fecha Auditoria").GetAttribute("Name");
                            if (xpath_LeftClickButtonLupaA.Count > 0 || name == "Seleccionar X Fecha Auditoria")
                            {
                                Thread.Sleep(500);
                                newFunctions_4.ScreenshotNuevo("Lupa Auditoria", true, file);
                                ReadOnlyCollection<WindowsElement> winElem_LeftClickButton_14_15 = desktopSession.FindElementsByXPath(xpath_LeftClickButton_14_15);
                                desktopSession.Mouse.MouseMove(winElem_LeftClickButton_14_15[0].Coordinates, 10, 15);
                                desktopSession.Mouse.Click(null);
                                desktopSession.Mouse.Click(null);
                                Thread.Sleep(500);
                                newFunctions_4.ScreenshotNuevo("Consultar Lupa Auditoria", true, file);
                            }
                        }
                        catch
                        {
                            Debug.WriteLine("No se encuentra la opcion Lupa de auditoria");
                        }
                    }
                   
                }
                return rootSession != null;
            });
            return errorMessages;
        }
        public static void BotonImprimir(WindowsDriver<WindowsElement> desktopSession, string file)
        {

            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            ////////////////////////////////////////////////
            
            WindowsDriver<WindowsElement> rootSession = null;
            WindowsElement printButton = desktopSession.FindElementByName("Imprimir");

            desktopSession.Mouse.MouseMove(printButton.Coordinates, printButton.Size.Width / 2, printButton.Size.Height / 2);
            desktopSession.Mouse.Click(null);

            Thread.Sleep(10000);

            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TfrxPrintDialog");

            var scrollelement = rootSession.FindElementsByClassName("TComboBox");
            rootSession.Mouse.MouseMove(scrollelement[5].Coordinates, 10, 10);
            rootSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Ventana Impresión PDF", true, file);
            rootSession.Keyboard.SendKeys("m");
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(500);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(1000);
            string pdf = @"C:\Reportes\ReportePDF1_" + "Imprimir" + "_" + Hora();
            rootSession = RootSession();
            rootSession.Keyboard.SendKeys(pdf);
            Thread.Sleep(500);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(5000);
            Process.Start(pdf + ".pdf");
            Thread.Sleep(6000);
            newFunctions_4.ScreenshotNuevo("Impresión PDF", true, file);
            LimpiarProcesos();
        }
        }
    }
    


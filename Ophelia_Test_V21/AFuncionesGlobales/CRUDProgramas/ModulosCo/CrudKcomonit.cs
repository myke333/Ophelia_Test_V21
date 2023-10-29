﻿using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Drawing;
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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKcomonit : FuncionesVitales
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

        public static void KCoMonitCRUD(WindowsDriver<WindowsElement> desktopSession, Dictionary<int, string> variables, string className, int key, int flag, string editValue)
        {

            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName(className);
            //System.Diagnostics.//Debugger.Launch();

            if (flag == 0)
            {
                foreach (KeyValuePair<int, string> elem in variables)
                {
                    editFields[elem.Key].ClearCache();
                    desktopSession.Mouse.MouseMove(editFields[elem.Key].Coordinates, editFields[elem.Key].Size.Width / 2, editFields[elem.Key].Size.Height / 2);
                    desktopSession.Mouse.DoubleClick(null);
                    //Thread.Sleep(500);
                    //desktopSession.Mouse.DoubleClick(null);
                    Thread.Sleep(1000);
                    desktopSession.Keyboard.SendKeys(elem.Value);
                    Thread.Sleep(1000);
                }
            }
            else
            {
                foreach (KeyValuePair<int, string> elem in variables)
                {
                    if (elem.Key == key)
                    {
                        editFields[elem.Key].ClearCache();
                        editFields[elem.Key].Clear();
                        desktopSession.Mouse.MouseMove(editFields[elem.Key].Coordinates, editFields[elem.Key].Size.Width / 2, editFields[elem.Key].Size.Height / 2);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);
                        desktopSession.Keyboard.SendKeys(editValue);
                        Thread.Sleep(1000);
                    }


                }
            }



        }

        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {
                                
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                
        }

        public static void ClickButtonAceptar(WindowsDriver<WindowsElement> desktopSession)
        {
            var ElementList = desktopSession.FindElementsByName("Aceptar");
            ElementList[0].Click();
        }

        static public void Click(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {

            //var btnTDVI2 = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 153, 96);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 472, 362);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(1500);

            }
        }

    }
}
using OpenQA.Selenium.Appium;
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
    class CrudKbifamil 
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



        public static void KBifamilCRUD(WindowsDriver<WindowsElement> desktopSession, Dictionary<int, string> variables, string className, int key, int flag, string editValue) {

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
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Keyboard.SendKeys(elem.Value);
                    Thread.Sleep(1000);
                }
            }
            else {
                foreach (KeyValuePair<int, string> elem in variables)
                {
                    if (elem.Key == key) {
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

        static public void CrudPrincipal(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal, string file)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
  
            if (bandera==0)
            {
                Lupa(desktopSession, crudPrincipal, file);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 690, 34);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudPrincipal[7]);
                for (int a = 0; a < 6; a++)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[11]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudPrincipal[10]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudPrincipal[9]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 690, 34);
                desktopSession.Mouse.Click(null);
               for (int a=0;a<4;a++)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }
                desktopSession.Keyboard.SendKeys(crudPrincipal[12]);
            }

          

        }

        static public void Lupa(WindowsDriver<WindowsElement> desktopSession, List<string> crudPrincipal, string file)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            int variable = 0;
          
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 767, 34);
                desktopSession.Mouse.Click(null);
            newFunctions_4.ScreenshotNuevo("Pestaña QBE", true, file);
            enter(desktopSession);
                enter(desktopSession);


            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TCaptura");

            newFunctions_4.ScreenshotNuevo("Pestaña", true, file);

            var btnTDVI2 = RootSession.FindElementsByClassName("TCaptura");



            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 220, 69);
            RootSession.Mouse.Click(null);
            Thread.Sleep(2000);
            RootSession.Keyboard.SendKeys(crudPrincipal[variable]);
            Thread.Sleep(2000);
            newFunctions_4.ScreenshotNuevo("Valor elegido", true, file);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

            enter(desktopSession);

        }

        static public void enter(WindowsDriver<WindowsElement> desktopSession)
        {


            //Dar enter
            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(1000);
          

        }

        static public void reportePreliminar(WindowsDriver<WindowsElement> desktopSession, int bandera, string file)
        {
            
            var btnTDVI = desktopSession.FindElementsByClassName("TGridPanel");
            
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 130, 17);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TFrmOpcRepo");
            Thread.Sleep(2000);
            var btnTDVI2 = RootSession.FindElementsByClassName("TFrmOpcRepo");

            if (bandera == 0)
            {
                newFunctions_4.ScreenshotNuevo("Primera opcion", true, file);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                
                newFunctions_4.ScreenshotNuevo("segunda opcion", true, file);
                Thread.Sleep(2000);
                RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 41, 94);
                RootSession.Mouse.Click(null);
                Thread.Sleep(2000);
                RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 119, 147);
                RootSession.Mouse.Click(null);
                Thread.Sleep(2000);

                RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 155, 46);
                
                RootSession.Mouse.Click(null);
                
                Thread.Sleep(2000);
               
                RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 187, 49);
                for (int a = 0; a < 23; a++)
                {
                    RootSession.Mouse.Click(null);

                }
                newFunctions_4.ScreenshotNuevo("Primera fecha seleccionada", true, file);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 230, 140);
                RootSession.Mouse.Click(null);

                Thread.Sleep(2000);
                newFunctions_4.ScreenshotNuevo("Segunda fecha seleccionada", true, file);
                RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 155, 46);
                RootSession.Mouse.Click(null);
                Thread.Sleep(2000);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

                newFunctions_4.ScreenshotNuevo("Fechas seleccionadas", true, file);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }


        }



    }
}

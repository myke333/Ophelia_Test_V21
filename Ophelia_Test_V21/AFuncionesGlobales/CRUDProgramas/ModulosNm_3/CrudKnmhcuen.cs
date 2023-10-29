using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;




namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmhcuen
    {
        public static WindowsDriver<WindowsElement> RootSession()
        {
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }

        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string Clase)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(Clase);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; } 
        }
        public static void CrudKNmhcuen(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            if (Tipo == 1)
            {
                List<string> lupa = new List<string> { data[0], data[1], data[2] };
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("(EPS) Entidad Promotora de Salud").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                PruebaCRUD.LupaDinamica(desktopSession, lupa);
                Thread.Sleep(500);
                ElementList[8].SendKeys(data[3]);
                Thread.Sleep(5000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 30, 317);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(data[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[2]);


            }
            else if (Tipo == 2)
            {
                Thread.Sleep(500);
                ElementList[8].Clear();
                Thread.Sleep(500);
                ElementList[8].SendKeys(data[4]);
                Thread.Sleep(500);
            }
        }
        public static List<string> PreliminarKNmhcuen(WindowsDriver<WindowsElement> desktopSession, string BanderaPreli, string file, string pdf1)
        {
            WindowsDriver<WindowsElement> rootSession;
            List<string> ErrorValidacion = new List<string> { };
            List<string> errors = new List<string> { };
            
            for(int i = 0; i < 7; i++)
            {
                FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                Thread.Sleep(1000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmTipoRep");
                var ElementList = rootSession.FindElementsByClassName("TGroupButton");
                rootSession.Mouse.MouseMove(ElementList[i].Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(500);
                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Aceptar").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(500);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                Thread.Sleep(500);               
            }
            for (int i = 0; i < 7; i++)
            {
                FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                Thread.Sleep(1000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmTipoRep");
                var ElementList = rootSession.FindElementsByClassName("TGroupButton");
                if(i == 0)
                {
                    rootSession.Mouse.MouseMove(ElementList[7].Coordinates);
                    rootSession.Mouse.Click(null);
                }
                rootSession.Mouse.MouseMove(ElementList[i].Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(500);
                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Aceptar").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(500);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                Thread.Sleep(500);
            }
            return ErrorValidacion;







        }

        static public void reportePreliminar(WindowsDriver<WindowsElement> desktopSession, int bandera, int bandera2, string file)
        {
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();

            var btnTDVI = desktopSession.FindElementsByClassName("TGridPanel");
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 130, 17);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TFrmTipoRep");

            var btnTDVI2 = RootSession.FindElementsByClassName("TFrmTipoRep");

         
           
                switch (bandera2)
                {
                    case 0:
                    if (bandera==0)
                    {
                        RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 40, 64);
                        RootSession.Mouse.Click(null);
                        newFunctions_4.ScreenshotNuevo("Opcion resumida", true, file);
                    }
                    else
                    {
                        RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 190, 64);
                        RootSession.Mouse.Click(null);
                        newFunctions_4.ScreenshotNuevo("Opcion Detallada", true, file);
                    }
                    RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 40, 113);
                    RootSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("Sub opcion N1", true, file);
                     RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                     Thread.Sleep(2000);
                      break;
                    case 1:
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        break;
                    case 2:
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        break;
                }
         
          



        }


    }
}

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
	class CrudKBiDatad:FuncionesVitales
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

        static public void CrudData(WindowsDriver<WindowsElement> desktopSession, string tipo, List<string> crudPrincipal, string file)
		{
            var Elementlist = desktopSession.FindElementsByClassName("TDBEdit");
            var selec = desktopSession.FindElementsByClassName("TDBEdit");

			switch (tipo)
			{
                case ("0"):
                   
                    Elementlist[5].SendKeys(crudPrincipal[0]);
                    Elementlist[6].SendKeys(crudPrincipal[7]);
                    Elementlist[7].SendKeys(crudPrincipal[8]);
                    Elementlist[8].SendKeys(crudPrincipal[6]);
                    Elementlist[9].SendKeys(crudPrincipal[5]);
                    Elementlist[10].SendKeys(crudPrincipal[4]);
                    Elementlist[11].SendKeys(crudPrincipal[3]);
                    Elementlist[12].SendKeys("Negro");
                    Elementlist[13].SendKeys(crudPrincipal[2]);
                    Elementlist[14].SendKeys(crudPrincipal[1]);
                    //Factor
                    desktopSession.Mouse.MouseMove(selec[14].Coordinates, -15, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> FacSelect = null;
                    FacSelect = PruebaCRUD.RootSession();
                    FacSelect.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> FacEnter = null;
                    FacEnter = PruebaCRUD.RootSession();
                    FacEnter.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);
                    //Grupo
                    desktopSession.Mouse.MouseMove(selec[14].Coordinates, -75, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> GpoSelect = null;
                    GpoSelect = PruebaCRUD.RootSession();
                    GpoSelect.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> GpoEnter = null;
                    GpoEnter = PruebaCRUD.RootSession();
                    GpoEnter.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    //Seleccionar Casilla
                    desktopSession.Mouse.MouseMove(selec[14].Coordinates, -75, 50);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.MouseMove(Elementlist[5].Coordinates, 230, 45);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[5].Coordinates, 10, 100);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> DatProf1 = null;
                    DatProf1 = PruebaCRUD.RootSession();
                    DatProf1.Keyboard.SendKeys("ProfesionPrueba");
                    Thread.Sleep(1000);
                    DatProf1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf1.Keyboard.SendKeys("2");
                    Thread.Sleep(1000);
                    DatProf1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf1.Keyboard.SendKeys("MCN 395");
                    Thread.Sleep(1000);
                    DatProf1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf1.Keyboard.SendKeys("Antioquia");
                    Thread.Sleep(1000);
                    DatProf1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf1.Keyboard.SendKeys("PruebaObs");
                    Thread.Sleep(1000);
                    DatProf1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf1.Keyboard.SendKeys("PruebaHobbies");
                    Thread.Sleep(1000);
                    DatProf1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf1.Keyboard.SendKeys("PruebaProps");
                    Thread.Sleep(1000);
                    DatProf1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf1.Keyboard.SendKeys("Gatos");
                    Thread.Sleep(1000);
                    DatProf1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf1.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.MouseMove(Elementlist[5].Coordinates, 240, 150);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.MouseMove(Elementlist[5].Coordinates, 680, 160);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    

                    break;

                case ("1"):

                    WindowsDriver<WindowsElement> DatosEdit = null;
                    DatosEdit = PruebaCRUD.RootSession();

                    desktopSession.Mouse.MouseMove(Elementlist[5].Coordinates, 10, 45);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[5].Coordinates, 150, 145);
                    desktopSession.Mouse.DoubleClick(null);
                    Thread.Sleep(1000);

                    
                    DatosEdit.Keyboard.SendKeys(crudPrincipal[9]);
                    Thread.Sleep(1000);

                    DatosEdit.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatosEdit.Keyboard.SendKeys(crudPrincipal[10]);
                    Thread.Sleep(1000);

                    DatosEdit.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatosEdit.Keyboard.SendKeys("Mestizo");
                    Thread.Sleep(1000);

                    DatosEdit.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatosEdit.Keyboard.SendKeys(crudPrincipal[11]);
                    Thread.Sleep(1000);

                    DatosEdit.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatosEdit.Keyboard.SendKeys(crudPrincipal[12]);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[5].Coordinates, 150, 255);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    DatosEdit.Keyboard.SendKeys(crudPrincipal[15]);
                    Thread.Sleep(1000);

                    DatosEdit.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatosEdit.Keyboard.SendKeys(crudPrincipal[16]);

                    ///******************
                    
                    break;

                case ("2"):

                    //Factor
                    desktopSession.Mouse.MouseMove(selec[14].Coordinates, -15, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> FacSelectEdit = null;
                    FacSelectEdit = PruebaCRUD.RootSession();
                    FacSelectEdit.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> FacEnterEdit = null;
                    FacEnterEdit = PruebaCRUD.RootSession();
                    FacEnterEdit.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);
                    //Grupo
                    desktopSession.Mouse.MouseMove(selec[14].Coordinates, -75, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> GpoSelectEdit = null;
                    GpoSelectEdit = PruebaCRUD.RootSession();
                    GpoSelectEdit.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> GpoEnterEdit = null;
                    GpoEnterEdit = PruebaCRUD.RootSession();
                    GpoEnterEdit.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(3000);

                    //Seleccionar Casilla
                    desktopSession.Mouse.MouseMove(selec[14].Coordinates, -75, 50);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.MouseMove(Elementlist[5].Coordinates, 230, 45);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[5].Coordinates, 10, 100);
                    desktopSession.Mouse.DoubleClick(null);
                    Thread.Sleep(1000);

                    //Pestaña Otros Datos
                    WindowsDriver<WindowsElement> DatProf2 = null;
                    DatProf2 = PruebaCRUD.RootSession();

                    DatProf2.Keyboard.SendKeys("Profesion Prueba Edit");
                    Thread.Sleep(1000);
                    DatProf2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf2.Keyboard.SendKeys("3");
                    Thread.Sleep(1000);
                    DatProf2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf2.Keyboard.SendKeys("MMN 39E");
                    Thread.Sleep(1000);
                    DatProf2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf2.Keyboard.SendKeys("ARMENIA");
                    Thread.Sleep(1000);
                    DatProf2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf2.Keyboard.SendKeys("PruebaObsEdit");
                    Thread.Sleep(1000);
                    DatProf2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf2.Keyboard.SendKeys("PruebaHobbiesEdit");
                    Thread.Sleep(1000);
                    DatProf2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf2.Keyboard.SendKeys("PruebaPropsEdit");
                    Thread.Sleep(1000);
                    DatProf2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf2.Keyboard.SendKeys("Perro");
                    Thread.Sleep(1000);
                    DatProf2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    DatProf2.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.MouseMove(Elementlist[5].Coordinates, 240, 150);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.MouseMove(Elementlist[5].Coordinates, 680, 160);
                    desktopSession.Mouse.Click(null);
                    break;




            }
            /*
                        var cont = 0;

                        foreach(var item in Elementlist)
                        {
                            if(cont == 3)
                            {
                                item.SendKeys("2");
                            }
                            else
                            {
                                item.SendKeys(Convert.ToString(cont));
                            }

                            cont = cont + 1;
                            Thread.Sleep(1000);
                        }
            */
        }


        static public void reportePreliminar(WindowsDriver<WindowsElement> desktopSession, int bandera, string file)
        {
            Thread.Sleep(4000);
            var btnTDVI = desktopSession.FindElementsByClassName("TGridPanel");
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 130, 17);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TFrmSeleccion");
            Thread.Sleep(4000);
            var btnTDVI2 = RootSession.FindElementsByClassName("TPanel");
   
            switch (bandera)
            {
                case 0:
                    //primera opcion 
                    newFunctions_4.ScreenshotNuevo("Primera opcion preliminar ", true, file);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    break;
                case 1:
                    //segunda opcion // Carpeta
                    RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 52, 77);
                    RootSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("Segunda opcion preliminar - carpeta ", true, file);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    break;
                case 2:
                    //Segunda opcion // caja
                    RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 52, 77);
                    RootSession.Mouse.Click(null);
                    RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 194, 97);
                    RootSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("Segunda opcion preliminar - caja ", true, file);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                  
                    break;
                case 3:
                    //Tercera opcion
                    RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 52, 122);
                    RootSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("tercera opcion preliminar ", true, file);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    break;
                case 4:
                    //Cuarta opcion
                    Thread.Sleep(4000);
                    RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 52, 163);
                    RootSession.Mouse.Click(null);
                    Thread.Sleep(4000);
                    newFunctions_4.ScreenshotNuevo("cuarta opcion preliminar ", true, file);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    break;
                case 5:
                    //Tercera opcion
                    RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 52, 205);
                    RootSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("quinta opcion preliminar ", true, file);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    break;

            }

        }

       


    }
}

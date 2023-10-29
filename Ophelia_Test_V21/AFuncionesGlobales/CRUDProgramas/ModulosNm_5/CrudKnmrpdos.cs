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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmrpdos : FuncionesVitales
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
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, List<string> data, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TCheckBox");
            var Boton = desktopSession.FindElementsByClassName("TScrollBox");

            ElementList[2].SendKeys(data[0]);
            ElementList[1].SendKeys(data[1]);

            foreach (var elem in ElementList2)
            {
                if (elem.Text == "Nómina")
                {
                    elem.Click();
                }
            }
            newFunctions_4.ScreenshotNuevo("Agregar Datos", true, file);
            Thread.Sleep(1000);
            desktopSession.Mouse.MouseMove(Boton[0].Coordinates, 130, 88);
            desktopSession.Mouse.Click(null);
        }
        public static void Preliminar(WindowsDriver<WindowsElement> desktopSession, string file,string pdf1, string Bandera)
        {
            for (int i = 0; i < 6; i++)
            {
                FuncionesGlobales.ReportePreliminar(desktopSession, Bandera, file, pdf1);
                WindowsDriver<WindowsElement> rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmRep");
                var Elemtlist = rootSession.FindElementsByClassName("TPageControl");
                Debug.WriteLine($"Cantidad de Pestañas: {Elemtlist.Count}");
                rootSession.Mouse.MouseMove(Elemtlist[0].Coordinates, 5, 5);
                rootSession.Mouse.Click(null);

                if (i <= 2)
                {       
                    var Elemtlist2 = rootSession.FindElementsByClassName("TCheckBox");
                    if (i == 0)
                    {
                        foreach (var elem2 in Elemtlist2)
                        {
                            if (elem2.Text == "Identificación")
                            {
                                elem2.Click();
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            }
                        }
                        NotePad(rootSession, file, i);
                    }//Final Primer reporte
                    else if (i == 1)
                    {
                        foreach (var elem2 in Elemtlist2)
                        {
                            if (elem2.Text == "Desprendibles Standar")
                            {
                                elem2.Click();
                            }
                        }
                        NotePad(rootSession, file, i);
                    }//Final Segundo reporte
                    else if (i == 2)
                    {
                        foreach (var elem2 in Elemtlist2)
                        {
                            if (elem2.Text == "Alfabetico ( Requiere Selección de Datos)")
                            {
                                elem2.Click();
                            }
                        }

                        PruebaCRUD.QBEQry(rootSession, "0", "", file);
                        NotePad(rootSession, file, i);
                    }//Final Tercer reporte

                }//Fin del si de la primera Pestaña
                else if (i == 3)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);                    
                    rootSession = ReloadSession(rootSession, "TFrmRep");
                    var Elemtlist2 = rootSession.FindElementsByClassName("TCheckBox");

                    foreach (var elem2 in Elemtlist2)
                    {
                        if (elem2.Text == "Centro de Costo/Empleados/Formas De Pago")
                        {
                            elem2.Click();
                            rootSession = RootSession();
                            rootSession = ReloadSession(rootSession, "TFrmRep");
                            var cheks = rootSession.FindElementsByClassName("TCheckBox");
                            foreach (var chk in cheks)
                            {
                                if (chk.Text != "Centro de Costo/Empleados/Formas De Pago")
                                {
                                    chk.Click();
                                }
                            }
                        }
                    }
                    NotePad(rootSession, file, i);

                }//Fin del si de la segunda Pestaña
                else if (i == 4)
                {
                    Elemtlist = rootSession.FindElementsByClassName("TPageControl");
                    Debug.WriteLine($"Cantidad de Pestañas: {Elemtlist.Count}");
                    rootSession.Mouse.MouseMove(Elemtlist[0].Coordinates, 5, 5);
                    rootSession.Mouse.Click(null);
                    for (int j = 1; j <= 2; j++) {
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                        Thread.Sleep(1000);
                    }
                    rootSession = ReloadSession(rootSession, "TFrmRep");
                    var Elemtlist2 = rootSession.FindElementsByClassName("TCheckBox");

                    foreach (var elem2 in Elemtlist2)
                    {
                        if (elem2.Text == "Resumen General Nómina")
                        {
                            elem2.Click();
                        }
                    }
                    NotePad(rootSession, file, i);

                }//Fin del si de la Tercera Pestaña
                else if (i == 5)
                {
                    Elemtlist = rootSession.FindElementsByClassName("TPageControl");
                    Debug.WriteLine($"Cantidad de Pestañas: {Elemtlist.Count}");
                    rootSession.Mouse.MouseMove(Elemtlist[0].Coordinates, 5, 5);
                    rootSession.Mouse.Click(null);
                    for (int j = 1; j <= 3; j++)
                    {
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                        Thread.Sleep(1000);
                    }

                    rootSession = ReloadSession(rootSession, "TFrmRep");

                    var Elemtlist2 = rootSession.FindElementsByClassName("TCheckBox");

                    foreach (var elem2 in Elemtlist2)
                    {
                        if (elem2.Text == "Entidades/Cuenta/Empleado( Requiere Seleccion de Datos)")
                        {
                            elem2.Click();
                        }
                    }
                    PruebaCRUD.QBEQry(rootSession, "0", "", file);
                    NotePad(rootSession, file, i);

                }//Fin del si de la Cuarta Pestaña                
            }
        }

        public static void NotePad(WindowsDriver<WindowsElement> rootSession, string file, int i)
        {
            var Boton = rootSession.FindElementsByClassName("TBitBtn");
            foreach (var btn in Boton)
            {
                if (btn.Text == "Aceptar")
                {
                    btn.Click();
                    if (i == 0 || i == 2 || i == 3 || i == 4 || i == 5)
                    {
                        Thread.Sleep(15000);
                    }
                    else if (i == 1)
                    {
                        Thread.Sleep(35000);
                    }

                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TMessageForm");
                    var Report = rootSession.FindElementsByClassName("TButton");
                    foreach (var rep in Report)
                    {
                        if (rep.Text == "Yes")
                        {
                            rep.Click();
                        }
                    }
                    Thread.Sleep(4000);
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "Notepad");
                    newFunctions_4.ScreenshotNuevo("Reporte Txt", true, file);
                    var cerrar = rootSession.FindElementByName("Cerrar");
                    cerrar.Click();
                }
            }

        }
    }
}

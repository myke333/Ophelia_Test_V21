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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_10
{
    class CrudKnmcinca : FuncionesVitales
    {
        static public WindowsDriver<WindowsElement> RootSession()
        {
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }

        static public WindowsDriver<WindowsElement> ReloadSession1(WindowsDriver<WindowsElement> session, string parametro)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(parametro);
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

        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            //var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();
            //for (int i = 0; i < btnTDVI.Count; i++)
            //{
            //    btnTDVI[i].SendKeys(i.ToString());
            //}


            if (bandera == 0)
            {
                btnTDVI[7].SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1500);
                btnTDVI[5].SendKeys(crudPrincipal[2]);
                Thread.Sleep(1000);

                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 101, 27);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                //Thread.Sleep(1500);

            }
            else
            {
                btnTDVI[6].Clear();
                btnTDVI[6].SendKeys(crudPrincipal[3]);
                Thread.Sleep(1500);

                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 101, 27);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                //Thread.Sleep(1500);

            }
        }

        static public void Click(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            var btnTDVI2 = desktopSession.FindElementsByClassName("TDBGrid");

            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 1292, 245);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);
            newFunctions_4.ScreenshotNuevo("Click QBE Det 1", true, file);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(1500);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TKQBE");

            var btnTDVI = RootSession.FindElementsByClassName("TTabSheet");

            RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 16);
            RootSession.Mouse.Click(null);
            Thread.Sleep(2000);

            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);
            newFunctions_4.ScreenshotNuevo("Registros QBE", true, file);

            WindowsDriver<WindowsElement> RootSession3 = null;
            RootSession3 = PruebaCRUD.RootSession();
            RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);
        }

        static public void AgregarCrud1(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet1)
        {

            //var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI = desktopSession.FindElementsByClassName("TDBGrid");
            //var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 35, 30);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet1[0]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet1[1]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1235, 30);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1220, 42);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1363, 30);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet1[2]);
                Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1363, 30);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet1[3]);
                Thread.Sleep(1000);

            }
        }

        static public void Clickk(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {

            //var btnTDVI2 = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI2 = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 482, 258);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 91, -10);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(1000);
                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 586, 170);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 22, 570);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession.Keyboard.SendKeys("ReportePDF1_Preliminar_08022022_161632");
                //Thread.Sleep(1000);
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            }
        }

        public static List<string> generarExcel(WindowsDriver<WindowsElement> Session, string file, string codProgram, string ruta)
        {
            List<string> errorMessages = new List<string>();
            try
            {
                Session = RootSession();
                Session = ReloadSession1(Session, "TFrmExcelex");
                var element = Session.FindElementsByXPath("//Window[contains(@ClassName, 'TFrmExcelex')]");
                if (element[0] != null)
                {
                    bool IsDisplayedExcel = false;
                    string Fecha = Hora();
                    Thread.Sleep(500);
                    if (element[0] != null)
                    {
                        IsDisplayedExcel = true;
                    }
                    else
                    {
                        errorMessages.Add($"No se puede visualizar la ventana Exportar a excel");
                    }

                    if (IsDisplayedExcel)
                    {
                        Session = RootSession();
                        Session = ReloadSession1(Session, "TFrmExcelex");
                        var ReportElements = Session.FindElementsByXPath("//Window[contains(@ClassName, 'TFrmExcelex')]");
                        Session.Mouse.MouseMove(ReportElements[0].Coordinates, 222, 207);
                        Session.Mouse.Click(null);
                        newFunctions_4.ScreenshotNuevo("Exportar datos Excel", true, file);
                        var buttonAccept = Session.FindElementByName("Aceptar");
                        buttonAccept.Click();

                        Thread.Sleep(5000);
                        try
                        {
                            Session = RootSession();
                            Session = ReloadSession1(Session, "XLMAIN");
                            if (Session != null)
                            {
                                int count = 0;
                                while (count < 240)
                                {
                                    try
                                    {
                                        //Session = RootSession();
                                        //Session = ReloadSession1(Session, "XLMAIN");
                                        Thread.Sleep(500);
                                        var saveExcel1 = Session.FindElementsByName("Maximizar");
                                        if (saveExcel1.Count > 0)
                                        {
                                            break;
                                        }
                                    }
                                    catch
                                    {
                                        Thread.Sleep(1000);
                                        count++;
                                    }
                                }
                                var saveExcel = Session.FindElementsByName("Maximizar");
                                var cantidad = saveExcel.Count;
                                if (cantidad == 2)
                                {
                                    saveExcel[1].Click();
                                }
                                else
                                {
                                    saveExcel[0].Click();
                                }
                                //Cambio Jose
                                Session.FindElementByName("Pestaña Archivo").Click();
                                //Fin Cambio Jose

                                //Session.FindElementByName("Guardar").Click();
                                Session.FindElementByName("Guardar").Click();
                                Session.FindElementByName("Examinar").Click();
                                Thread.Sleep(1000);
                                Session.Keyboard.SendKeys(ruta);
                                Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                Thread.Sleep(1000);
                                newFunctions_4.ScreenshotNuevo("Reporte Excel Guardado", true, file);
                                LimpiarProcesos();

                            }
                            else
                            {
                                //Debugger.Launch();
                                Thread.Sleep(1000);
                                Session = RootSession();
                                Session = ReloadSession1(Session, "Shell_TrayWnd");
                                Thread.Sleep(1000);
                                var venExcel = Session.FindElementByName("Excel - 1 ventana de ejecución");
                                if (venExcel != null)
                                {
                                    venExcel.Click();
                                    Thread.Sleep(4000);
                                    int count = 0;
                                    while (count < 240)
                                    {
                                        try
                                        {
                                            Session = RootSession();
                                            Session = ReloadSession1(Session, "XLMAIN");
                                            Thread.Sleep(500);
                                            var saveExcel1 = Session.FindElementsByName("Maximizar");
                                            if (saveExcel1.Count > 0)
                                            {
                                                break;
                                            }
                                        }
                                        catch
                                        {
                                            Thread.Sleep(1000);
                                            count++;
                                        }
                                    }
                                    var saveExcel = Session.FindElementsByName("Maximizar");
                                    var cantidad = saveExcel.Count;
                                    if (cantidad == 2)
                                    {
                                        saveExcel[1].Click();
                                    }
                                    else
                                    {
                                        saveExcel[0].Click();
                                    }
                                    //Cambio Jose
                                    Session.FindElementByName("Pestaña Archivo").Click();
                                    //Fin Cambio Jose

                                    //Session.FindElementByName("Guardar").Click();
                                    Session.FindElementByName("Guardar").Click();
                                    Session.FindElementByName("Examinar").Click();
                                    Thread.Sleep(1000);
                                    Session.Keyboard.SendKeys(ruta);
                                    Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                    Thread.Sleep(1000);
                                    newFunctions_4.ScreenshotNuevo("Reporte Excel Guardado", true, file);
                                    LimpiarProcesos();
                                }
                            }

                        }
                        catch
                        {
                            Thread.Sleep(1000);
                            /*string navigator = AbrirPrograma.SelectNavigator();
                            if (navigator == "Edge")
                            {
                                newFunctions_4.ScreenshotNuevo("Opción de Excel no trae registros", true, file);
                                PruebaCRUD.VentanaAzul(Session);
                                Thread.Sleep(2000);
                            }*/
                        }

                        if (Session == null)
                        {
                            Thread.Sleep(1000);
                            newFunctions_4.ScreenshotNuevo("Opción de Excel no trae registros", true, file);
                            Thread.Sleep(1000);
                            Session = RootSession();
                            PruebaCRUD.VentanaAzul(Session);
                            Thread.Sleep(500);
                        }

                    }
                }
            }
            catch
            {
                Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                errorMessages.Add($"Error al intentar generar el excel");
            }

            return errorMessages;
        }

        public static List<Point> coordinatesFinder(WindowsDriver<WindowsElement> session)
        {
            Screenshot image = ((ITakesScreenshot)session).GetScreenshot();
            string name = "ProgramaExcel";
            string path = "C:\\imagenesTest\\" + name + ".Png";
            Image imgSource;
            try
            {
                image.SaveAsFile(string.Format("C:\\imagenesTest\\{0}.Png", name), ScreenshotImageFormat.Png);
                Thread.Sleep(1000);
                imgSource = Image.FromFile(path);
            }
            catch (Exception e)
            {
                image.SaveAsFile(string.Format("C:\\imagenesTest\\{0}.Png", name), ScreenshotImageFormat.Png);
                Thread.Sleep(1000);
                imgSource = Image.FromFile(path);
            }

            Point puntos = new Point();
            int lastx = 0;
            int lasty = 0;
            List<Point> coordenadas = new List<Point>();
            Bitmap bmpSource = new Bitmap(imgSource);
            Bitmap bmpDest = new Bitmap(bmpSource.Width, bmpSource.Height);
            int romper = 0;
            for (int y = 0; y < bmpSource.Height; y++)
            {
                for (int x = 0; x < bmpSource.Width; x++)
                {
                    Color clrPixel = bmpSource.GetPixel(x, y);
                    if (clrPixel.R == 151 && clrPixel.G == 179 && clrPixel.B == 143)
                    {
                        puntos.X = x;
                        puntos.Y = y;
                        if (x > lastx + 10 || y > lasty + 10)
                        {
                            coordenadas.Add(puntos);
                            lastx = x;
                            lasty = y;
                            romper = 1;
                            break;
                        }
                    }
                }
                if (romper == 1)
                {
                    break;
                }
            }
            return coordenadas;
        }

        public static List<string> ExpExcel(WindowsDriver<WindowsElement> desktopSession, string banExcel, string file, string codProgram, string ruta)
        {
            List<string> errorMessages = new List<string>();
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(10),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            try
            {
                wait.Until(driver =>
                {
                    WindowsDriver<WindowsElement> rootSession = null;
                    List<Point> coordenadas = coordinatesFinder(desktopSession);
                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));

                    string navigator = AbrirPrograma.SelectNavigator();
                    //EDEGE
                    var ExcelOption1 = desktopSession.FindElementsByClassName("TKNavegador");

                    if (navigator == "IE")
                    {
                        //IE
                        string ClickButtonExcel = "//Pane[contains(@ClassName, 'TKNavegador')]/..";
                        ExcelOption1 = desktopSession.FindElementsByXPath(ClickButtonExcel);
                        Thread.Sleep(2000);

                        //WindowsDriver<WindowsElement> RootSession3 = null;
                        //RootSession3 = PruebaCRUD.RootSession();
                        //RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        //Thread.Sleep(2000);
                        //RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        //Thread.Sleep(2000);

                        //rootSession = RootSession();
                        //rootSession = ReloadSession1(rootSession, "TForm");
                        //Thread.Sleep(1000);
                                                
                        //rootSession.FindElementByName("Exportación estándar").Click();
                        //Thread.Sleep(1000);
                        //newFunctions_4.ScreenshotNuevo("Opción Excel", true, file);
                        //Thread.Sleep(1000);
                        //rootSession.FindElementByName("Aceptar").Click();
                        //Thread.Sleep(4000);

                    }

                    

                    if (ExcelOption1.Count > 0)
                    {
                        desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                        desktopSession.Mouse.Click(null);
                        newFunctions_4.ScreenshotNuevo("Opciones Exportar Excel", true, file);

                        
                        bool optionExcel = false;
                        int cont = 0;
                        string clase = "";
                        if (banExcel == "0")
                        {
                            WindowsDriver<WindowsElement> RootSession3 = null;
                            RootSession3 = PruebaCRUD.RootSession();
                            RootSession3 = ReloadSession(RootSession3, "TForm");

                            RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            Thread.Sleep(2000);
                            newFunctions_4.ScreenshotNuevo("Exportacion Estandar", true, file);
                            RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            Thread.Sleep(2000);

                            Thread.Sleep(1000);
                            try
                            {
                                generarExcel(rootSession, file, codProgram, ruta);
                                rootSession = RootSession();
                            }
                            catch
                            {
                                PruebaCRUD.VentanaAzul(desktopSession);
                                rootSession = RootSession();
                            }
                        }

                        else if (banExcel == "1")
                        {
                            rootSession = RootSession();
                            rootSession = ReloadSession1(rootSession, "TForm");
                        }
                        else if (banExcel == "2")
                        {
                            rootSession = RootSession();
                            rootSession = ReloadSession1(rootSession, "TFrmMenuExcel");
                        }
                        else if (banExcel == "3")
                        {

                            rootSession = RootSession();
                            rootSession = ReloadSession1(rootSession, "TFrmMenExpo");
                        }
                        else if (banExcel == "4")
                        {

                            rootSession = RootSession();
                            rootSession = ReloadSession1(rootSession, "TFrmTipExpo");
                        }

                        if (banExcel == "1" || banExcel == "2" || banExcel == "3" || banExcel == "4")
                        {
                            ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                            var contButton = rootSession.FindElementsByClassName("TGroupButton");
                            int cont2 = contButton.Count() - 1;
                            rootSession.FindElementsByClassName("TGroupButton");
                            rootSession.FindElementsByClassName("TGroupButton")[cont2].Click();
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            try
                            {
                                generarExcel(rootSession, file, codProgram, ruta);
                            }
                            catch
                            {
                                Thread.Sleep(2000);
                                PruebaCRUD.VentanaAzul(desktopSession);
                                Thread.Sleep(2000);
                            }
                            int i;
                            //Debugger.Launch();
                            for (i = cont2 - 1; i > -1; i--)
                            {
                                ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                                desktopSession.Mouse.Click(null);
                                if (banExcel == "1")
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                else
                                {
                                    rootSession.FindElementsByClassName("TGroupButton")[cont2 - 1].Click();
                                }
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                Thread.Sleep(1000);
                                newFunctions_4.ScreenshotNuevo("Opción Excel", true, file);
                                Thread.Sleep(1000);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                Thread.Sleep(1000);
                                try
                                {
                                    generarExcel(rootSession, file, codProgram, ruta);
                                }
                                catch
                                {
                                    Thread.Sleep(7000);
                                    PruebaCRUD.VentanaAzul(desktopSession);
                                    Thread.Sleep(2000);
                                }
                                cont2 = cont2 - 1;
                            }
                        }

                        if (banExcel == "5")
                        {
                            rootSession = RootSession();
                            rootSession = ReloadSession1(rootSession, "TForm");
                            Thread.Sleep(1000);
                            rootSession.FindElementByName("Exportación estándar").Click();
                            Thread.Sleep(1000);
                            newFunctions_4.ScreenshotNuevo("Opción Excel", true, file);
                            Thread.Sleep(1000);
                            rootSession.FindElementByName("Aceptar").Click();
                            Thread.Sleep(4000);
                            rootSession = RootSession();
                            rootSession = ReloadSession1(rootSession, "XLMAIN");
                            Thread.Sleep(3000);
                            //Debugger.Launch();
                            if (rootSession == null)
                            {
                                Thread.Sleep(1000);
                                newFunctions_4.ScreenshotNuevo("No hay registros para exportar o el generar excel", true, file);
                                Thread.Sleep(2000);
                                PruebaCRUD.VentanaAzul(desktopSession);
                                Thread.Sleep(2000);
                            }
                            else
                            {
                                generarExcel(rootSession, file, codProgram, ruta);
                            }
                            for (int i = 1; i < 16; i++)
                            {
                                //Clic Excel
                                ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                                desktopSession.Mouse.Click(null);
                                Thread.Sleep(1000);
                                rootSession = RootSession();
                                rootSession = ReloadSession1(rootSession, "TForm");
                                Thread.Sleep(1000);
                                rootSession.FindElementByName("Exportación estándar").Click();
                                Thread.Sleep(1000);
                                for (int j = 0; j < i; j++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                Thread.Sleep(1000);
                                newFunctions_4.ScreenshotNuevo("Opción Excel", true, file);
                                Thread.Sleep(1000);
                                rootSession.FindElementByName("Aceptar").Click();
                                Thread.Sleep(4000);
                                rootSession = RootSession();
                                rootSession = ReloadSession1(rootSession, "XLMAIN");
                                if (rootSession == null)
                                {
                                    Thread.Sleep(1000);
                                    newFunctions_4.ScreenshotNuevo("No hay registros para exportar o el generar excel", true, file);
                                    Thread.Sleep(2000);
                                    PruebaCRUD.VentanaAzul(desktopSession);
                                    Thread.Sleep(2000);
                                }
                                else
                                {
                                    generarExcel(rootSession, file, codProgram, ruta);
                                }

                            }
                        }

                    }
                    return rootSession != null;
                });
            }

            catch
            {
                Debug.WriteLine("No se encuentra la opcion de excel para este programa o no hay registros para exportar");
            }
            //string pathImg = "C:\\imagenesTest\\Programa.Png";
            //if (File.Exists(pathImg))
            //{
            //    File.Delete(pathImg);
            //}
            return errorMessages;
        }

        public static List<string> ExpExcel1(WindowsDriver<WindowsElement> desktopSession, string banExcel, string file, string codProgram, string ruta)
        {
            List<string> errorMessages = new List<string>();
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(10),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            try
            {
                wait.Until(driver =>
                {
                    WindowsDriver<WindowsElement> rootSession = null;
                    List<Point> coordenadas = coordinatesFinder(desktopSession);
                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));

                    string navigator = AbrirPrograma.SelectNavigator();
                    //EDEGE
                    var ExcelOption1 = desktopSession.FindElementsByClassName("TKNavegador");

                    if (navigator == "IE")
                    {
                        //IE
                        string ClickButtonExcel = "//Pane[contains(@ClassName, 'TKNavegador')]/..";
                        ExcelOption1 = desktopSession.FindElementsByXPath(ClickButtonExcel);
                        Thread.Sleep(2000);

                        //WindowsDriver<WindowsElement> RootSession3 = null;
                        //RootSession3 = PruebaCRUD.RootSession();
                        //RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        //Thread.Sleep(2000);
                        //RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        //Thread.Sleep(2000);

                        //rootSession = RootSession();
                        //rootSession = ReloadSession1(rootSession, "TForm");
                        //Thread.Sleep(1000);

                        //rootSession.FindElementByName("Exportación estándar").Click();
                        //Thread.Sleep(1000);
                        //newFunctions_4.ScreenshotNuevo("Opción Excel", true, file);
                        //Thread.Sleep(1000);
                        //rootSession.FindElementByName("Aceptar").Click();
                        //Thread.Sleep(4000);

                    }



                    if (ExcelOption1.Count > 0)
                    {
                        desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                        desktopSession.Mouse.Click(null);
                        newFunctions_4.ScreenshotNuevo("Opciones Exportar Excel", true, file);


                        bool optionExcel = false;
                        int cont = 0;
                        string clase = "";
                        if (banExcel == "0")
                        {
                            WindowsDriver<WindowsElement> RootSession3 = null;
                            RootSession3 = PruebaCRUD.RootSession();
                            RootSession3 = ReloadSession(RootSession3, "TForm");

                            RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                            Thread.Sleep(2000);
                            newFunctions_4.ScreenshotNuevo("Exportacion con Detalle", true, file);
                            RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            Thread.Sleep(2000);
                            RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            Thread.Sleep(2000);

                            Thread.Sleep(1000);
                            try
                            {
                                generarExcel(rootSession, file, codProgram, ruta);
                                rootSession = RootSession();
                            }
                            catch
                            {
                                PruebaCRUD.VentanaAzul(desktopSession);
                                rootSession = RootSession();
                            }
                        }

                        else if (banExcel == "1")
                        {
                            rootSession = RootSession();
                            rootSession = ReloadSession1(rootSession, "TForm");
                        }
                        else if (banExcel == "2")
                        {
                            rootSession = RootSession();
                            rootSession = ReloadSession1(rootSession, "TFrmMenuExcel");
                        }
                        else if (banExcel == "3")
                        {

                            rootSession = RootSession();
                            rootSession = ReloadSession1(rootSession, "TFrmMenExpo");
                        }
                        else if (banExcel == "4")
                        {

                            rootSession = RootSession();
                            rootSession = ReloadSession1(rootSession, "TFrmTipExpo");
                        }

                        if (banExcel == "1" || banExcel == "2" || banExcel == "3" || banExcel == "4")
                        {
                            ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                            var contButton = rootSession.FindElementsByClassName("TGroupButton");
                            int cont2 = contButton.Count() - 1;
                            rootSession.FindElementsByClassName("TGroupButton");
                            rootSession.FindElementsByClassName("TGroupButton")[cont2].Click();
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            try
                            {
                                generarExcel(rootSession, file, codProgram, ruta);
                            }
                            catch
                            {
                                Thread.Sleep(2000);
                                PruebaCRUD.VentanaAzul(desktopSession);
                                Thread.Sleep(2000);
                            }
                            int i;
                            //Debugger.Launch();
                            for (i = cont2 - 1; i > -1; i--)
                            {
                                ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                                desktopSession.Mouse.Click(null);
                                if (banExcel == "1")
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                else
                                {
                                    rootSession.FindElementsByClassName("TGroupButton")[cont2 - 1].Click();
                                }
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                Thread.Sleep(1000);
                                newFunctions_4.ScreenshotNuevo("Opción Excel", true, file);
                                Thread.Sleep(1000);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                Thread.Sleep(1000);
                                try
                                {
                                    generarExcel(rootSession, file, codProgram, ruta);
                                }
                                catch
                                {
                                    Thread.Sleep(7000);
                                    PruebaCRUD.VentanaAzul(desktopSession);
                                    Thread.Sleep(2000);
                                }
                                cont2 = cont2 - 1;
                            }
                        }

                        if (banExcel == "5")
                        {
                            rootSession = RootSession();
                            rootSession = ReloadSession1(rootSession, "TForm");
                            Thread.Sleep(1000);
                            rootSession.FindElementByName("Exportación estándar").Click();
                            Thread.Sleep(1000);
                            newFunctions_4.ScreenshotNuevo("Opción Excel", true, file);
                            Thread.Sleep(1000);
                            rootSession.FindElementByName("Aceptar").Click();
                            Thread.Sleep(4000);
                            rootSession = RootSession();
                            rootSession = ReloadSession1(rootSession, "XLMAIN");
                            Thread.Sleep(3000);
                            //Debugger.Launch();
                            if (rootSession == null)
                            {
                                Thread.Sleep(1000);
                                newFunctions_4.ScreenshotNuevo("No hay registros para exportar o el generar excel", true, file);
                                Thread.Sleep(2000);
                                PruebaCRUD.VentanaAzul(desktopSession);
                                Thread.Sleep(2000);
                            }
                            else
                            {
                                generarExcel(rootSession, file, codProgram, ruta);
                            }
                            for (int i = 1; i < 16; i++)
                            {
                                //Clic Excel
                                ruta = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora();
                                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                                desktopSession.Mouse.Click(null);
                                Thread.Sleep(1000);
                                rootSession = RootSession();
                                rootSession = ReloadSession1(rootSession, "TForm");
                                Thread.Sleep(1000);
                                rootSession.FindElementByName("Exportación estándar").Click();
                                Thread.Sleep(1000);
                                for (int j = 0; j < i; j++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                Thread.Sleep(1000);
                                newFunctions_4.ScreenshotNuevo("Opción Excel", true, file);
                                Thread.Sleep(1000);
                                rootSession.FindElementByName("Aceptar").Click();
                                Thread.Sleep(4000);
                                rootSession = RootSession();
                                rootSession = ReloadSession1(rootSession, "XLMAIN");
                                if (rootSession == null)
                                {
                                    Thread.Sleep(1000);
                                    newFunctions_4.ScreenshotNuevo("No hay registros para exportar o el generar excel", true, file);
                                    Thread.Sleep(2000);
                                    PruebaCRUD.VentanaAzul(desktopSession);
                                    Thread.Sleep(2000);
                                }
                                else
                                {
                                    generarExcel(rootSession, file, codProgram, ruta);
                                }

                            }
                        }

                    }
                    return rootSession != null;
                });
            }

            catch
            {
                Debug.WriteLine("No se encuentra la opcion de excel para este programa o no hay registros para exportar");
            }
            //string pathImg = "C:\\imagenesTest\\Programa.Png";
            //if (File.Exists(pathImg))
            //{
            //    File.Delete(pathImg);
            //}
            return errorMessages;
        }
    }
}
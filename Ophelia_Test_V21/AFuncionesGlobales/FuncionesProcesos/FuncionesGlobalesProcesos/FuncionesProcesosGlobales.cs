using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Appium;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using OpenQA.Selenium;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System.IO;

namespace Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesGlobalesProcesos
{
    class FuncionesProcesosGlobales : FuncionesVitales
    {
        public static List<string> AgregarContrato(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file)
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            try
            {
                var TDBEdit = desktopSession.FindElementsByClassName("TDBEdit");
                var Edit = desktopSession.FindElementsByClassName("Edit");
                string NomPestaña = "";

                //Se envia parametro al campo Identificacion
                TDBEdit[10].SendKeys(crudVars[0]);
                TDBEdit[10].SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                //Se envia parametro al campo Nro Contrato
                TDBEdit[9].SendKeys(crudVars[0]);
                TDBEdit[9].SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                //Pestaña Tipo Contrato
                if (crudVars[1] == "Término Fijo")
                {
                    var TipoContr = desktopSession.FindElementsByName(crudVars[1]);
                    //Se envia parametro segun tipo de contrato
                    TipoContr[0].Click();
                    TipoContr[0].Click();
                    //AGREGAR MENSAJE DE ERROR PARA QUE TENGA FECHA DE VENCIMIENTO DEL CONTRATO
                }
                Thread.Sleep(500);
                if (crudVars[2] == "Procedimiento 2")
                {
                    //Se envia parametro al campo Tipo Retencion
                    Edit[4].Click();
                    desktopSession.Keyboard.SendKeys(Keys.ArrowDown);
                }
                Thread.Sleep(500);
                ////Pestaña Regimen Cesantias
                NomPestaña = "Regimen Cesantias";
                var Pestaña = desktopSession.FindElementsByName(NomPestaña);
                if (crudVars[14] == "Tradicional" || crudVars[14] == "No-Aplica" || crudVars[14] == "Consignación Mensual")
                {
                    Pestaña[0].Click();
                    Pestaña[0].Click();
                    Thread.Sleep(500);
                    var TipoRegimen = desktopSession.FindElementsByName(crudVars[14]);
                    TipoRegimen[0].Click();
                    TipoRegimen[0].Click();
                }
                Thread.Sleep(500);

                //////////Pestaña Seguridad Social
                ////NomPestaña = "Seguridad Social";
                ////Pestaña = desktopSession.FindElementsByName(NomPestaña);
                ////Pestaña[0].Click();
                ////Pestaña[0].Click();
                //////crudVars[3]  ///Tipo Cotizante
                ////Thread.Sleep(500);

                //////Pestaña Tipo Salario
                NomPestaña = "Tipo Salario";
                Pestaña = desktopSession.FindElementsByName(NomPestaña);
                //Se envia parametro segun tupo de salario
                if (crudVars[4] == "Integral" || crudVars[4] == "Variable")
                {
                    Pestaña[0].Click();
                    Pestaña[0].Click();
                    Thread.Sleep(500);
                    var TipoSalario = desktopSession.FindElementsByName(crudVars[4]);
                    TipoSalario[0].Click();
                    TipoSalario[0].Click();
                }

                //////Pestaña Fechas Contrato
                NomPestaña = "Fechas Contrato";
                Pestaña = desktopSession.FindElementsByName(NomPestaña);
                Pestaña[0].Click();
                Pestaña[0].Click();
                Thread.Sleep(500);
                TDBEdit = desktopSession.FindElementsByClassName("TDBEdit");
                TDBEdit[27].SendKeys(crudVars[5]);
                Thread.Sleep(1000);
                if (crudVars[1] == "Término Fijo")
                {
                    //Se envia parametro fecha de vencimiento
                    TDBEdit[18].Click();
                    Thread.Sleep(5000);
                    TDBEdit[18].Click();
                    Thread.Sleep(500);
                    TDBEdit[18].SendKeys(crudVars[12]);
                }
                Thread.Sleep(1000);

                //////Pestaña Condiciones de Pago
                NomPestaña = "Condiciones de Pago";
                Pestaña = desktopSession.FindElementsByName(NomPestaña);
                Pestaña[0].Click();
                Pestaña[0].Click();
                Thread.Sleep(500);
                Pestaña[0].Click();
                Thread.Sleep(500);
                TDBEdit = desktopSession.FindElementsByClassName("TDBEdit");
                //Se envia parametro cargo
                TDBEdit[26].SendKeys(crudVars[6]);
                Thread.Sleep(500);
                TDBEdit[25].Click();
                Thread.Sleep(500);
                var ApplicationWindow = desktopSession.FindElementByClassName("IEFrame");
                new Actions(desktopSession).MoveToElement(ApplicationWindow, 0, 0).MoveByOffset((ApplicationWindow.Size.Width / 2) + 90, (ApplicationWindow.Size.Height / 2) + 70).DoubleClick().Perform();
                Thread.Sleep(1000);
                //Se envia parametro Centro Costo
                TDBEdit[25].Click();
                TDBEdit[25].Click();
                TDBEdit[25].SendKeys(crudVars[7]);
                Thread.Sleep(1000);
                //Se envia parametro Centro Trabajo
                TDBEdit[24].Click();
                TDBEdit[24].Click();
                TDBEdit[24].SendKeys(crudVars[8]);
                Thread.Sleep(1000);
                //Se envia parametro Grupo Prototipo
                TDBEdit[22].Click();
                TDBEdit[22].Click();
                TDBEdit[22].SendKeys(crudVars[9]);
                Thread.Sleep(1000);
                if (crudVars[1] == "Término Fijo")
                {
                    TDBEdit[29].Click();
                    Thread.Sleep(1000);
                    new Actions(desktopSession).MoveToElement(ApplicationWindow, 0, 0).MoveByOffset((ApplicationWindow.Size.Width / 2) + 60, (ApplicationWindow.Size.Height / 2) + 30).DoubleClick().Perform();
                    Thread.Sleep(1000);
                    //Se envia parametro Clase Nomina
                    List<string> crudVarsoff = new List<string> { crudVars[10] };
                    List<int> lupasOffIndex = new List<int> { 0, 1, 2, 3, 4, 5, 6, 8};
                    newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, crudVarsoff, lupasOffIndex);
                    Thread.Sleep(1000);
                }
                else
                {
                    //Se envia parametro Clase Nomina
                    TDBEdit[23].Click();
                    TDBEdit[23].Click();
                    TDBEdit[23].SendKeys(crudVars[10]);
                    Thread.Sleep(1000);
                }
                //Se envia parametro Sueldo
                TDBEdit[29].Click();
                TDBEdit[29].Click();
                TDBEdit[29].SendKeys(crudVars[11]);

                //////Pestaña Clasificación
                NomPestaña = "Clasificación";
                Pestaña = desktopSession.FindElementsByName(NomPestaña);
                Pestaña[0].Click();
                Pestaña[0].Click();
                Thread.Sleep(500);
                TDBEdit = desktopSession.FindElementsByClassName("TDBEdit");
                TDBEdit[13].SendKeys(crudVars[13]);

                //AGREGAR REGISTRO
                var ElementList = PruebaCRUD.NavClass(desktopSession);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 215, 15);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(5000);
                //Cerrar Ventana Error
                ApplicationWindow = desktopSession.FindElementByClassName("IEFrame");
                new Actions(desktopSession).MoveToElement(ApplicationWindow, 0, 0).MoveByOffset((ApplicationWindow.Size.Width / 2), (ApplicationWindow.Size.Height / 2) + 110).DoubleClick().Perform();
                Thread.Sleep(1000);

            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;
        }
        public static List<string> AgregarNovedad(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file)
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            try
            {
                var TDBEdit = desktopSession.FindElementsByClassName("TDBEdit");
                //Se envia parametro identificacion
                TDBEdit[16].SendKeys(crudVars[0]);
                TDBEdit[16].SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                //Se envia parametro concepto
                TDBEdit[3].SendKeys(crudVars[1]);
                TDBEdit[3].SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                ///     recargo nocturno           Dominical               HE diurnas              HE nocturnas         HE festivas diurnas     HE festivas nocturnas  Recargo nocturno festivo
                if (crudVars[1] == "1075" || crudVars[1] == "1066" || crudVars[1] == "1060" || crudVars[1] == "1063" || crudVars[1] == "1069" || crudVars[1] == "1072" || crudVars[1] == "1081")
                {
                    //Se envia parametro cantidad
                    TDBEdit[10].Click();
                    Thread.Sleep(1000);
                    TDBEdit[10].SendKeys(crudVars[3]);
                    TDBEdit[10].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(1000);
                } else if (crudVars[1] == "1274"){
                    //Se agregan los concpetos que necesitan solo cuotas
                    //Se envia parametro Valor cuota
                    TDBEdit[11].Click();
                    Thread.Sleep(500);
                    TDBEdit[11].Clear();
                    Thread.Sleep(500);
                    TDBEdit[11].SendKeys(crudVars[2]);
                    TDBEdit[11].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(3000);
                }
                else
                {   
                    //Se envia parametro Valor cuota y Cantidad
                    TDBEdit[10].Click();
                    Thread.Sleep(500);
                    TDBEdit[10].SendKeys(crudVars[3]);
                    TDBEdit[10].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(3000);

                    TDBEdit[11].Click();
                    Thread.Sleep(500);
                    TDBEdit[11].Clear();
                    Thread.Sleep(500);
                    TDBEdit[11].SendKeys(crudVars[2]);
                    TDBEdit[11].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(3000);
                }
                Thread.Sleep(500);
                TDBEdit[5].Clear();
                TDBEdit[5].SendKeys(crudVars[4]);
                //TDBEdit[11].SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(3000);
                TDBEdit[6].Clear();
                TDBEdit[6].SendKeys(crudVars[5]);
                //TDBEdit[10].SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(3000);

                ////ACEPTAR REGISTRO
                var ElementList = PruebaCRUD.NavClass(desktopSession);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 215, 15);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(3000);
                ////Debugger.Launch();
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;
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
        static public WindowsDriver<WindowsElement> RootSession()
        {
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }
        public static List<string> AgregarAusentismo(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file)
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            try
            {
                var TDBEdit = desktopSession.FindElementsByClassName("TDBEdit");
                //Se envia parametro identificacion
                TDBEdit[24].SendKeys(crudVars[0]);
                Thread.Sleep(500);
                //Se envia parametro concepto
                TDBEdit[14].SendKeys(crudVars[1]);
                Thread.Sleep(500);
                //Se envia motivo ausencia
                TDBEdit[13].SendKeys(crudVars[2]);
                Thread.Sleep(500);
                //Seleccion Tipo de Ausentismo
                var TipodeAusentismo = desktopSession.FindElementsByName(crudVars[3]);
                TipodeAusentismo[0].Click();
                Thread.Sleep(500);
                //Ventana de fechas TBitBtn
                var ButtonFechas = desktopSession.FindElementsByClassName("TBitBtn");
                ButtonFechas[2].Click();
                    //Cambio de Frame TfrmCambioFechas
                    
                WindowsDriver<WindowsElement> rootSession = null;
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TfrmCambioFechas");
                    //Rellenar las fechas Inicio y Final 
                    var TDBEditFechas = rootSession.FindElementsByClassName("TEdit");
                    TDBEditFechas[1].SendKeys(crudVars[4]);
                    TDBEditFechas[0].SendKeys(crudVars[5]);
                    //Click Aceptar
                    var BotonAceptar = rootSession.FindElementsByName("Aceptar");
                    BotonAceptar[0].Click();
                //AGREGAR REGISTRO
                var ElementList = PruebaCRUD.NavClass(desktopSession);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 215, 15);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(5000);


            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;
        }
        
        public static List<string> LiquidarNomina(WindowsDriver<WindowsElement> desktopSession, string tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            try
            {
                var Elementlist = desktopSession.FindElementsByClassName("TEdit");

                switch (tipo)
                {
                    //Se insertan las fechas de liquidacion y de pago
                    case ("0"):
                        Elementlist[1].SendKeys(data[0]);
                        Thread.Sleep(1000);
                        Elementlist[0].SendKeys(data[1]);
                        break;
                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;
        }

        public static List<string> ValidacionesProcesos(WindowsDriver<WindowsElement> desktopSession, string IdInicial, string IdFinal, int cont, List<string> NomCamposValidarProcesoP, List<string> VarValorGlobalProcesoP, List<string> SelNumCamposTabProceso, string IdentificacionM)
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            int IdInicialInt = Convert.ToInt32(IdInicial);
            int IdFinalInt = Convert.ToInt32(IdFinal);
            int IdentificacionInt = Convert.ToInt32(IdentificacionM);
            //int CantidadDataValidarInt = Convert.ToInt32(CantidadDataValidar); 
            bool IsDisplayedValue = false;
            try
            {
                var TDBGridInplaceEdit = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
                TDBGridInplaceEdit[0].Click();
                Thread.Sleep(1000);
                //Valida la posición donde arranca
                if (cont == 0)
                {
                    for (int i = 0; i < 500; i++)
                    {
                        string Identificacion = TDBGridInplaceEdit[0].Text;
                        IdentificacionInt = Convert.ToInt32(Identificacion);
                        if (IdentificacionInt >= IdInicialInt && IdentificacionInt <= IdFinalInt)
                        {
                            IsDisplayedValue = true;
                            break;
                        }
                        else
                        {
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                            Thread.Sleep(2000);
                        }
                    }
                }
                if (cont != 0)
                {
                    IsDisplayedValue = true;
                }
                //Console.WriteLine(IsDisplayedValue);
                if (IsDisplayedValue)
                {
                    if (cont != 0)
                    {
                        Thread.Sleep(2000);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                        Thread.Sleep(500);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Home);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Home);
                    }

                    //Valida identificacion
                    bool IsIdentification = false;
                    while (IsIdentification == false)
                    {
                        string IdentificacionRealP = TDBGridInplaceEdit[0].Text;
                        int IdentificacionRealIntP = Convert.ToInt32(IdentificacionRealP);
                        if (IdentificacionInt == IdentificacionRealIntP)
                        {
                            IsIdentification = true;
                            break;
                        }
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                        Thread.Sleep(1500);
                    }
                    if (IsIdentification)
                    {
                        string IdentificacionReal = TDBGridInplaceEdit[0].Text;
                        int IdentificacionRealInt = Convert.ToInt32(IdentificacionReal);
                        if (IdentificacionRealInt >= IdInicialInt && IdentificacionRealInt <= IdFinalInt)
                        {

                            for (int j = 0; j < SelNumCamposTabProceso.Count; j++)
                            {
                                var Value = SelNumCamposTabProceso[j];
                                if (Value.Length == 2)
                                {
                                    var FirstValue = Value[0];
                                    var SecondValue = Value[1];
                                    int ValueTab = (int)Char.GetNumericValue(FirstValue);
                                    MoverEnGrilla1(desktopSession, ValueTab, SecondValue);
                                }
                                else if (Value.Length == 3) 
                                {
                                    //Debugger.Launch();
                                    var FirstValue = Value[0];
                                    var SecondValue = Value[1];
                                    int ValueTab = (int)Char.GetNumericValue(FirstValue) * 10 + (int)Char.GetNumericValue(SecondValue);
                                    var ThirdValue = Value[2];
                                    MoverEnGrilla1(desktopSession, ValueTab, ThirdValue);
                                }
                                else
                                {
                                    var FirstValue = Value[0];
                                    var SecondValue = Value[1];
                                    int ValueTab = (int)Char.GetNumericValue(FirstValue);
                                    var ThreeValue = Value[2];
                                    var FourValue = Value[3];
                                    int ValueTabS = (int)Char.GetNumericValue(ThreeValue);
                                    MoverEnGrilla1(desktopSession, ValueTab, SecondValue);
                                    MoverEnGrilla2(desktopSession, ValueTabS, FourValue);
                                }

                                //Validacion Grilla
                                Thread.Sleep(1000);
                                string ValueObt = TDBGridInplaceEdit[0].Text;
                                Thread.Sleep(1000);
                                if (ValueObt != VarValorGlobalProcesoP[j])
                                {
                                    errorMessages.Add($"El {NomCamposValidarProcesoP[j]} del usuario {IdentificacionReal} es incorrecto el valor esperado es: {VarValorGlobalProcesoP[j]} y el encontrado es: {ValueObt}");
                                }
                                Thread.Sleep(2000);
                            }
                        }
                        else
                        {
                            errorMessages.Add($"Valor {IdentificacionRealInt} no se encuentra dentro del rango estipulado");
                        }
                    }

                }
                else
                {
                    errorMessages.Add($"No puede encontrar valor dentro de la grilla entre ese rango de identificación");
                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;

        }

        //public static List<string> ValidacionesProcesoslices(WindowsDriver<WindowsElement> desktopSession, string IdInicial, string IdFinal, int cont, List<string> NomCamposValidarProcesoP, List<string> VarValorGlobalProcesoP, List<string> SelNumCamposTabProceso, string IdentificacionM, string CodPrograma)
        //{
        //    List<string> errorMessages = new List<string>();
        //    List<string> ErrorValidacion = new List<string>();
        //    List<string> errors = new List<string>();
        //    int IdInicialInt = Convert.ToInt32(IdInicial);
        //    int IdFinalInt = Convert.ToInt32(IdFinal);
        //    int IdentificacionInt = Convert.ToInt32(IdentificacionM);
        //    //int CantidadDataValidarInt = Convert.ToInt32(CantidadDataValidar); 
        //    bool IsDisplayedValue = false;
        //    try
        //    {
                

        //        var TDBGridInplaceEdit = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
        //        TDBGridInplaceEdit[0].Click();
        //        if (CodPrograma == "KNmCesan") { desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight); }
        //        Thread.Sleep(1000);
        //        //Valida la posición donde arranca
        //        if (cont == 0)
        //        {
        //            for (int i = 0; i < 500; i++)
        //            {
        //                string Identificacion = TDBGridInplaceEdit[0].Text;
        //                IdentificacionInt = Convert.ToInt32(Identificacion);
        //                if (IdentificacionInt >= IdInicialInt && IdentificacionInt <= IdFinalInt)
        //                {
        //                    IsDisplayedValue = true;
        //                    break;
        //                }
        //                else
        //                {
        //                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
        //                    Thread.Sleep(2000);
        //                }
        //            }
        //        }
        //        if (cont != 0)
        //        {
        //            IsDisplayedValue = true;
        //        }
        //        //Console.WriteLine(IsDisplayedValue);
        //        if (IsDisplayedValue)
        //        {
        //            if (cont != 0)
        //            {
        //                Thread.Sleep(2000);
        //                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
        //                Thread.Sleep(500);
        //                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Home);
        //                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Home);
        //            }

        //            //Valida identificacion
        //            bool IsIdentification = false;
        //            while (IsIdentification == false)
        //            {
        //                string IdentificacionRealP = TDBGridInplaceEdit[0].Text;
        //                int IdentificacionRealIntP = Convert.ToInt32(IdentificacionRealP);
        //                if (IdentificacionInt == IdentificacionRealIntP)
        //                {
        //                    IsIdentification = true;
        //                    break;
        //                }
        //                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
        //                Thread.Sleep(1500);
        //            }
        //            if (IsIdentification)
        //            {
        //                string IdentificacionReal = TDBGridInplaceEdit[0].Text;
        //                int IdentificacionRealInt = Convert.ToInt32(IdentificacionReal);
        //                if (IdentificacionRealInt >= IdInicialInt && IdentificacionRealInt <= IdFinalInt)
        //                {

        //                    for (int j = 0; j < SelNumCamposTabProceso.Count; j++)
        //                    {
        //                        var Value = SelNumCamposTabProceso[j];
        //                        if (Value.Length == 2)
        //                        {
        //                            var FirstValue = Value[0];
        //                            var SecondValue = Value[1];
        //                            int ValueTab = (int)Char.GetNumericValue(FirstValue);
        //                            MoverEnGrilla1(desktopSession, ValueTab, SecondValue);
        //                        }
        //                        else
        //                        {
        //                            var FirstValue = Value[0];
        //                            var SecondValue = Value[1];
        //                            int ValueTab = (int)Char.GetNumericValue(FirstValue);
        //                            var ThreeValue = Value[2];
        //                            var FourValue = Value[3];
        //                            int ValueTabS = (int)Char.GetNumericValue(ThreeValue);
        //                            MoverEnGrilla1(desktopSession, ValueTab, SecondValue);
        //                            MoverEnGrilla2(desktopSession, ValueTabS, FourValue);
        //                        }

        //                        //Validacion Grilla
        //                        Thread.Sleep(1000);
        //                        string ValueObt = TDBGridInplaceEdit[0].Text;
        //                        Thread.Sleep(1000);
        //                        if (ValueObt != VarValorGlobalProcesoP[j])
        //                        {
        //                            errorMessages.Add($"El {NomCamposValidarProcesoP[j]} del usuario {IdentificacionReal} es incorrecto el valor esperado es: {VarValorGlobalProcesoP[j]} y el encontrado es: {ValueObt}");
        //                        }
        //                        Thread.Sleep(2000);
        //                    }
        //                }
        //                else
        //                {
        //                    errorMessages.Add($"Valor {IdentificacionRealInt} no se encuentra dentro del rango estipulado");
        //                }
        //            }

        //        }
        //        else
        //        {
        //            errorMessages.Add($"No puede encontrar valor dentro de la grilla entre ese rango de identificación");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        errorMessages.Add($"Error de Appium: {ex.ToString()}");
        //    }
        //    return errorMessages;

        //}


        public static List<string> AgregarIncapacidad(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file)
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            try
            {
                var TDBEdit = desktopSession.FindElementsByClassName("TDBEdit");
                
                //Agrego Identificacion
                TDBEdit[13].SendKeys(crudVars[0]);
                TDBEdit[13].SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                //Agrego Concepto
                TDBEdit[15].SendKeys(crudVars[1]);
                TDBEdit[15].SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                //Agrego Codigo Diagnostico
                TDBEdit[17].SendKeys(crudVars[2]);
                TDBEdit[17].SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                //Agrego Numero Incapacidad
                TDBEdit[16].SendKeys(crudVars[3]);
                Thread.Sleep(500);                
                //Agrego Fecha Desde
                TDBEdit[26].SendKeys(crudVars[4]);
                Thread.Sleep(500);
                //Agrego Fecha Hasta
                TDBEdit[25].SendKeys(crudVars[5]);
                Thread.Sleep(1000);
                //Agrego Fecha Inicio Pago
                TDBEdit[23].SendKeys(crudVars[6]);
                Thread.Sleep(500);
                //Agrego Prorroga
                if (crudVars[7] == "1")
                {
                    var Checkbox = desktopSession.FindElementsByName("Indicador Prorroga");
                    Checkbox[0].Click();
                }
                //Agrego Entidad
                TDBEdit[27].SendKeys(crudVars[8]);
                TDBEdit[27].SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(3000);
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }

            return null;
        }

        public static List<string> ValidacionListaDatosProceso1(WindowsDriver<WindowsElement> desktopSession, List<string> VarValorGlobalProceso, List<string> NomCamposValidarProceso, List<string> VarValorGlobalProcesoNom)
        {
            //Seleccionado solamente los campos que se deben validar y elimina los nulos
            List<string> SelNomCamposValidarProceso = new List<string> { };
            foreach (var campo in NomCamposValidarProceso)
            {
                if (campo != "")
                {
                    SelNomCamposValidarProceso.Add(campo);
                }
            }
            return SelNomCamposValidarProceso;
        }

        public static List<string> ValidacionListaDatosProceso2(WindowsDriver<WindowsElement> desktopSession, List<string> VarValorGlobalProceso, List<string> NomCamposValidarProceso, List<string> VarValorGlobalProcesoNom, List<string> SelNomCamposValidarProceso)
        {
            //Seleccionando solamente la información que se necesita para validar los datos
            List<string> IdeVarValorGlobalProceso = new List<string> { };
            for (int i = 0; i < SelNomCamposValidarProceso.Count; i++)
            {
                for (int j = 0; j < VarValorGlobalProceso.Count; j++)
                {
                    if (SelNomCamposValidarProceso[i] == VarValorGlobalProcesoNom[j])
                    {
                        IdeVarValorGlobalProceso.Add(VarValorGlobalProceso[j]);
                    }
                }
            }
            return IdeVarValorGlobalProceso;
        }

        public static List<string> ValidacionListaDatosTab(WindowsDriver<WindowsElement> desktopSession, List<string> NumCamposTabProceso)
        {
            //Seleccionado solamente los campos que se deben dar tab
            List<string> SelNumCamposTabProceso = new List<string> { };
            foreach (var campo in NumCamposTabProceso)
            {
                if (campo != "")
                {
                    SelNumCamposTabProceso.Add(campo);
                }
            }
            return SelNumCamposTabProceso;
        }

        public static void MoverEnGrilla1(WindowsDriver<WindowsElement> desktopSession, int ValueTab, char SecondValue)
        {
            //Proceso de Tabs para 2 variables
            for (int k = 0; k < ValueTab; k++)
            {
                switch (SecondValue.ToString())
                {
                    case "T":
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        break;
                    case "D":
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                        break;
                    case "I":
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowLeft);
                        break;
                    default:
                        break;
                }
            }
        }

        public static void MoverEnGrilla2(WindowsDriver<WindowsElement> desktopSession, int ValueTabS, char FourValue)
        {
            //Proceso de Tabs para 4 variables
            for (int k = 0; k < ValueTabS; k++)
            {
                switch (FourValue.ToString())
                {
                    case "T":
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        break;
                    case "D":
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                        break;
                    case "I":
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowLeft);
                        break;
                    default:
                        break;
                }
            }
        }
    }
}


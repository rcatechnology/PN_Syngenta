using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;

namespace PN_Syngenta
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    //If you want to use an add-on identifier for the development license, you can specify an add-on identifier string as the second parameter.
                    //oApp = new Application(args[0], "XXXXX");
                    oApp = new Application(args[0]);
                }
                oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
                
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        static SAPbobsCOM.Company oCompany;

        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.FormType == 134 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
            {
                SAPbouiCOM.Form oPNForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                SAPbouiCOM.Button obutton;
                SAPbouiCOM.Item oitem;
                SAPbouiCOM.Item oNewItem;

                oNewItem = oPNForm.Items.Add("RplPN", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                obutton = (SAPbouiCOM.Button)oNewItem.Specific;
                obutton.Caption = "Criar Fornecedor";

                oitem = oPNForm.Items.Item("540002072"); // UI element in the system form to use for positional reference
                oNewItem.Top = oitem.Top;
                oNewItem.Height = oitem.Height;
                oNewItem.Width = oitem.Width;
                oNewItem.Left = oitem.Left - oitem.Width - 20;

                //oNewItem.Visible = false;
            }
            if (pVal.FormType == 134 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && pVal.Before_Action == false)
            {
                SAPbouiCOM.Form oPNForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                if(oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType",0) == "C")
                {
                    SAPbouiCOM.Item oNewItem;
                    oNewItem = oPNForm.Items.Item("RplPN");
                    oNewItem.Visible = true;
                }
                else
                {
                    SAPbouiCOM.Item oNewItem;
                    oNewItem = oPNForm.Items.Item("RplPN");
                    oNewItem.Visible = false;
                }
            }
            if (pVal.FormType == 134 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.Before_Action == true)
            {
                SAPbouiCOM.Form oPNForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                if (pVal.ItemUID == "RplPN")
                {
                    if(oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0) != "C")
                    {
                        string MsgTXT = "Esta função está disponível apenas para clientes.";
                        Application.SBO_Application.StatusBar.SetText(MsgTXT);
                        Application.SBO_Application.MessageBox(MsgTXT);
                        return;
                    }

                    //Replicar PN
                    //Application.SBO_Application.MessageBox("Replicar CardCode: " + oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0));
                    //Application.SBO_Application.MessageBox("Replicar Address: " + GetSAddr(oPNForm));
                    string PN_Addr = GetSAddr(oPNForm);
                    if(PN_Addr == "")
                    {
                        string MsgTXT = "Não encontrado endereço sem fornecedor cadastrado para replicação.";
                        Application.SBO_Application.StatusBar.SetText(MsgTXT);
                        Application.SBO_Application.MessageBox(MsgTXT);
                        return;
                    }
                    ReplicaPN(oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0), PN_Addr);
                }
            }
        }

        static void ReplicaPN(string CardCode, string Address)
        {
            SAPbobsCOM.BusinessPartners oCRD = (SAPbobsCOM.BusinessPartners)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
            SAPbobsCOM.BusinessPartners NoCRD = (SAPbobsCOM.BusinessPartners)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
            oCRD.GetByKey(CardCode);

            NoCRD.CardName = oCRD.CardName;
            NoCRD.CardForeignName = oCRD.CardForeignName;
            NoCRD.EmailAddress = oCRD.EmailAddress;
            NoCRD.Phone1 = oCRD.Phone1;
            NoCRD.Phone2 = oCRD.Phone2;
            NoCRD.MainUsage = oCRD.MainUsage;
            NoCRD.Territory = oCRD.Territory;

            NoCRD.CardType = SAPbobsCOM.BoCardTypes.cSupplier;
            NoCRD.GroupCode = 101;
            NoCRD.Series = 76;

            for(int i = 0; i < oCRD.Addresses.Count; i++)
            {
                if(oCRD.Addresses.AddressType == SAPbobsCOM.BoAddressType.bo_BillTo)
                    NoCRD.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo;
                else
                    NoCRD.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;

                NoCRD.Addresses.AddressName = oCRD.Addresses.AddressName;
                NoCRD.Addresses.Block = oCRD.Addresses.Block;
                NoCRD.Addresses.City = oCRD.Addresses.City;
                NoCRD.Addresses.County = oCRD.Addresses.County;
                NoCRD.Addresses.Country = oCRD.Addresses.Country;
                NoCRD.Addresses.State = oCRD.Addresses.State;
                NoCRD.Addresses.Street = oCRD.Addresses.Street;
                NoCRD.Addresses.StreetNo = oCRD.Addresses.StreetNo;
                NoCRD.Addresses.Add();
            }

            for (int i = 0; i < oCRD.ContactEmployees.Count; i++)
            {
                NoCRD.ContactEmployees.FirstName = oCRD.ContactEmployees.FirstName;
                /*NoCRD.ContactEmployees. = oCRD.ContactEmployees.FirstName;
                NoCRD.ContactEmployees.FirstName = oCRD.ContactEmployees.FirstName;
                NoCRD.ContactEmployees.FirstName = oCRD.ContactEmployees.FirstName;
                NoCRD.ContactEmployees.FirstName = oCRD.ContactEmployees.FirstName;*/
                NoCRD.ContactEmployees.Add();
            }

            NoCRD.Add();



            Application.SBO_Application.StatusBar.SetText("Criado PN: " + oCompany.GetNewObjectKey(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);


        }

        static string GetSAddr(SAPbouiCOM.Form oPNForm)
        {
            string retVal = "";
            for(int i = 0; i < oPNForm.DataSources.DBDataSources.Item("CRD1").Size; i++)
            {
                if(oPNForm.DataSources.DBDataSources.Item("CRD1").GetValue("AdresType", i) == "S" && oPNForm.DataSources.DBDataSources.Item("CRD1").GetValue("U_SD_CardCodeS", i) == "")
                {
                    retVal = oPNForm.DataSources.DBDataSources.Item("CRD1").GetValue("Address", i);
                }
            }
            return retVal;
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }
    }
}

using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;

namespace B1SSyngentaAddOn
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
                Application.SBO_Application.FormDataEvent += SBO_Application_FormDataEvent;
                
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        
        private static void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                /* Etapas de Apovaçãoo*/
                if (pVal.FormTypeEx == "50101" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && pVal.BeforeAction == false)
                {
                    SAPbouiCOM.Form oOWSTForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.Item oNewItem;
                    oNewItem = oOWSTForm.Items.Item("lblDpto");
                    oNewItem.Visible = true;

                    oNewItem = oOWSTForm.Items.Item("cmbDpto");
                    oNewItem.Visible = true;


                }
            }
            catch (Exception a)
            {
                Application.SBO_Application.MessageBox("Erro: " + a.Message);
            }

            try
            {
                /* Modelos de Autorização*/
                if (pVal.FormTypeEx == "50102" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && pVal.BeforeAction == false)
                {
                    SAPbouiCOM.Form oOWTMForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                    SAPbouiCOM.Item oNewItem;
                    SAPbouiCOM.Item oNewItem2;

                    oNewItem = oOWTMForm.Items.Item("chkNecJust");
                    oNewItem.Visible = true;

                    oNewItem2 = oOWTMForm.Items.Item("appHome");
                    oNewItem2.Visible = true;

                }
            }
            catch (Exception a)
            {
                Application.SBO_Application.MessageBox("Erro: " + a.Message);
            }


            try
            {

            
            if (pVal.FormTypeEx == "134" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && pVal.BeforeAction == false)
            {
                SAPbouiCOM.Form oPNForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                if (oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0) == "C")
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
                    if (oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("ConnBP", 0) != "")
                    {
                        string MsgTXT = "Este fornecedor possui um cliente conectado, as alterações devem ser realizadas no cadastro do cliente.";
                        Application.SBO_Application.StatusBar.SetText(MsgTXT, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        Application.SBO_Application.MessageBox(MsgTXT);
                        return;
                    }
                }
            }

                if (pVal.FormTypeEx == "134" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE && pVal.BeforeAction == true)
                {

                    SAPbouiCOM.Form oPNForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                    SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    recordset.DoQuery($@"SELECT ""CardType"" FROM OCRD WHERE ""CardCode"" = '{oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0)}'");

                    if (oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0) == "S" && oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("ConnBP", 0) != "")
                    {
                        string MsgTXT = "Este fornecedor possui um cliente conectado, as alterações devem ser realizadas no cadastro do cliente.";
                        Application.SBO_Application.StatusBar.SetText(MsgTXT);
                        Application.SBO_Application.MessageBox(MsgTXT);
                        BubbleEvent = false;
                        return;
                    }
                    if (oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0) == "C" && oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("ConnBP", 0) != "")
                    {
                        string MsgTXT = "Será atualizado o Fornecedor Conectado: " + oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("ConnBP", 0);
                        Application.SBO_Application.StatusBar.SetText(MsgTXT, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        //Application.SBO_Application.MessageBox(MsgTXT);
                        return;
                    }

                    if (oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0) == "L" && recordset.Fields.Item(0).Value.ToString() != "L")
                    {
                        string MsgTXT = "Não é mais possível indicar este cadastro como Cliente Potencial.";
                        Application.SBO_Application.StatusBar.SetText(MsgTXT);
                        Application.SBO_Application.MessageBox(MsgTXT);
                        BubbleEvent = false;
                        return;
                    }

                    if (oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0) != "L" && recordset.Fields.Item(0).Value.ToString() == "L")
                    {
                        if (!Application.SBO_Application.MessageBox("A ação de alterar o tipo de Cliente Potencial é irreverssível.\nDeseja continuar?", 1, "Sim", "Não").Equals(1))
                        {
                            BubbleEvent = false;
                            return;
                        }

                    }

                }
            }
            catch (Exception a)
            {

                Application.SBO_Application.MessageBox("Erro: " + a.Message);
            }


        }

        public static SAPbobsCOM.Company oCompany;

        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                /*Modelos de Aprovadores */
                if (pVal.FormType == 50101 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                {
                    SAPbouiCOM.Form oOWSTForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                    SAPbouiCOM.StaticText obutton;
                    SAPbouiCOM.Item oitem;
                    SAPbouiCOM.Item oNewItem;
                    SAPbouiCOM.Item oNewCmb;
                    oNewItem = oOWSTForm.Items.Add("lblDpto", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oNewCmb = oOWSTForm.Items.Add("cmbDpto", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    SAPbouiCOM.ComboBox cmbDpto = ((SAPbouiCOM.ComboBox)(oOWSTForm.Items.Item("cmbDpto").Specific));
                    cmbDpto.DataBind.SetBound(true, "OWST", "U_B1S_EXT_Depart");
                    cmbDpto.Item.DisplayDesc = true;
                    SAPbobsCOM.Recordset oRs = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRs.DoQuery(@"SELECT * FROM ""@B1S_EXT_DEPARTMENT""");

                    while (cmbDpto.ValidValues.Count > 0) { cmbDpto.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index); }
                    while (!oRs.EoF) { cmbDpto.ValidValues.Add(oRs.Fields.Item(0).Value.ToString(), oRs.Fields.Item(1).Value.ToString()); oRs.MoveNext(); }

                    obutton = (SAPbouiCOM.StaticText)oNewItem.Specific;
                    obutton.Caption = "Departamento";

                    oitem = oOWSTForm.Items.Item("1760000107"); // UI element in the system form to use for positional reference
                    oNewItem.Top = oitem.Top + 16;
                    oNewItem.Height = oitem.Height;
                    oNewItem.Width = oitem.Width;
                    oNewItem.Left = oitem.Left;
                    oNewItem.LinkTo = "cmbDpto";

                    oNewCmb.Top = oitem.Top + 16;
                    oNewCmb.Height = oitem.Height;
                    oNewCmb.Width = oitem.Width + 3;
                    oNewCmb.Left = oitem.Left + oitem.Width + 2;

                    oNewItem.Visible = true;
                    oNewCmb.Visible = true;
                }
            }
            catch (Exception a)
            {
                Application.SBO_Application.MessageBox("Erro: " + a.Message);
            }

            try
            {
                /*Modelos de Autorizção */
                if (pVal.FormType == 50102 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                {
                    SAPbouiCOM.Form oOWTMForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                    SAPbouiCOM.CheckBox obutton;
                    SAPbouiCOM.Item oitem;
                    SAPbouiCOM.Item oNewItem;
                    oNewItem = oOWTMForm.Items.Add("chkNecJust", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    SAPbouiCOM.CheckBox chk = ((SAPbouiCOM.CheckBox)(oOWTMForm.Items.Item("chkNecJust").Specific));
                    chk.DataBind.SetBound(true, "OWTM", "U_B1S_EXT_Justif");
                    obutton = (SAPbouiCOM.CheckBox)oNewItem.Specific;
                    obutton.Caption = "Necessita Jusitificativa?";

                    oitem = oOWTMForm.Items.Item("13"); // UI element in the system form to use for positional reference
                    oNewItem.Top = oitem.Top + 16;
                    oNewItem.Height = oitem.Height;
                    oNewItem.Width = oitem.Width;
                    oNewItem.Left = oitem.Left;

                    oNewItem.Visible = true;


                    SAPbouiCOM.CheckBox obutton2;
                    SAPbouiCOM.Item oitem2;
                    SAPbouiCOM.Item oNewItem2;
                    oNewItem2 = oOWTMForm.Items.Add("appHome", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    SAPbouiCOM.CheckBox chk2 = ((SAPbouiCOM.CheckBox)(oOWTMForm.Items.Item("appHome").Specific));
                    chk2.DataBind.SetBound(true, "OWTM", "U_B1S_EXT_HomeApproval");
                    obutton = (SAPbouiCOM.CheckBox)oNewItem.Specific;
                    obutton.Caption = "Aprova pela HOME?";

                    oitem2 = oOWTMForm.Items.Item("chkNecJust"); // UI element in the system form to use for positional reference
                    oNewItem2.Top = oitem2.Top;
                    oNewItem2.Height = oitem.Height;
                    oNewItem2.Width = oitem.Width;
                    oNewItem2.Left = oitem.Left + 100;

                    oNewItem2.Visible = true;
                }

            }
            catch (Exception b)
            {

                Application.SBO_Application.MessageBox("Erro: " + b.Message);
            }

            try
            {

            
            /*PN*/
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

                oNewItem.Visible = false;
            }

            if (pVal.FormType == 134 && pVal.Before_Action == true && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN || pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST) && pVal.ItemUID.Equals("1004"))
            {

                BubbleEvent = false;
                return;
                SAPbouiCOM.Form oPNForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                recordset.DoQuery(@"SELECT ""ConnBP"" FROM OCRD WHERE ""CardCode"" = '" + (oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0)) + "'");
                SAPbouiCOM.EditText editText = ((SAPbouiCOM.EditText)(oPNForm.Items.Item(pVal.ItemUID).Specific));
                if (editText.Value.ToString() != recordset.Fields.Item(0).Value.ToString())
                {
                    Application.SBO_Application.MessageBox("Não é possível alterar o campo de PN Conectado após sua inserção");
                    return;
                }

            }

            if (pVal.FormType == 134 && pVal.Before_Action == true && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) && pVal.ItemUID.Equals("40"))
            {

                
                SAPbouiCOM.Form oPNForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                SAPbouiCOM.ComboBox comboBox = ((SAPbouiCOM.ComboBox)oPNForm.Items.Item("40").Specific);

                if (oPNForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE && oPNForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    recordset.DoQuery($@"SELECT ""CardType"" FROM OCRD WHERE ""CardCode"" = '{oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0)}'");
                    BubbleEvent = true;

                    if (oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0) == "C")
                    {
                        string MsgTXT = "Não é mais possível alterar o tipo do cliente.";
                        Application.SBO_Application.StatusBar.SetText(MsgTXT);
                        Application.SBO_Application.MessageBox(MsgTXT);
                        BubbleEvent = false;
                        return;
                    }
                    string a = recordset.Fields.Item(0).Value.ToString();
                    if (comboBox.Value.ToString().Trim() == "L" && recordset.Fields.Item(0).Value.ToString() == "L")
                    {
                        if (!Application.SBO_Application.MessageBox("A ação de alterar o tipo de Cliente Potencial é irreverssível.\nDeseja continuar?", 1, "Sim", "Não").Equals(1))
                        {
                            BubbleEvent = false;
                            return;
                        }

                        return;

                    }

                }
                

            }

                if (pVal.FormType == 134 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.Before_Action == true)
                {
                    SAPbouiCOM.Form oPNForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                    if (pVal.ItemUID == "RplPN")
                    {

                        if (oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0) != "C")
                        {
                            string MsgTXT = "Esta função está disponível apenas para clientes.";
                            Application.SBO_Application.StatusBar.SetText(MsgTXT);
                            Application.SBO_Application.MessageBox(MsgTXT);
                            return;
                        }

                        if (oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("ConnBP", 0) != "")
                        {
                            string MsgTXT = "Este cliente já possui um fornecedor conectado.";
                            Application.SBO_Application.StatusBar.SetText(MsgTXT);
                            Application.SBO_Application.MessageBox(MsgTXT);
                            return;
                        }
                        if (oPNForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            string MsgTXT = "";
                            switch (oPNForm.Mode)
                            {
                                case SAPbouiCOM.BoFormMode.fm_ADD_MODE:
                                    MsgTXT = "Não é possível replicar um cadastro antes de adiciona-lo";
                                    break;

                                case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE:
                                    MsgTXT = "Finalize as edições no cadastro antes de replica-lo";
                                    break;

                                case SAPbouiCOM.BoFormMode.fm_FIND_MODE:
                                    MsgTXT = "Esta função não é permitida";
                                    break;
                            }
                            Application.SBO_Application.StatusBar.SetText(MsgTXT);
                            Application.SBO_Application.MessageBox(MsgTXT);
                            return;
                        }
                        //Replicar PN
                        //Application.SBO_Application.MessageBox("Replicar CardCode: " + oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0));
                        //Application.SBO_Application.MessageBox("Replicar Address: " + GetSAddr(oPNForm));

                        ReplicaPN(oPNForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0));
                    }
                }
            }
            catch (Exception c)
            {

                Application.SBO_Application.MessageBox("Erro: " + c.Message);
            }

        }

        static void ReplicaPN(string CardCode)
        {
            int RetVal;
            string NCardCode = "";
            Application.SBO_Application.StatusBar.SetText("Realizando a criação do Fornecedor Conectado...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            oCompany.StartTransaction();
            try
            {
                SAPbobsCOM.BusinessPartners oCRD = (SAPbobsCOM.BusinessPartners)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                SAPbobsCOM.BusinessPartners NoCRD = (SAPbobsCOM.BusinessPartners)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);



                oCRD.GetByKey(CardCode);

                //NoCRD.CardCode = oCRD.CardCode;
                NoCRD.CardName = oCRD.CardName;
                //NoCRD.Address = oCRD.Address;
                //NoCRD.ZipCode = oCRD.ZipCode;
                //NoCRD.MailAddress = oCRD.MailAddress;
                //NoCRD.MailZipCode = oCRD.MailZipCode;
                NoCRD.Phone1 = oCRD.Phone1;
                NoCRD.Phone2 = oCRD.Phone2;
                NoCRD.Fax = oCRD.Fax;
                //NoCRD.ContactPerson = oCRD.ContactPerson;
                NoCRD.Notes = oCRD.Notes;
                //NoCRD.PayTermsGrpCode = oCRD.PayTermsGrpCode;
                NoCRD.CreditLimit = oCRD.CreditLimit;
                NoCRD.MaxCommitment = oCRD.MaxCommitment;
                //NoCRD.DiscountPercent = oCRD.DiscountPercent;
                //NoCRD.VatLiable = oCRD.VatLiable;
                NoCRD.FederalTaxID = oCRD.FederalTaxID;
                NoCRD.DeductibleAtSource = oCRD.DeductibleAtSource;
                NoCRD.DeductionPercent = oCRD.DeductionPercent;
                NoCRD.DeductionValidUntil = oCRD.DeductionValidUntil;
                NoCRD.PriceListNum = oCRD.PriceListNum;
                NoCRD.IntrestRatePercent = oCRD.IntrestRatePercent;
                NoCRD.CommissionPercent = oCRD.CommissionPercent;
                NoCRD.CommissionGroupCode = oCRD.CommissionGroupCode;
                NoCRD.FreeText = oCRD.FreeText;
                NoCRD.SalesPersonCode = oCRD.SalesPersonCode;
                NoCRD.Currency = oCRD.Currency;
                NoCRD.RateDiffAccount = oCRD.RateDiffAccount;
                NoCRD.Cellular = oCRD.Cellular;
                NoCRD.AvarageLate = oCRD.AvarageLate;
                //NoCRD.City = oCRD.City;
                //NoCRD.County = oCRD.County;
                //NoCRD.Country = oCRD.Country;
                //NoCRD.MailCity = oCRD.MailCity;
                //NoCRD.MailCounty = oCRD.MailCounty;
                //NoCRD.MailCountry = oCRD.MailCountry;
                NoCRD.EmailAddress = oCRD.EmailAddress;
                //NoCRD.Picture = oCRD.Picture;
                NoCRD.DefaultAccount = oCRD.DefaultAccount;
                NoCRD.DefaultBranch = oCRD.DefaultBranch;
                NoCRD.DefaultBankCode = oCRD.DefaultBankCode;
                NoCRD.AdditionalID = oCRD.AdditionalID;
                NoCRD.Pager = oCRD.Pager;
                //NoCRD.FatherCard = oCRD.FatherCard;
                NoCRD.CardForeignName = oCRD.CardForeignName;
                //NoCRD.FatherType = oCRD.FatherType;
                //NoCRD.DeductionOffice = oCRD.DeductionOffice;
                //NoCRD.ExportCode = oCRD.ExportCode;
                //NoCRD.MinIntrest = oCRD.MinIntrest;
                //NoCRD.VatGroup = oCRD.VatGroup;
                if (oCRD.ShippingType != 0)
                    NoCRD.ShippingType = oCRD.ShippingType;

                NoCRD.Password = oCRD.Password;
                NoCRD.Indicator = oCRD.Indicator;
                NoCRD.IBAN = oCRD.IBAN;
                //NoCRD.CreditCardCode = oCRD.CreditCardCode;
                //NoCRD.CreditCardNum = oCRD.CreditCardNum;
                //NoCRD.CreditCardExpiration = oCRD.CreditCardExpiration;
                //NoCRD.DebitorAccount = oCRD.DebitorAccount;
                NoCRD.Valid = oCRD.Valid;
                NoCRD.ValidFrom = oCRD.ValidFrom;
                NoCRD.ValidTo = oCRD.ValidTo;
                NoCRD.ValidRemarks = oCRD.ValidRemarks;
                NoCRD.Frozen = oCRD.Frozen;
                NoCRD.FrozenFrom = oCRD.FrozenFrom;
                NoCRD.FrozenTo = oCRD.FrozenTo;
                NoCRD.FrozenRemarks = oCRD.FrozenRemarks;
                //NoCRD.Block = oCRD.Block;
                //NoCRD.BillToState = oCRD.BillToState;
                NoCRD.ExemptNum = oCRD.ExemptNum;
                if(oCRD.Territory != 0)
                    NoCRD.Territory = oCRD.Territory;
                NoCRD.Website = oCRD.Website;
                //NoCRD.Priority = oCRD.Priority;
                //NoCRD.FormCode1099 = oCRD.FormCode1099;
                //NoCRD.Box1099 = oCRD.Box1099;
                //NoCRD.PeymentMethodCode = oCRD.PeymentMethodCode;
                //NoCRD.BackOrder = oCRD.BackOrder;
                //NoCRD.PartialDelivery = oCRD.PartialDelivery;
                //NoCRD.BlockDunning = oCRD.BlockDunning;
                NoCRD.BankCountry = oCRD.BankCountry;
                NoCRD.HouseBank = oCRD.HouseBank;
                NoCRD.HouseBankCountry = oCRD.HouseBankCountry;
                NoCRD.HouseBankAccount = oCRD.HouseBankAccount;
                //NoCRD.ShipToDefault = oCRD.ShipToDefault;
                //NoCRD.CollectionAuthorization = oCRD.CollectionAuthorization;
                //NoCRD.DME = oCRD.DME;
                //NoCRD.InstructionKey = oCRD.InstructionKey;
                //NoCRD.SinglePayment = oCRD.SinglePayment;
                //NoCRD.ISRBillerID = oCRD.ISRBillerID;
                //NoCRD.PaymentBlock = oCRD.PaymentBlock;
                //NoCRD.ReferenceDetails = oCRD.ReferenceDetails;
                //NoCRD.HouseBankBranch = oCRD.HouseBankBranch;
                //NoCRD.OwnerIDNumber = oCRD.OwnerIDNumber;
                //NoCRD.PaymentBlockDescription = oCRD.PaymentBlockDescription;
                //NoCRD.TaxExemptionLetterNum = oCRD.TaxExemptionLetterNum;
                //NoCRD.MaxAmountOfExemption = oCRD.MaxAmountOfExemption;
                //NoCRD.ExemptionValidityDateFrom = oCRD.ExemptionValidityDateFrom;
                //NoCRD.ExemptionValidityDateTo = oCRD.ExemptionValidityDateTo;
                //NoCRD.LinkedBusinessPartner = oCRD.LinkedBusinessPartner;
                //NoCRD.LastMultiReconciliationNum = oCRD.LastMultiReconciliationNum;
                //NoCRD.Equalization = oCRD.Equalization;
                //NoCRD.SubjectToWithholdingTax = oCRD.SubjectToWithholdingTax;
                //NoCRD.CertificateNumber = oCRD.CertificateNumber;
                //NoCRD.ExpirationDate = oCRD.ExpirationDate;
                //NoCRD.NationalInsuranceNum = oCRD.NationalInsuranceNum;
                //NoCRD.AccrualCriteria = oCRD.AccrualCriteria;
                //NoCRD.WTCode = oCRD.WTCode;
                //NoCRD.DeferredTax = oCRD.DeferredTax;
                //NoCRD.BillToBuildingFloorRoom = oCRD.BillToBuildingFloorRoom;
                //NoCRD.DownPaymentClearAct = oCRD.DownPaymentClearAct;
                //NoCRD.ChannelBP = oCRD.ChannelBP;
                /*NoCRD.DefaultTechnician = oCRD.DefaultTechnician;
                NoCRD.BilltoDefault = oCRD.BilltoDefault;
                NoCRD.CustomerBillofExchangDisc = oCRD.CustomerBillofExchangDisc;
                NoCRD.ShipToBuildingFloorRoom = oCRD.ShipToBuildingFloorRoom;
                NoCRD.CustomerBillofExchangPres = oCRD.CustomerBillofExchangPres;
                NoCRD.ProjectCode = oCRD.ProjectCode;
                NoCRD.VatGroupLatinAmerica = oCRD.VatGroupLatinAmerica;
                NoCRD.DunningTerm = oCRD.DunningTerm;
                
                NoCRD.OtherReceivablePayable = oCRD.OtherReceivablePayable;
                NoCRD.ClosingDateProcedureNumber = oCRD.ClosingDateProcedureNumber;
                NoCRD.Profession = oCRD.Profession;
                NoCRD.BillofExchangeonCollection = oCRD.BillofExchangeonCollection;
                */
                NoCRD.CompanyPrivate = oCRD.CompanyPrivate;
                NoCRD.LanguageCode = oCRD.LanguageCode;
                /*
                NoCRD.UnpaidBillofExchange = oCRD.UnpaidBillofExchange;
                NoCRD.WithholdingTaxDeductionGroup = oCRD.WithholdingTaxDeductionGroup;
                NoCRD.BankChargesAllocationCode = oCRD.BankChargesAllocationCode;
                NoCRD.TaxRoundingRule = oCRD.TaxRoundingRule;
                NoCRD.CompanyRegistrationNumber = oCRD.CompanyRegistrationNumber;
                NoCRD.VerificationNumber = oCRD.VerificationNumber;
                NoCRD.OperationCode347 = oCRD.OperationCode347;
                NoCRD.InsuranceOperation347 = oCRD.InsuranceOperation347;
                NoCRD.DiscountBaseObject = oCRD.DiscountBaseObject;
                NoCRD.DiscountRelations = oCRD.DiscountRelations;
                NoCRD.TypeReport = oCRD.TypeReport;
                NoCRD.ThresholdOverlook = oCRD.ThresholdOverlook;
                NoCRD.SurchargeOverlook = oCRD.SurchargeOverlook;
                NoCRD.DownPaymentInterimAccount = oCRD.DownPaymentInterimAccount;
                NoCRD.HierarchicalDeduction = oCRD.HierarchicalDeduction;
                NoCRD.ShaamGroup = oCRD.ShaamGroup;
                NoCRD.WithholdingTaxCertified = oCRD.WithholdingTaxCertified;
                NoCRD.BookkeepingCertified = oCRD.BookkeepingCertified;
                NoCRD.PlanningGroup = oCRD.PlanningGroup;
                NoCRD.Affiliate = oCRD.Affiliate;
                NoCRD.Industry = oCRD.Industry;
                NoCRD.VatIDNum = oCRD.VatIDNum;
                NoCRD.DatevAccount = oCRD.DatevAccount;
                NoCRD.DatevFirstDataEntry = oCRD.DatevFirstDataEntry;
                NoCRD.GTSRegNo = oCRD.GTSRegNo;
                NoCRD.GTSBankAccountNo = oCRD.GTSBankAccountNo;
                NoCRD.GTSBillingAddrTel = oCRD.GTSBillingAddrTel;
                NoCRD.ETaxWebSite = oCRD.ETaxWebSite;
                NoCRD.AutomaticPosting = oCRD.AutomaticPosting;
                NoCRD.InterestAccount = oCRD.InterestAccount;
                NoCRD.FeeAccount = oCRD.FeeAccount;
                NoCRD.CampaignNumber = oCRD.CampaignNumber;
                NoCRD.VATRegistrationNumber = oCRD.VATRegistrationNumber;
                NoCRD.RepresentativeName = oCRD.RepresentativeName;
                NoCRD.IndustryType = oCRD.IndustryType;
                NoCRD.BusinessType = oCRD.BusinessType;
                
                NoCRD.DefaultBlanketAgreementNumber = oCRD.DefaultBlanketAgreementNumber;
                NoCRD.EffectiveDiscount = oCRD.EffectiveDiscount;
                NoCRD.NoDiscounts = oCRD.NoDiscounts;
                NoCRD.GlobalLocationNumber = oCRD.GlobalLocationNumber;
                NoCRD.EDISenderID = oCRD.EDISenderID;
                NoCRD.EDIRecipientID = oCRD.EDIRecipientID;
                NoCRD.ResidenNumber = oCRD.ResidenNumber;
                NoCRD.UnifiedFederalTaxID = oCRD.UnifiedFederalTaxID;
                NoCRD.RelationshipDateFrom = oCRD.RelationshipDateFrom;
                NoCRD.RelationshipDateTill = oCRD.RelationshipDateTill;
                NoCRD.RelationshipCode = oCRD.RelationshipCode;*/
                //NoCRD.AttachmentEntry = oCRD.AttachmentEntry;
                //NoCRD.TypeOfOperation = oCRD.TypeOfOperation;
                NoCRD.OwnerCode = oCRD.OwnerCode;
                NoCRD.AliasName = oCRD.AliasName;
                //NoCRD.EndorsableChecksFromBP = oCRD.EndorsableChecksFromBP;
                //NoCRD.AcceptsEndorsedChecks = oCRD.AcceptsEndorsedChecks;
                //NoCRD.BlockSendingMarketingContent = oCRD.BlockSendingMarketingContent;
                //NoCRD.AgentCode = oCRD.AgentCode;
                /*NoCRD.EDocGenerationType = oCRD.EDocGenerationType;
                NoCRD.EDocStreet = oCRD.EDocStreet;
                NoCRD.EDocStreetNumber = oCRD.EDocStreetNumber;
                NoCRD.EDocBuildingNumber = oCRD.EDocBuildingNumber;
                NoCRD.EDocZipCode = oCRD.EDocZipCode;
                NoCRD.EDocCity = oCRD.EDocCity;
                NoCRD.EDocCountry = oCRD.EDocCountry;
                NoCRD.EDocDistrict = oCRD.EDocDistrict;
                NoCRD.EDocRepresentativeFirstName = oCRD.EDocRepresentativeFirstName;
                NoCRD.EDocRepresentativeSurname = oCRD.EDocRepresentativeSurname;
                NoCRD.EDocRepresentativeCompany = oCRD.EDocRepresentativeCompany;
                NoCRD.EDocRepresentativeFiscalCode = oCRD.EDocRepresentativeFiscalCode;
                NoCRD.EDocRepresentativeAdditionalId = oCRD.EDocRepresentativeAdditionalId;
                NoCRD.EDocPECAddress = oCRD.EDocPECAddress;
                NoCRD.IPACodeForPA = oCRD.IPACodeForPA;
                NoCRD.ExemptionMaxAmountValidationType = oCRD.ExemptionMaxAmountValidationType;
                NoCRD.ECommerceMerchantID = oCRD.ECommerceMerchantID;
                NoCRD.UseBillToAddrToDetermineTax = oCRD.UseBillToAddrToDetermineTax;
                NoCRD.PriceMode = oCRD.PriceMode;
                NoCRD.EffectivePrice = oCRD.EffectivePrice;
                NoCRD.UseShippedGoodsAccount = oCRD.UseShippedGoodsAccount;
                NoCRD.DefaultTransporterEntry = oCRD.DefaultTransporterEntry;
                NoCRD.DefaultTransporterLineNumber = oCRD.DefaultTransporterLineNumber;
                NoCRD.FCERelevant = oCRD.FCERelevant;
                NoCRD.FCEValidateBaseDelivery = oCRD.FCEValidateBaseDelivery;
                NoCRD.EffectivePriceConsidersPriceBeforeDiscount = oCRD.EffectivePriceConsidersPriceBeforeDiscount;
                NoCRD.MainUsage = oCRD.MainUsage;
                NoCRD.EBooksVATExemptionCause = oCRD.EBooksVATExemptionCause;
                NoCRD.LegalText = oCRD.LegalText;
                NoCRD.ExchangeRateForIncomingPayment = oCRD.ExchangeRateForIncomingPayment;
                NoCRD.ExchangeRateForOutgoingPayment = oCRD.ExchangeRateForOutgoingPayment;
                NoCRD.CertificateDetails = oCRD.CertificateDetails;
                NoCRD.DefaultCurrency = oCRD.DefaultCurrency;
                NoCRD.EORINumber = oCRD.EORINumber;
                NoCRD.FCEAsPaymentMeans = oCRD.FCEAsPaymentMeans;
                NoCRD.DeferCommitmentLimitOnDueDate = oCRD.DeferCommitmentLimitOnDueDate;
                NoCRD.DeferCommitmentLimitOnDueDateMonths = oCRD.DeferCommitmentLimitOnDueDateMonths;
                NoCRD.DeferCommitmentLimitOnDueDateDays = oCRD.DeferCommitmentLimitOnDueDateDays;*/


                NoCRD.CardType = SAPbobsCOM.BoCardTypes.cSupplier;
                NoCRD.GroupCode = 101;
                //TO-DO: criar parametro
                NoCRD.Series = 78;

                recordset.DoQuery(@"SELECT * FROM OPYM T0 WHERE T0.""Type"" = 'O'");

                while (!recordset.EoF)
                {
                    NoCRD.BPPaymentMethods.PaymentMethodCode = recordset.Fields.Item("PayMethCod").Value.ToString();
                    NoCRD.BPPaymentMethods.Add();
                    recordset.MoveNext();
                }

                for (int i = 1; i <= 64; i++)
                {
                    NoCRD.Properties[i] = oCRD.Properties[i];
                }
                

                for (int i = 0; i < oCRD.UserFields.Fields.Count; i++)
                {
                    if (oCRD.UserFields.Fields.Item(i).Name.Equals("U_AGRT_UUID_CardCode"))
                        NoCRD.UserFields.Fields.Item(i).Value = "";
                    else if (oCRD.UserFields.Fields.Item(i).Name.Equals("U_AGRT_UUID_PDR"))
                        NoCRD.UserFields.Fields.Item(i).Value = "";
                    else
                        NoCRD.UserFields.Fields.Item(i).Value = oCRD.UserFields.Fields.Item(i).Value;

                }

                for (int i = 0; i < oCRD.Addresses.Count; i++)
                {
                    

                    oCRD.Addresses.SetCurrentLine(i);
                    if(i > 0)
                        NoCRD.Addresses.Add();
                    NoCRD.Addresses.SetCurrentLine(i);

                    if (oCRD.Addresses.AddressType == SAPbobsCOM.BoAddressType.bo_BillTo)
                        NoCRD.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo;
                    else
                        NoCRD.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                    NoCRD.Addresses.AddressName = oCRD.Addresses.AddressName;
                    NoCRD.Addresses.Street = oCRD.Addresses.Street;
                    NoCRD.Addresses.Block = oCRD.Addresses.Block;
                    NoCRD.Addresses.ZipCode = oCRD.Addresses.ZipCode;
                    NoCRD.Addresses.City = oCRD.Addresses.City;
                    NoCRD.Addresses.County = oCRD.Addresses.County;
                    NoCRD.Addresses.Country = oCRD.Addresses.Country;
                    NoCRD.Addresses.State = oCRD.Addresses.State;

                    NoCRD.Addresses.FederalTaxID = oCRD.Addresses.FederalTaxID;
                    NoCRD.Addresses.TaxCode = oCRD.Addresses.TaxCode;
                    NoCRD.Addresses.BuildingFloorRoom = oCRD.Addresses.BuildingFloorRoom;
                    NoCRD.Addresses.AddressName2 = oCRD.Addresses.AddressName2;
                    NoCRD.Addresses.AddressName3 = oCRD.Addresses.AddressName3;
                    NoCRD.Addresses.TypeOfAddress = oCRD.Addresses.TypeOfAddress;
                    NoCRD.Addresses.StreetNo = oCRD.Addresses.StreetNo;
                    NoCRD.Addresses.GlobalLocationNumber = oCRD.Addresses.GlobalLocationNumber;
                    NoCRD.Addresses.Nationality = oCRD.Addresses.Nationality;
                    NoCRD.Addresses.TaxOffice = oCRD.Addresses.TaxOffice;
                    NoCRD.Addresses.GSTIN = oCRD.Addresses.GSTIN;
                    NoCRD.Addresses.GstType = oCRD.Addresses.GstType;
                    NoCRD.Addresses.MYFType = oCRD.Addresses.MYFType;
                    NoCRD.Addresses.TaasEnabled = oCRD.Addresses.TaasEnabled;
                    //NoCRD.Addresses.Add();

                    for (int iEnd = 0; iEnd < oCRD.Addresses.UserFields.Fields.Count; iEnd++)
                    {

                        if (oCRD.Addresses.UserFields.Fields.Item(iEnd).Name.Equals("U_AGRT_UUID_PPR"))
                            NoCRD.Addresses.UserFields.Fields.Item(iEnd).Value = "";
                        else if (oCRD.Addresses.UserFields.Fields.Item(iEnd).Name.Equals("U_AGRT_PropriedadeRural"))
                                NoCRD.Addresses.UserFields.Fields.Item(iEnd).Value = "N";
                        else
                            NoCRD.Addresses.UserFields.Fields.Item(iEnd).Value = oCRD.Addresses.UserFields.Fields.Item(iEnd).Value;
                        
                    }

                }

                for (int i = 0; i < oCRD.BPBankAccounts.Count; i++)
                {
                    oCRD.BPBankAccounts.SetCurrentLine(i);
                    NoCRD.BPBankAccounts.ABARoutingNumber = oCRD.BPBankAccounts.ABARoutingNumber;
                    NoCRD.BPBankAccounts.AccountName = oCRD.BPBankAccounts.AccountName;
                    NoCRD.BPBankAccounts.AccountNo = oCRD.BPBankAccounts.AccountNo;
                    NoCRD.BPBankAccounts.BankCode = oCRD.BPBankAccounts.BankCode;
                    NoCRD.BPBankAccounts.BICSwiftCode = oCRD.BPBankAccounts.BICSwiftCode;
                    NoCRD.BPBankAccounts.BIK = oCRD.BPBankAccounts.BIK;
                    NoCRD.BPBankAccounts.Block = oCRD.BPBankAccounts.Block;
                    NoCRD.BPBankAccounts.BPCode = oCRD.BPBankAccounts.BPCode;
                    NoCRD.BPBankAccounts.Branch = oCRD.BPBankAccounts.Branch;
                    NoCRD.BPBankAccounts.BuildingFloorRoom = oCRD.BPBankAccounts.BuildingFloorRoom;
                    NoCRD.BPBankAccounts.City = oCRD.BPBankAccounts.City;
                    NoCRD.BPBankAccounts.ControlKey = oCRD.BPBankAccounts.ControlKey;
                    NoCRD.BPBankAccounts.CorrespondentAccount = oCRD.BPBankAccounts.CorrespondentAccount;
                    NoCRD.BPBankAccounts.Country = oCRD.BPBankAccounts.Country;
                    NoCRD.BPBankAccounts.County = oCRD.BPBankAccounts.County;
                    NoCRD.BPBankAccounts.CustomerIdNumber = oCRD.BPBankAccounts.CustomerIdNumber;
                    NoCRD.BPBankAccounts.Fax = oCRD.BPBankAccounts.Fax;
                    NoCRD.BPBankAccounts.IBAN = oCRD.BPBankAccounts.IBAN;
                    //NoCRD.BPBankAccounts.InternalKey = oCRD.BPBankAccounts.InternalKey;
                    NoCRD.BPBankAccounts.ISRBillerID = oCRD.BPBankAccounts.ISRBillerID;
                    NoCRD.BPBankAccounts.ISRType = oCRD.BPBankAccounts.ISRType;
                    NoCRD.BPBankAccounts.MandateExpDate = oCRD.BPBankAccounts.MandateExpDate;
                    NoCRD.BPBankAccounts.MandateID = oCRD.BPBankAccounts.MandateID;
                    NoCRD.BPBankAccounts.Phone = oCRD.BPBankAccounts.Phone;
                    NoCRD.BPBankAccounts.SEPASeqType = oCRD.BPBankAccounts.SEPASeqType;
                    NoCRD.BPBankAccounts.State = oCRD.BPBankAccounts.State;
                    NoCRD.BPBankAccounts.Street = oCRD.BPBankAccounts.Street;
                    NoCRD.BPBankAccounts.ZipCode = oCRD.BPBankAccounts.ZipCode;
                }

                for (int i = 0; i < oCRD.ContactEmployees.Count; i++)
                {
                    oCRD.ContactEmployees.SetCurrentLine(i);
                    NoCRD.ContactEmployees.Position = oCRD.ContactEmployees.Position;
                    NoCRD.ContactEmployees.Address = oCRD.ContactEmployees.Address;
                    NoCRD.ContactEmployees.Phone1 = oCRD.ContactEmployees.Phone1;
                    NoCRD.ContactEmployees.MobilePhone = oCRD.ContactEmployees.MobilePhone;
                    NoCRD.ContactEmployees.Fax = oCRD.ContactEmployees.Fax;
                    NoCRD.ContactEmployees.E_Mail = oCRD.ContactEmployees.E_Mail;
                    NoCRD.ContactEmployees.Pager = oCRD.ContactEmployees.Pager;
                    NoCRD.ContactEmployees.Remarks1 = oCRD.ContactEmployees.Remarks1;
                    NoCRD.ContactEmployees.Remarks2 = oCRD.ContactEmployees.Remarks2;
                    NoCRD.ContactEmployees.Password = oCRD.ContactEmployees.Password;
                    NoCRD.ContactEmployees.Name = oCRD.ContactEmployees.Name;
                    NoCRD.ContactEmployees.PlaceOfBirth = oCRD.ContactEmployees.PlaceOfBirth;
                    NoCRD.ContactEmployees.DateOfBirth = oCRD.ContactEmployees.DateOfBirth;
                    NoCRD.ContactEmployees.Gender = oCRD.ContactEmployees.Gender;
                    NoCRD.ContactEmployees.Profession = oCRD.ContactEmployees.Profession;
                    NoCRD.ContactEmployees.Title = oCRD.ContactEmployees.Title;
                    NoCRD.ContactEmployees.CityOfBirth = oCRD.ContactEmployees.CityOfBirth;
                    NoCRD.ContactEmployees.Active = oCRD.ContactEmployees.Active;
                    NoCRD.ContactEmployees.FirstName = oCRD.ContactEmployees.FirstName;
                    NoCRD.ContactEmployees.MiddleName = oCRD.ContactEmployees.MiddleName;
                    NoCRD.ContactEmployees.LastName = oCRD.ContactEmployees.LastName;
                    NoCRD.ContactEmployees.EmailGroupCode = oCRD.ContactEmployees.EmailGroupCode;
                    NoCRD.ContactEmployees.BlockSendingMarketingContent = oCRD.ContactEmployees.BlockSendingMarketingContent;
                    NoCRD.ContactEmployees.ConnectedAddressType = oCRD.ContactEmployees.ConnectedAddressType;
                    NoCRD.ContactEmployees.ConnectedAddressName = oCRD.ContactEmployees.ConnectedAddressName;
                    //NoCRD.ContactEmployees.ForeignCountry = oCRD.ContactEmployees.ForeignCountry;
                    //for (int iContact = 0; iContact < oCRD.ContactEmployees.UserFields.Fields.Count; iContact++)
                    //{
                    //    if (!oCRD.ContactEmployees.UserFields.Fields.Item(iContact).Name.Contains("AGRT"))
                    //        NoCRD.ContactEmployees.UserFields.Fields.Item(iContact).Value = oCRD.ContactEmployees.UserFields.Fields.Item(iContact).Value;

                    //}

                    NoCRD.ContactEmployees.Add();

                    
                }

                //for (int i = 0; i < oCRD.Addresses.Count; i++)
                //{


                //    oCRD.Addresses.SetCurrentLine(i);
                //    if (i > 0)
                //        NoCRD.Addresses.Add();
                //    NoCRD.Addresses.SetCurrentLine(i);

                //    if (oCRD.Addresses.AddressType == SAPbobsCOM.BoAddressType.bo_BillTo)
                //        NoCRD.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo;
                //    else
                //        NoCRD.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                //    NoCRD.Addresses.AddressName = oCRD.Addresses.AddressName;
                //}

                
                NoCRD.FiscalTaxID.TaxId0 = oCRD.FiscalTaxID.TaxId0;
                NoCRD.FiscalTaxID.TaxId1 = oCRD.FiscalTaxID.TaxId1;
                NoCRD.FiscalTaxID.TaxId4 = oCRD.FiscalTaxID.TaxId4;
                NoCRD.FiscalTaxID.Address = "";

                NoCRD.FiscalTaxID.Add();

                //NoCRD.SaveToFile(@"C:\temp\bp.xml");

                RetVal = NoCRD.Add();

                if (RetVal != 0)
                {
                    if (oCompany.InTransaction)
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                    oCompany.GetLastError(out int ErrCode, out string ErrMsg);
                    throw new Exception("Erro criando PN: " + ErrMsg);
                }
                else
                {
                    NCardCode = oCompany.GetNewObjectKey();

                    oCRD.LinkedBusinessPartner = NCardCode;

                    RetVal = oCRD.Update();
                    if (RetVal != 0)
                    {
                        if (oCompany.InTransaction)
                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                        oCompany.GetLastError(out int ErrCode, out string ErrMsg);
                        throw new Exception("Erro atualizando cliente: " + ErrMsg + ", processo abortado.");
                    }

                    if (oCompany.InTransaction)
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                    Application.SBO_Application.ActivateMenuItem("1304");
                    Application.SBO_Application.StatusBar.SetText("Criado PN: " + NCardCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }


                

            }
            catch (Exception ex)
            {
                if (oCompany.InTransaction)
                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                Application.SBO_Application.StatusBar.SetText(ex.Message);

                //throw new Exception(ex.Message);
            }

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

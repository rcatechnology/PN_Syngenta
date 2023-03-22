using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace B1SSyngentaAddOn.UIForms.SystemForms
{
    [FormAttribute("392", "UIForms/SystemForms/frm392_JournalEntry.b1f")]
    class frm392_JournalEntry : SystemFormBase
    {
        private SAPbouiCOM.EditText edit_TransId;
        private SAPbouiCOM.Matrix mtx_Lines;
        private SAPbouiCOM.EditText edit_refdate;
        private SAPbouiCOM.EditText edit_duedate;
        private SAPbouiCOM.EditText edit_taxdate;
        private SAPbouiCOM.EditText edit_memo;
        private SAPbouiCOM.EditText edit_project;
        private SAPbouiCOM.EditText edit_refOne;
        private SAPbouiCOM.EditText edit_refTwo;
        private SAPbouiCOM.EditText edit_refThree;

        private SAPbouiCOM.ComboBox cmb_series;
        private SAPbouiCOM.ComboBox cmb_ecdType;
        private SAPbouiCOM.ComboBox cmb_indicator;
        private SAPbouiCOM.ComboBox cmb_transcode;
        private SAPbouiCOM.ComboBox cmb_matriz;

        private SAPbouiCOM.CheckBox chk_cambio;
        private SAPbouiCOM.CheckBox chk_estorno;
        private SAPbouiCOM.CheckBox chk_comp;

        private SAPbouiCOM.Button btn_main;
        public frm392_JournalEntry()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.edit_TransId = this.GetSpecificItem<SAPbouiCOM.EditText>("5");
            this.edit_refdate = this.GetSpecificItem<SAPbouiCOM.EditText>("6");
            this.edit_duedate = this.GetSpecificItem<SAPbouiCOM.EditText>("102");
            this.edit_taxdate = this.GetSpecificItem<SAPbouiCOM.EditText>("97");
            this.edit_memo = this.GetSpecificItem<SAPbouiCOM.EditText>("10");
            this.edit_project = this.GetSpecificItem<SAPbouiCOM.EditText>("26");
            this.edit_refOne = this.GetSpecificItem<SAPbouiCOM.EditText>("7");
            this.edit_refTwo = this.GetSpecificItem<SAPbouiCOM.EditText>("8");
            this.edit_refThree = this.GetSpecificItem<SAPbouiCOM.EditText>("540002023");
            this.cmb_ecdType = this.GetSpecificItem<SAPbouiCOM.ComboBox>("1980000004");
            this.cmb_transcode = this.GetSpecificItem<SAPbouiCOM.ComboBox>("9");
            this.cmb_series = this.GetSpecificItem<SAPbouiCOM.ComboBox>("137");
            this.cmb_indicator = this.GetSpecificItem<SAPbouiCOM.ComboBox>("93");
            this.cmb_matriz = this.GetSpecificItem<SAPbouiCOM.ComboBox>("1320002034");
            this.chk_cambio = this.GetSpecificItem<SAPbouiCOM.CheckBox>("82");
            this.chk_estorno = this.GetSpecificItem<SAPbouiCOM.CheckBox>("99");
            this.chk_comp = this.GetSpecificItem<SAPbouiCOM.CheckBox>("95");
            this.btn_main = this.GetSpecificItem<SAPbouiCOM.Button>("1");
            this.mtx_Lines = this.GetSpecificItem<SAPbouiCOM.Matrix>("76");

            this.OnCustomInitialize();
        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += this.PreLcm_LoadAfter;
            this.DataLoadAfter += this.LoadDataAfter;
        }

        private void LoadDataAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            string transId = edit_TransId.Value.ToString();
            string identificadorRh = GetSdrIntRhValue(transId);

            //caso seja nulo significa que o pre lcm e da integração
            if (String.IsNullOrWhiteSpace(identificadorRh) || identificadorRh == "0")
                ChangeFormState(true);
            else
                ChangeFormState(false);
        }

        private void OnCustomInitialize()
        {
            this.btn_main.ClickBefore += Btn_main_ClickBefore;
        }

        private void Btn_main_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string transId = edit_TransId.Value.ToString();
            string identificadorRh = GetSdrIntRhValue(transId);

            //caso seja nulo significa que o pre lcm e da integração
            if (String.IsNullOrWhiteSpace(identificadorRh))
                return;

            Application.SBO_Application.SetStatusBarMessage("Não é permitido alterar lançamentos inseridos pela integração RH de forma manual.");
            BubbleEvent = false;
            btn_main.Item.Enabled = false;
        }

        private void PreLcm_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            string transId = edit_TransId.Value.ToString();
            string identificadorRh = GetSdrIntRhValue(transId);

            //caso seja nulo significa que o pre lcm e da integração
            if (String.IsNullOrWhiteSpace(identificadorRh))
                return;

            ChangeFormState(false);
        }

        private string GetSdrIntRhValue(string transId)
        {
            if (String.IsNullOrWhiteSpace(transId))
                return "";
            string query = $"SELECT \"U_SDR_IntRh\" FROM OJDT WHERE \"TransId\" = {transId} ";

            SAPbobsCOM.Recordset oRec = ((SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

            oRec.DoQuery(query);

            if (oRec.RecordCount == 0) return "";

            if (oRec.Fields.Item("U_SDR_IntRh").Value == null) return "";

            return oRec.Fields.Item("U_SDR_IntRh").Value.ToString();
        }

        private void ChangeFormState(bool newState)
        {
            try
            {
                edit_duedate.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                edit_memo.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                cmb_indicator.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                edit_project.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                cmb_transcode.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                edit_refOne.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                edit_refTwo.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                edit_refThree.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                cmb_ecdType.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                cmb_matriz.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                btn_main.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                mtx_Lines.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                chk_cambio.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                chk_estorno.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                chk_comp.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

        }
    }
}

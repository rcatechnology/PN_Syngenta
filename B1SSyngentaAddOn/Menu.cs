using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace B1SSyngentaAddOn
{
    class Menu
    {
        public void AddMenuItems()
        {
            /*
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Application.SBO_Application.Menus.Item("43520"); // moudles'

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "PN_Syngenta";
            oCreationPackage.String = "PN_Syngenta";
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = -1;

            oMenus = oMenuItem.SubMenus;

            try
            {
                //  If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception e)
            {

            }

            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("PN_Syngenta");
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "PN_Syngenta.Form1";
                oCreationPackage.String = "Form1";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception er)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            */
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            

            try
            {
                SAPbouiCOM.Form form = Application.SBO_Application.Forms.ActiveForm;
                if (pVal.MenuUID == "6913" && (form.TypeEx == "50101" || form.TypeEx == "50102"))
                {
                    BubbleEvent = false;
                    return;
                }
            }
            catch (Exception ex)
            {
                //Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
            
        }

    }
}

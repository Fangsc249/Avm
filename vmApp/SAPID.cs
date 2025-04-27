namespace vmApp
{
    class SAPID
    {
        /*
         * 应该是彻底解决了在Code指令加载一些类型定义，在后续Code指令种使用的问题，
         * 这样以来就可以在程序中以及yaml脚本中共用这里的SAPID
         * 2025-4-22
         */
        public static string Q01 => "Q01 One ERP Quality Assurance System";
        public static string Q08 => "Q08 Quality Assurance System";
        public static string P01 => "P01 One ERP Production System";
        public static string SRP => "SRP PS ERP Production System";
        public static string SRQ => "SRQ PS ERP Quality Assurance System";
        public static string P24 => "P24 APO Production System";
        public static string ERP => "AAA-EHP6";
        public static string APP => "APP PS SCM Production System";
        /// <summary>
        /// Logon UserID
        /// </summary>
        public static string ID_USERID => @"wnd[0]/usr/txtRSYST-BNAME";
        /// <summary>
        /// Logon Password
        /// </summary>
        public static string ID_PASSWORD => @"wnd[0]/usr/pwdRSYST-BCODE";
        /// <summary>
        /// Logon Lang Code
        /// </summary>
        public static string ID_LANGCODE => @"wnd[0]/usr/txtRSYST-LANGU";

        #region VL06O
        /// <summary>
        /// VL06O ShipPoint 多选输入框。
        /// </summary>
        public static string ID_VL06O_SHIPPOINT_MULTI_BUTTON
            => @"wnd[0]/usr/btn%_IF_VSTEL_%_APP_%-VALU_PUSH";
        /// <summary>
        /// VL06O ShipPoint 多选输入框，删除所有
        /// </summary>
        public static string ID_VL06O_SPWINDOW_DELETE_ALL_SHIPPOINT_BUTTON
            => @"wnd[1]/tbar[0]/btn[16]";//DELETE ENTIRE SELECTION, 清除旧的选择，如果有的化。
        /// <summary>
        /// 
        /// </summary>
        public static string ID_VL06O_SPWINDOW_COPY_F8_BUTTON
            => @"wnd[1]/tbar[0]/btn[8]"; // 应用填充的多个ShipPoint
        /// <summary>
        /// 
        /// </summary>
        public static string ID_VL06O_SPWINDOW_CHECK_BUTTON
            => @"wnd[1]/tbar[0]/btn[0]";
        /// <summary>
        /// 
        /// </summary>
        public static string ID_VL06O_SPWINDOW_TABLECONTROL
            => @"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE";
        /// <summary>
        /// 
        /// </summary>
        public static string ID_VL06O_LAYOUT_SELECT_GRID // 选择Layout对话框中的GridVIew
            => @"wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell";
        /// <summary>
        /// MM60
        /// </summary>
        public static string ID_VL06O_OUTBOUND_BTN => @"wnd[0]/usr/btnBUTTON6";
        #endregion VL06O

        #region VL03N
        /// <summary>
        /// 
        /// </summary>
        public static string ID_SAPLV51GTC_HU_003_GuiTableControl
            => @"wnd[0]/usr/tabsTS_HU_VERP/tabpUE6HUS/ssubTAB:SAPLV51G:6020/tblSAPLV51GTC_HU_003";
        /// <summary>
        /// 
        /// </summary>
        public static string ID_VL03N_HU_OUTPUT_TABLE
            => @"wnd[0]/usr/tblSAPDV70ATC_NAST3";
        /// <summary>
        /// VL03N 
        /// </summary>
        public static string ID_VL03N_HU_OUTPUT_MENU
                => @"wnd[0]/mbar/menu[3]/menu[4]";
        #endregion VL03N

        #region MM60
        /// <summary>
        /// MM60 Material 
        /// </summary>
        public static string ID_MM60_MATERIAL => @"wnd[0]/usr/ctxtMS_MATNR-LOW";
        /// <summary>
        /// MM60 Plant
        /// </summary>
        public static string ID_MM60_PLANT => @"wnd[0]/usr/ctxtMS_WERKS-LOW";
        /// <summary>
        /// MM60 Query Result GridView
        /// </summary>
        public static string ID_MM60_GRID => @"wnd[0]/usr/cntlGRID1/shellcont/shell";
        /// <summary>
        /// MM60 Material Type
        /// </summary>
        public static string ID_MM60_MATERIAL_TYPE => @"wnd[0]/usr/ctxtMTART-LOW";

        #endregion MM60
        #region CS03
        /// <summary>
        /// CS03 MATERIAL
        /// </summary>
        public static string ID_CS03_MATERIAL => @"wnd[0]/usr/ctxtRC29N-MATNR";
        /// <summary>
        /// CS03 PLANT
        /// </summary>
        public static string ID_CS03_PLANT => @"wnd[0]/usr/ctxtRC29N-WERKS";
        /// <summary>
        /// CS03 BOM USAGE
        /// </summary>
        public static string ID_CS03_BOM_USAGE => @"wnd[0]/usr/ctxtRC29N-STLAN";
        /// <summary>
        /// CS03 TAB STRIP
        /// </summary>
        public static string ID_CS03_TABSTRIP => @"wnd[0]/usr/tabsTS_ITOV";
        /// <summary>
        /// CS03 BOM TABLECONTROL
        /// </summary>
        public static string ID_CS03_BOM_TABLE => @"wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT";
        /// <summary>
        /// 存在多个Alternative BOM时出现的一个表格，从中选择一个BOM
        /// </summary>
        public static string ID_CS03_ALT_BOM_TABLE => @"wnd[0]/usr/tblSAPLCSDITCALT";
        #endregion CS03
    }
}

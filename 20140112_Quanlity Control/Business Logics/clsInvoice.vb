Public Class clsInvoice
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Order Or pVal.FormTypeEx = frm_Delivery Or pVal.FormTypeEx = frm_Invoice Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.CharPressed <> 9 And pVal.ItemUID = "edCR" Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                               
                                If pVal.FormTypeEx = frm_Order Then
                                    oApplication.Utilities.AddControls(oForm, "stCR", "86", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", , , , "Crew Briefing Ref")
                                    oApplication.Utilities.AddControls(oForm, "edCR", "46", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN")
                                    oEditText = oForm.Items.Item("edCR").Specific
                                    oEditText.DataBind.SetBound(True, "ORDR", "U_Z_CrewRef")
                                    oForm.Items.Item("edCR").Enabled = False

                                ElseIf pVal.FormTypeEx = frm_Delivery Then
                                    oApplication.Utilities.AddControls(oForm, "stCR", "86", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", , , , "Crew Briefing Ref")
                                    oApplication.Utilities.AddControls(oForm, "edCR", "46", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN")
                                    oEditText = oForm.Items.Item("edCR").Specific
                                    oEditText.DataBind.SetBound(True, "ODLN", "U_Z_CrewRef")
                                    oForm.Items.Item("edCR").Enabled = False
                                ElseIf pVal.FormTypeEx = frm_Invoice Then
                                    oApplication.Utilities.AddControls(oForm, "stCR", "86", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", , , , "Crew Briefing Ref")
                                    oApplication.Utilities.AddControls(oForm, "edCR", "46", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN")
                                    oEditText = oForm.Items.Item("edCR").Specific
                                    oEditText.DataBind.SetBound(True, "OINV", "U_Z_CrewRef")
                                    oForm.Items.Item("edCR").Enabled = False
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                        End Select
                End Select
            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region
    Private Function AddtoUDT(ByVal aform As SAPbouiCOM.Form, ByVal ItemCode As String, ByVal ItemName As String) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim otemp, otemp1 As SAPbobsCOM.Recordset
        Dim strqry, strCode, strqry1, strProCode, ProName, strGLAcc As String
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oUserTable = oApplication.Company.UserTables.Item("Z_OCBS")

        If ItemCode <> "" Then
            strCode = ItemCode
            oUserTable.GetByKey(strCode)
            oUserTable.Code = strCode
            oUserTable.Name = strCode
            oUserTable.Update()
        Else
            strCode = oApplication.Utilities.getMaxCode("@Z_OCBS", "Code")
            oUserTable.Code = strCode
            oUserTable.Name = strCode
            oUserTable.Add()
        End If
        oApplication.Utilities.SetEditText(aform, "edCR", strCode)

        Return True
    End Function

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "CR"
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    Dim ItemCode, ItemName As String
                    
                    If pVal.BeforeAction = False Then
                        If oForm.TypeEx = frm_Order Then
                            If oApplication.Utilities.getEdittextvalue(oForm, "4") <> "" Then
                                ItemCode = oApplication.Utilities.getEdittextvalue(oForm, "edCR")
                                ItemName = oApplication.Utilities.getEdittextvalue(oForm, "8")
                                If AddtoUDT(oForm, ItemCode, ItemName) = True Then
                                    ItemCode = oApplication.Utilities.getEdittextvalue(oForm, "edCR")
                                    ItemName = oApplication.Utilities.getEdittextvalue(oForm, "8")
                                    Dim objct As New clsQCMaster
                                    objct.LoadForm(ItemCode, ItemName, "SalesOrder")
                                End If
                            Else
                                oApplication.Utilities.Message("Select the Customer... ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        Else
                            ItemCode = oApplication.Utilities.getEdittextvalue(oForm, "edCR")
                            If ItemCode = "" Then
                                oApplication.Utilities.Message("Crew Briefing Sheet details are not available", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Sub
                            Else
                                Dim objct As New clsQCMaster
                                objct.LoadForm(ItemCode, "", "DelInvoice")
                            End If
                        End If
                       

                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
        'If eventInfo.FormUID = "RightClk" Then
        If oForm.TypeEx = frm_Order Or oForm.TypeEx = frm_Delivery Or oForm.TypeEx = frm_Invoice Then
            If (eventInfo.BeforeAction = True) Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                        oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = "CR"
                        oCreationPackage.String = "Crew Breifing Sheet"
                        oCreationPackage.Enabled = True
                        oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                        oMenus = oMenuItem.SubMenus
                        oMenus.AddEx(oCreationPackage)


                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Else
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        oApplication.SBO_Application.Menus.RemoveEx("CR")
                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        End If
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()

            End If
           
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class

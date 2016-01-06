
Public Class clsQCMaster
    Inherits clsBase
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oCheckbox As SAPbouiCOM.CheckBox
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private oTemp As SAPbobsCOM.Recordset
    Private InvBaseDocNo, strname As String
    Private InvForConsumedItems As Integer
    Private oMenuobject As Object
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Public Sub LoadForm(ByVal ItemCode As String, ByVal ItemName As String, ByVal FormType As String)
        oForm = oApplication.Utilities.LoadForm(xml_QCItemMaster, frm_QCItemMaster)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oApplication.Utilities.setEdittextvalue(oForm, "7", ItemCode)
        If FormType = "SalesOrder" Then
            oApplication.Utilities.setEdittextvalue(oForm, "9", ItemName)
            oForm.Items.Item("9").Visible = True
            oForm.Items.Item("8").Visible = True
            oForm.Items.Item("13").Visible = True
            oForm.Items.Item("3").Visible = True
            oForm.Items.Item("4").Visible = True
        Else
            oForm.Items.Item("9").Visible = False
            oForm.Items.Item("8").Visible = False
            oForm.Items.Item("13").Visible = False
            oForm.Items.Item("3").Visible = False
            oForm.Items.Item("4").Visible = False
        End If
        Databind(oForm, ItemCode, FormType)
        'AddChooseFromList(oForm)
    End Sub

#Region "Databind"
    Private Sub Databind(ByVal aform As SAPbouiCOM.Form, ByVal ItemCode As String, ByVal FormType As String)
        Try
            aform.Freeze(True)
            oGrid = aform.Items.Item("5").Specific
            dtTemp = oGrid.DataTable
            '   dtTemp.ExecuteQuery("Select * from [@Z_CBS1] where U_Z_RefCode='" & ItemCode & "' order by Code")
            dtTemp.ExecuteQuery("SELECT T0.[Code], T0.[Name], T0.[U_Z_Dept], T0.[U_Z_DepDate], T0.[U_Z_ETD], T0.[U_Z_Arrival], T0.[U_Z_ArvDate], T0.[U_Z_ETA], T0.[U_Z_RefCode] FROM [dbo].[@Z_CBS1]  T0 where U_Z_RefCode='" & ItemCode & "' order by Code")
            oGrid.DataTable = dtTemp
            Formatgrid(oGrid, FormType)
            assignLineNo(aform)
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub
#End Region

    Private Sub assignLineNo(ByVal aForm As SAPbouiCOM.Form)
        aForm.Freeze(True)
        oGrid = aForm.Items.Item("5").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            '   oGrid.DataTable.SetValue("RowId", intRow, intRow + 1)
            oGrid.RowHeaders.SetText(intRow, intRow + 1)
        Next
        aForm.Freeze(False)
    End Sub

#Region "FormatGrid"
    Private Sub Formatgrid(ByVal agrid As SAPbouiCOM.Grid, ByVal FormType As String)
        If FormType = "SalesOrder" Then
            agrid.Columns.Item("Code").Visible = False
            agrid.Columns.Item("Name").Visible = False
            agrid.Columns.Item("U_Z_Dept").TitleObject.Caption = "Departure"
            agrid.Columns.Item("U_Z_DepDate").TitleObject.Caption = "Dep.Date"
            agrid.Columns.Item("U_Z_ETD").TitleObject.Caption = "ETD"
            agrid.Columns.Item("U_Z_Arrival").TitleObject.Caption = "Arrival"
            agrid.Columns.Item("U_Z_ArvDate").TitleObject.Caption = "Arv.Date"
            agrid.Columns.Item("U_Z_ETA").TitleObject.Caption = "ETA"
            agrid.Columns.Item("U_Z_RefCode").TitleObject.Caption = "Reference Code"
            agrid.Columns.Item("U_Z_RefCode").Visible = False
            agrid.Columns.Item("RowsHeader").TitleObject.Caption = "#"
        Else
            '   agrid.Columns.Item("RowsHeader").Visible = False
            agrid.Columns.Item("Code").Visible = False
            agrid.Columns.Item("Name").Visible = False
            agrid.Columns.Item("U_Z_Dept").TitleObject.Caption = "Departure"
            agrid.Columns.Item("U_Z_Dept").Editable = False
            agrid.Columns.Item("U_Z_DepDate").TitleObject.Caption = "Dep.Date"
            agrid.Columns.Item("U_Z_DepDate").Editable = False
            agrid.Columns.Item("U_Z_ETD").TitleObject.Caption = "ETD"
            agrid.Columns.Item("U_Z_ETD").Editable = False
            agrid.Columns.Item("U_Z_Arrival").TitleObject.Caption = "Arrival"
            agrid.Columns.Item("U_Z_Arrival").Editable = False
            agrid.Columns.Item("U_Z_ArvDate").TitleObject.Caption = "Arv.Date"
            agrid.Columns.Item("U_Z_ArvDate").Editable = False
            agrid.Columns.Item("U_Z_ETA").TitleObject.Caption = "ETA"
            agrid.Columns.Item("U_Z_ETA").Editable = False
            agrid.Columns.Item("U_Z_RefCode").TitleObject.Caption = "Reference Code"
            agrid.Columns.Item("U_Z_RefCode").Visible = False
            agrid.Columns.Item("RowsHeader").TitleObject.Caption = "#"
        End If
        agrid.AutoResizeColumns()
        agrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
#End Region

    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        oGrid = aForm.Items.Item("5").Specific
        oEditTextColumn = oGrid.Columns.Item("U_Z_Dept")

        Dim strCode As String
        If oGrid.DataTable.Rows.Count - 1 <= 0 Then
            oGrid.DataTable.Rows.Add()
        End If
        'oEditTextColumn = oGrid.Columns.Item("U_Z_HRPeoobjCode")
        strCode = oEditTextColumn.GetText(oGrid.DataTable.Rows.Count - 1).ToString
        ' strCode = oEditTextColumn.GetTex(oGrid.DataTable.Rows.Count - 1).Value
        If strCode <> "" Then
            oGrid.DataTable.Rows.Add()
            If aForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And aForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        End If
        assignLineNo(aForm)
        oGrid.Columns.Item("RowsHeader").Click(oGrid.DataTable.Rows.Count - 1, False)
        oGrid.Columns.Item("U_Z_Dept").Click(oGrid.DataTable.Rows.Count - 1, False)
    End Sub
#Region "DeleteRow"
    Private Sub DeleteRow(ByVal aForm As SAPbouiCOM.Form)

        oGrid = aForm.Items.Item("5").Specific
        Dim strCode As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.Rows.IsSelected(intRow) Then
                strCode = oGrid.DataTable.GetValue("Code", intRow)
                oTemp.DoQuery("Update [@Z_CBS1] set Name=Name+'_XD' where Code='" & strCode & "'")
                oGrid.DataTable.Rows.Remove(intRow)
                'If aForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And aForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                '    aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                'End If
                assignLineNo(aForm)
                Exit Sub
            End If
        Next
    End Sub

#End Region
    Private Function Validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strSql, strETD, strETA, strDepdate, strArrDate As String
        Dim oRec, oRec1 As SAPbobsCOM.Recordset
        Dim Depdate, Arrivedate As Date
        oGrid = aform.Items.Item("5").Specific
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1

            If oGrid.DataTable.GetValue("U_Z_Dept", intRow) <> "" Then
                strDepdate = oGrid.DataTable.GetValue("U_Z_DepDate", intRow)
                strArrDate = oGrid.DataTable.GetValue("U_Z_ArvDate", intRow)
                strETA = oGrid.DataTable.GetValue("U_Z_ETA", intRow)
                strETD = oGrid.DataTable.GetValue("U_Z_ETD", intRow)
                If strDepdate = "" Then
                    oApplication.Utilities.Message("Enter Departure Date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If strETD = "" Then
                    oApplication.Utilities.Message("Enter ETD", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If CInt(strETD) = 0 Then
                    oApplication.Utilities.Message("Enter ETD", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If oGrid.DataTable.GetValue("U_Z_Arrival", intRow) = "" Then
                    oApplication.Utilities.Message("Enter Arrival", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If strArrDate = "" Then
                    oApplication.Utilities.Message("Enter Arrival Date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If strETD = "" Then
                    oApplication.Utilities.Message("Enter ETA", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If CInt(strETD) = 0 Then
                    oApplication.Utilities.Message("Enter ETA", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                Dim Etime, Earrive As Integer
                Depdate = oGrid.DataTable.GetValue("U_Z_DepDate", intRow)
                Arrivedate = oGrid.DataTable.GetValue("U_Z_ArvDate", intRow)
                Earrive = oGrid.DataTable.GetValue("U_Z_ETA", intRow)
                Etime = oGrid.DataTable.GetValue("U_Z_ETD", intRow)
                strETA = Earrive.ToString("00:00")
                strETD = Etime.ToString("00:00")
                strSql = "Select * from [@Z_CBS1] where '" & Depdate.ToString("yyyy-MM-dd") & "' = '" & Arrivedate.ToString("yyyy-MM-dd") & "'"
                oRec.DoQuery(strSql)
                If oRec.RecordCount > 0 Then
                    strSql = "Select * from [@Z_CBS1] where '" & Depdate.ToString("yyyy-MM-dd") & " " & strETD & "' < '" & Arrivedate.ToString("yyyy-MM-dd") & " " & strETA & "'"
                    oRec1.DoQuery(strSql)
                    If oRec1.RecordCount = 0 Then
                        oApplication.Utilities.Message("Arrival time should be greater than departure time...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oGrid.Columns.Item("U_Z_ETA").Click(intRow)
                        Return False
                    End If
                Else
                    strSql = "Select * from [@Z_CBS1] where '" & Depdate.ToString("yyyy-MM-dd") & "' > '" & Arrivedate.ToString("yyyy-MM-dd") & "'"
                    oRec.DoQuery(strSql)
                    If oRec.RecordCount > 0 Then
                        oApplication.Utilities.Message("Arrival Date should be greater than or equal to departure date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oGrid.Columns.Item("U_Z_ArvDate").Click(intRow)
                        Return False
                    End If
                End If
            End If
        Next
        Return True
    End Function
#Region "AddtoUDT"
    Private Function AddtoUDT1(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim oRec As SAPbobsCOM.Recordset
        Dim strCode, strECode, strESocial, strEname, strETax, strGLAcc As String
        Dim OCHECKBOXCOLUMN As SAPbouiCOM.CheckBoxColumn
        Dim Etime, Earrive As Integer
        oGrid = aform.Items.Item("5").Specific
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oUserTable = oApplication.Company.UserTables.Item("Z_CBS1")
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If (oGrid.DataTable.GetValue("U_Z_Dept", intRow)) <> "" Then
                strCode = oGrid.DataTable.GetValue("Code", intRow)
                Etime = oGrid.DataTable.GetValue("U_Z_ETD", intRow)
                Earrive = oGrid.DataTable.GetValue("U_Z_ETA", intRow)
                If oUserTable.GetByKey(strCode) Then
                    Dim str As String = Etime.ToString("00:00")
                    Dim strArrive As String = Earrive.ToString("00:00")
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_Dept").Value = (oGrid.DataTable.GetValue("U_Z_Dept", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_DepDate").Value = (oGrid.DataTable.GetValue("U_Z_DepDate", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_ETD").Value = str ' (oGrid.DataTable.GetValue("U_Z_ETD", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_Arrival").Value = (oGrid.DataTable.GetValue("U_Z_Arrival", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_ArvDate").Value = (oGrid.DataTable.GetValue("U_Z_ArvDate", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_ETA").Value = strArrive ' (oGrid.DataTable.GetValue("U_Z_ETA", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_RefCode").Value = oApplication.Utilities.getEdittextvalue(aform, "7")
                    If oUserTable.Update() <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                Else
                    'Dim int As Integer = 320
                    Dim str As String = Etime.ToString("00:00")
                    Dim strArrive As String = Earrive.ToString("00:00")
                    strCode = oApplication.Utilities.getMaxCode("@Z_CBS1", "Code")
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_Dept").Value = (oGrid.DataTable.GetValue("U_Z_Dept", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_DepDate").Value = (oGrid.DataTable.GetValue("U_Z_DepDate", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_ETD").Value = str '(oGrid.DataTable.GetValue("U_Z_ETD", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_Arrival").Value = (oGrid.DataTable.GetValue("U_Z_Arrival", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_ArvDate").Value = (oGrid.DataTable.GetValue("U_Z_ArvDate", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_ETA").Value = strArrive ' (oGrid.DataTable.GetValue("U_Z_ETA", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_RefCode").Value = oApplication.Utilities.getEdittextvalue(aform, "7")
                    If oUserTable.Add() <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            End If
        Next
        oRec.DoQuery("Delete from [@Z_CBS1] where Name like '%_XD' and U_Z_RefCode='" & oApplication.Utilities.getEdittextvalue(aform, "7") & "'")
        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Databind(aform, oApplication.Utilities.getEdittextvalue(aform, "7"), "SalesOrder")
    End Function
#End Region

    Private Sub Committrans(ByVal aForm As SAPbouiCOM.Form, ByVal aChoice As String)
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Update [@Z_CBS1] set Name=Code where Name like '%_XD' and U_Z_RefCode='" & oApplication.Utilities.getEdittextvalue(aForm, "7") & "'")
    End Sub



#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_QCItemMaster Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "2" Then
                                    Committrans(oForm, "Cancel")
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oGrid = oForm.Items.Item("5").Specific
                                'If pVal.ItemUID = "5" And pVal.ColUID = "Code" And pVal.CharPressed <> 9 Then
                                '    If oGrid.DataTable.GetValue("Ref", pVal.Row) <> "" Then
                                '        BubbleEvent = False
                                '        Exit Sub
                                '    End If
                                'End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "13" Then
                                    oGrid = oForm.Items.Item("5").Specific
                                    If Validation(oForm) = True Then
                                        AddtoUDT1(oForm)
                                    Else
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                               
                                If pVal.ItemUID = "3" Then
                                    oGrid = oForm.Items.Item("5").Specific
                                    AddRow(oForm)
                                End If
                                If pVal.ItemUID = "4" Then
                                    oGrid = oForm.Items.Item("5").Specific
                                    DeleteRow(oForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim oItm As SAPbobsCOM.Items
                                Dim sCHFL_ID, val As String
                                Dim intChoice, introw As Integer
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        oForm.Freeze(True)
                                        oForm.Update()
                                        'If pVal.ItemUID = "5" Then
                                        '    oGrid = oForm.Items.Item("5").Specific
                                        '    val = oDataTable.GetValue("FormatCode", 0)
                                        '    Try

                                        '        oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, val)
                                        '    Catch ex As Exception
                                        '    End Try
                                        'End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                    'MsgBox(ex.Message)
                                End Try
                        End Select
                End Select
            End If

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID

                Case mnu_ADD_ROW
                Case mnu_DELETE_ROW
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                'Select Case pVal.MenuUID
                '    Case mnu_LeaveMaster
                '        oMenuobject = New clsEarning
                '        oMenuobject.MenuEvent(pVal, BubbleEvent)
                'End Select
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class

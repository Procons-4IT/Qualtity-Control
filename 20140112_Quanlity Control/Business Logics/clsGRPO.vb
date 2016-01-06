Public Class clsGRPO
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
    Private oDT_GrossProfit As DataTable
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub populateGrossprofitBase(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim frmSource As SAPbouiCOM.Form
            Dim oSourcematrix As SAPbouiCOM.Matrix
            Dim strBaseType, strBaseEntry, strBaseLine As String
            Dim intSOEntry, intSOLine, intPOEntry, intPOLine, intPIEntry, intPILine As Integer
            Dim dblsalesprice, dblPurchaseprice As Double
            frmSource = aform
            oSourcematrix = frmSource.Items.Item("38").Specific

            Dim oRs As SAPbobsCOM.Recordset
            oRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oDT_GrossProfit = New DataTable()
            oDT_GrossProfit.Columns.Add("ARLineNum")
            oDT_GrossProfit.Columns.Add("ARItemCode")
            oDT_GrossProfit.Columns.Add("SaleBaseType")
            oDT_GrossProfit.Columns.Add("SaleBaseEntry")
            oDT_GrossProfit.Columns.Add("SaleBaseLine")
            oDT_GrossProfit.Columns.Add("GPBasePrice")

            For intRow As Integer = 1 To oSourcematrix.VisualRowCount
                If oSourcematrix.Columns.Item("1").Cells.Item(intRow).Specific.Value.ToString() = "" Then
                Else
                    If oSourcematrix.Columns.Item("43").Cells.Item(intRow).Specific.Value.ToString() = "17" Then
                        Dim dr As DataRow
                        Dim oStr As String = "Select isnull(T0.DropShip,'N') from  OWHS T0  Where T0.WhsCode='" & oSourcematrix.Columns.Item("24").Cells.Item(intRow).Specific.Value.ToString() & "'"
                        Dim oTest As SAPbobsCOM.Recordset
                        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oTest.DoQuery(oStr)
                        If 1 = 1 Then 'oTest.Fields.Item(0).Value = "Y" Then
                            Dim strQry As String = "Select * from POR1 T1 join OPOR T0 on T1.DocEntry=T0.docEntry Where T1.TargetType=18 and T1.BaseType='17' and T1.BaseEntry='" & oSourcematrix.Columns.Item("44").Cells.Item(intRow).Specific.Value.ToString() & "' and T1.BaseLine='" & oSourcematrix.Columns.Item("46").Cells.Item(intRow).Specific.Value.ToString() & "'"
                            oRs.DoQuery(strQry)
                            If oRs.RecordCount > 0 Then
                                strQry = "Select T1.Price as 'GPPrice' from PCH1 T1 join OPCH T0 on T1.DocEntry=T0.DocEntry where T1.BaseType=22 and T1.BaseEntry=" & oRs.Fields.Item("DocEntry").Value & " and T1.BaseLine=" & oRs.Fields.Item("LineNum").Value
                                oRs.DoQuery(strQry)
                                If oRs.RecordCount > 0 Then
                                    Dim BasePrice As Double = oRs.Fields.Item("GPPrice").Value
                                    dr = oDT_GrossProfit.NewRow()
                                    dr("ARLineNum") = intRow
                                    dr("ARItemCode") = oSourcematrix.Columns.Item("1").Cells.Item(intRow).Specific.Value.ToString()
                                    dr("SaleBaseType") = oSourcematrix.Columns.Item("43").Cells.Item(intRow).Specific.Value.ToString()
                                    dr("SaleBaseEntry") = oSourcematrix.Columns.Item("44").Cells.Item(intRow).Specific.Value.ToString()
                                    dr("SaleBaseLine") = oSourcematrix.Columns.Item("46").Cells.Item(intRow).Specific.Value.ToString()
                                    dr("GPBasePrice") = BasePrice
                                    oDT_GrossProfit.Rows.Add(dr)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            If Not IsNothing(oDT_GrossProfit) And oDT_GrossProfit.Rows.Count > 0 Then
                Try
                    oApplication.SBO_Application.ActivateMenuItem(mnu_Grossprofit)
                Catch ex As Exception
                End Try

                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If oForm.TypeEx = frm_Grossprofit Then
                    Dim oMatrix_GP = oForm.Items.Item("3").Specific
                    Dim dr4 As DataRow
                    For intRow As Integer = 1 To oDT_GrossProfit.Rows.Count
                        dr4 = oDT_GrossProfit.Rows(intRow - 1)
                        oMatrix_GP.Columns.Item("3").Cells.Item(Convert.ToInt32(dr4("ARLineNum").ToString())).Specific.Value = dr4("GPBasePrice").ToString()
                    Next
                End If
                Try

                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Catch ex As Exception
                    aform.Freeze(False)
                End Try
                aform.Freeze(False)
            End If
            aform.Freeze(False)

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub

    Private Sub populateGrossprofitBase_menuClick(ByVal aform As SAPbouiCOM.Form, ByVal aSource As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim frmSource As SAPbouiCOM.Form
            Dim oSourcematrix As SAPbouiCOM.Matrix
            Dim strBaseType, strBaseEntry, strBaseLine As String
            Dim intSOEntry, intSOLine, intPOEntry, intPOLine, intPIEntry, intPILine As Integer
            Dim dblsalesprice, dblPurchaseprice As Double
            frmSource = aform
            oSourcematrix = frmSource.Items.Item("38").Specific

            Dim oRs As SAPbobsCOM.Recordset
            oRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oDT_GrossProfit = New DataTable()
            oDT_GrossProfit.Columns.Add("ARLineNum")
            oDT_GrossProfit.Columns.Add("ARItemCode")
            oDT_GrossProfit.Columns.Add("SaleBaseType")
            oDT_GrossProfit.Columns.Add("SaleBaseEntry")
            oDT_GrossProfit.Columns.Add("SaleBaseLine")
            oDT_GrossProfit.Columns.Add("GPBasePrice")

            For intRow As Integer = 1 To oSourcematrix.VisualRowCount
                If oSourcematrix.Columns.Item("1").Cells.Item(intRow).Specific.Value.ToString() = "" Then
                Else
                    If oSourcematrix.Columns.Item("43").Cells.Item(intRow).Specific.Value.ToString() = "17" Then
                        Dim dr As DataRow
                        Dim oStr As String = "Select isnull(T0.DropShip,'N') from  OWHS T0  Where T0.WhsCode='" & oSourcematrix.Columns.Item("24").Cells.Item(intRow).Specific.Value.ToString() & "'"
                        Dim oTest As SAPbobsCOM.Recordset
                        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oTest.DoQuery(oStr)
                        If 1 = 1 Then 'oTest.Fields.Item(0).Value = "Y" Then
                            Dim strQry As String = "Select T1.Price as 'GPPrice',*  from POR1 T1 join OPOR T0 on T1.DocEntry=T0.docEntry Where T1.TargetType=18 and T1.BaseType='17' and T1.BaseEntry='" & oSourcematrix.Columns.Item("44").Cells.Item(intRow).Specific.Value.ToString() & "' and T1.BaseLine='" & oSourcematrix.Columns.Item("46").Cells.Item(intRow).Specific.Value.ToString() & "'"
                            oRs.DoQuery(strQry)
                            If oRs.RecordCount > 0 Then
                                strQry = "Select T1.Price as 'GPPrice' from PCH1 T1 join OPCH T0 on T1.DocEntry=T0.DocEntry where T1.BaseType=22 and T1.BaseEntry=" & oRs.Fields.Item("DocEntry").Value & " and T1.BaseLine=" & oRs.Fields.Item("LineNum").Value
                                oRs.DoQuery(strQry)
                                If oRs.RecordCount > 0 Then
                                    Dim BasePrice As Double = oRs.Fields.Item("GPPrice").Value
                                    dr = oDT_GrossProfit.NewRow()
                                    dr("ARLineNum") = intRow
                                    dr("ARItemCode") = oSourcematrix.Columns.Item("1").Cells.Item(intRow).Specific.Value.ToString()
                                    dr("SaleBaseType") = oSourcematrix.Columns.Item("43").Cells.Item(intRow).Specific.Value.ToString()
                                    dr("SaleBaseEntry") = oSourcematrix.Columns.Item("44").Cells.Item(intRow).Specific.Value.ToString()
                                    dr("SaleBaseLine") = oSourcematrix.Columns.Item("46").Cells.Item(intRow).Specific.Value.ToString()
                                    dr("GPBasePrice") = BasePrice
                                    oDT_GrossProfit.Rows.Add(dr)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            If Not IsNothing(oDT_GrossProfit) And oDT_GrossProfit.Rows.Count > 0 Then
                'oForm = oApplication.SBO_Application.Forms.ActiveForm()
                oForm = aSource
                If oForm.TypeEx = frm_Grossprofit Then
                    Dim oMatrix_GP = oForm.Items.Item("3").Specific
                    Dim dr4 As DataRow
                    For intRow As Integer = 1 To oDT_GrossProfit.Rows.Count
                        dr4 = oDT_GrossProfit.Rows(intRow - 1)
                        oMatrix_GP.Columns.Item("3").Cells.Item(Convert.ToInt32(dr4("ARLineNum").ToString())).Specific.Value = dr4("GPBasePrice").ToString()
                    Next
                End If
                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                'oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                aform.Freeze(False)
            End If
            aform.Freeze(False)

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_GRPO Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    '  populateGrossprofitBase(oForm)
                                End If

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    '  oForm.Freeze(False)
                                End If


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
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS


            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region


    Private Function InitializeApprisal(ByVal aDocEntry As Integer, ByVal aCardcode As String, ByVal aCardName As String) As Boolean
        Try
            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            oApplication.Company.StartTransaction()
            Dim oUserTable As SAPbobsCOM.UserTable
            Dim strCode, strECode, strESocial, strEname, strETax, strGLAcc, strDept As String
            Dim oGeneralService, oGeneralService1 As SAPbobsCOM.GeneralService
            Dim oGeneralData, oGeneralData1 As SAPbobsCOM.GeneralData
            Dim oGeneralParams, oGeneralParams1 As SAPbobsCOM.GeneralDataParams
            Dim oCompanyService As SAPbobsCOM.CompanyService
            Dim oChildren, oChildren1, oChildren2 As SAPbobsCOM.GeneralDataCollection
            oCompanyService = oApplication.Company.GetCompanyService()
            Dim otestRs, oRec As SAPbobsCOM.Recordset
            Dim oChild, oChild1, oChild2, oChild3 As SAPbobsCOM.GeneralData
            Dim blnRecordExists As Boolean = False
            oGeneralService = oCompanyService.GetGeneralService("Z_OQCL")
            otestRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            Dim oCheckbox, ocheckbox1 As SAPbouiCOM.CheckBoxColumn
            Dim blnDownpayment As Boolean = False
            Dim blnDocumentItem As Boolean
            Dim ReStdate, reEndDate As Date
            '  oDtAppraisal.Rows.Add(oGrid.DataTable.Rows.Count)
            Dim strCardCode, strCardName As String
            strCardCode = ""
            strCardName = ""
            oRec.DoQuery("Select * from OPDN where isnull(U_Z_QCDocNum,'')='' and DocEntry=" & aDocEntry)
            oGeneralData1 = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            strCardCode = oRec.Fields.Item("CardCode").Value
            strCardName = oRec.Fields.Item("CardName").Value
            oGeneralData1.SetProperty("U_Z_GRPONO", oRec.Fields.Item("DocNum").Value.ToString)
            oGeneralData1.SetProperty("U_Z_Approved", "")
            oGeneralData1.SetProperty("U_Z_STNO", "")
            Try
                oGeneralData1.SetProperty("U_Z_Status", "O")
            Catch ex As Exception

            End Try

            oChildren1 = oGeneralData1.Child("Z_QCL1")
            oRec.DoQuery("Select * from PDN1 where DocEntry=" & aDocEntry)
            Dim blnline As Boolean = False
            For intRow As Integer = 0 To oRec.RecordCount - 1
                blnDocumentItem = False
                'oGeneralData1 = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                otestRs.DoQuery("Select isnull(U_Z_STWhsCode,'') from OITM where isnull(U_Z_QCReq,'N')='Y' and  ItemCode='" & oRec.Fields.Item("ItemCode").Value & "'")
                If otestRs.RecordCount > 0 Then
                    blnline = True
                    oChild = oChildren1.Add()
                    oChild.SetProperty("U_Z_ItemCode", oRec.Fields.Item("ItemCode").Value)
                    oChild.SetProperty("U_Z_ItemName", oRec.Fields.Item("Dscription").Value)
                    oChild.SetProperty("U_Z_CardCode", aCardcode)
                    oChild.SetProperty("U_Z_CardName", aCardName)
                    If otestRs.Fields.Item(0).Value = "" Then
                        oChild.SetProperty("U_Z_STWhsCode", oRec.Fields.Item("WhsCode").Value)
                    Else
                        oChild.SetProperty("U_Z_STWhsCode", otestRs.Fields.Item(0).Value)
                    End If
                    oChild.SetProperty("U_Z_QCWhsCode", oRec.Fields.Item("WhsCode").Value)

                    oChild.SetProperty("U_Z_AcctQty", CDbl(oRec.Fields.Item("Quantity").Value))
                    oChild.SetProperty("U_Z_RecQty", CDbl(oRec.Fields.Item("Quantity").Value))
                    oChild.SetProperty("U_Z_RegQty", 0)
                End If
                oRec.MoveNext()
            Next
            If blnline = True Then
                oGeneralParams = oGeneralService.Add(oGeneralData1)
                Dim strDocEntry As String = oGeneralParams.GetProperty("DocEntry")
                oRec.DoQuery("Update OPDN set U_Z_QCDocNum='" & strDocEntry & "' where DocEntry=" & aDocEntry)
                oApplication.SBO_Application.MessageBox("Quality Control Document created Successfully : " & strDocEntry)

            End If
            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            Return False
        End Try
    End Function

#Region "Form Data Event"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                Dim strDoc As String
                Dim oDoc As SAPbobsCOM.Documents
                oApplication.Company.GetNewObjectCode(strDoc)
                oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                If oDoc.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                    Dim ststatus As String = oDoc.DocumentStatus
                    If ststatus = 0 Then


                        InitializeApprisal(oDoc.DocEntry, oDoc.CardCode, oDoc.CardName)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region
End Class

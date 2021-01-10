VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetRoleAndWarehouseForm 
   Caption         =   "Select Role & Warehouse"
   ClientHeight    =   1845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5370
   OleObjectBlob   =   "SetRoleAndWarehouseForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SetRoleAndWarehouseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bInitializing As Boolean
Dim bInitializationError As Boolean
Dim startingRole As String
Dim startingWarehouse As String


Public Sub ShowMe(bForceOpen As Boolean)
    Dim sql As String
    bInitializationError = False
    'Need to check to see if we haven't already called this. It could happen when a use isn't logged in and they go to Configs and then open Roles & Warehouses
    If Not bInitializing Then
        bInitializing = True
        'Get warehouse and Role - This also makes sure we are logged in
        sql = "select IFNULL(current_role(),'') ||','|| IFNULL(current_warehouse(),'') ||','|| IFNULL(current_database(),'') ||','|| IFNULL(current_schema(),'')"
        dbObjsString = execSQLSingleValueOnly(sql)
        
        'split up result string into array
        arrDBObjs = Split(dbObjsString, ",")
        startingRole = arrDBObjs(0)
        startingWarehouse = arrDBObjs(1)
        'persist the warehouse and role
        Utils.CustomRange(sgRangeWarehouse) = startingWarehouse
        Utils.CustomRange(sgRangeRole) = startingRole
        If bForceOpen Or startingRole = "" Or startingWarehouse = "" Then
            ' This calls the FormCommon.SetRoleAndWarehouseFormInit which calls the UserForm_InitializeCustom Sub below. It allows the progress window to open
            Call StatusForm.execMethod("FormCommon", "SetRoleAndWarehouseFormInit")
            If startingWarehouse <> "" And Not bInitializationError Then
                Me.Show
            End If
        End If
        bInitializing = False
    End If 'bInitializing
End Sub
Public Sub UserForm_InitializeCustom()
   Dim currentRole As String
   Dim manuallyEnteredWarehouse As String
    If startingWarehouse = "" Then
        startingWarehouse = getManaullyEnteredWarehouse()
    End If
    'If there is still no valid warehouse then exit
    If startingWarehouse = "" Then
        StatusForm.Hide
        Exit Sub
    End If
    
    On Error GoTo ErrorHandlerInitialization
    bInitializing = True ' This prevents the roles drop down from excuting
    Call SetRoleAndWarehouseForm.getRoles(cbRoles)
    Call SetRoleAndWarehouseForm.getWarehouses(cbWarehouses)
    bInitializing = False
    StatusForm.Hide
    Exit Sub
ErrorHandlerInitialization:
    If err.Number <> giCancelEvent Then
        MsgBox "ERROR: Problem initializing: " & err.Description
    End If
    bInitializing = False
    bInitializationError = True
End Sub

Sub getWarehouses(cbWarehouses As comboBox)
    Dim tempWarehouseValue As String
    If cbWarehouses = "" Then
        tempWarehouseValue = startingWarehouse
    Else
        tempWarehouseValue = cbWarehouses
    End If
    'populate Warehouses Combobox
    Call FormCommon.getWarehousesCombobox(cbWarehouses)
    If cbWarehouses.ListCount = 0 Then
        tempWarehouseValue = getManaullyEnteredWarehouse()
    End If
       ' Select Role from list
    index = FormCommon.indexOfValueInList(cbWarehouses, tempWarehouseValue)
    If index > -1 Then
        cbWarehouses.ListIndex = index
    Else
        If cbWarehouses.ListCount > 0 Then
            cbWarehouses.ListIndex = 0
        End If
    End If
End Sub

Sub getRoles(cbRoles As comboBox)
    'populate Roles Combobox
    Call FormCommon.getRolesCombobox(cbRoles)
    ' Select Role from list
    index = FormCommon.indexOfValueInList(cbRoles, startingRole)
    If index > -1 Then
        cbRoles.ListIndex = index
    Else
        If cbRoles.ListCount > 0 Then
            cbRoles.ListIndex = 0
        End If
    End If
End Sub

Function getManaullyEnteredWarehouse()
    Set GetValueForm = Nothing
    GetValueForm.setMessage ("Unable to retrieve list of Warehouses.")
    GetValueForm.setValue ("Enter Warehouse")
RequestWarehouse:
    GetValueForm.Show
    If GetValueForm.okClicked Then
        manuallyEnteredWarehouse = GetValueForm.Getvalue()
        On Error GoTo ErrorHandlerInvalidWarehouse
        Call Utils.execSQLFireAndForget("use warehouse " & manuallyEnteredWarehouse)
        ' if we got this far then the warehouse is valid
        getManaullyEnteredWarehouse = manuallyEnteredWarehouse
        Utils.CustomRange(sgRangeWarehouse) = manuallyEnteredWarehouse
    Else
        MsgBox "A warehouse is needed before querying or loading data. "
        getManaullyEnteredWarehouse = ""
    End If
    Exit Function
ErrorHandlerInvalidWarehouse:
    MsgBox "The Warehouse entered does not exist."
    Resume RequestWarehouse
End Function

Private Sub btCancel_Click()
    If cbRoles <> startingRole Then
        Call Utils.execSQLFireAndForget("use role " & startingRole)
        Call Utils.execSQLFireAndForget("use warehouse " & startingWarehouse)
    End If
    Me.Hide
    Set SetRoleAndWarehouseForm = Nothing
End Sub

Private Sub UserForm_Initialize()
    'Center window in Excel
    Call FormCommon.setUserFormPosition(Me)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    btCancel_Click
End Sub

Private Sub btDone_Click()
    Me.Hide
    On Error GoTo ErrorHandlerSettingWarehouse
    Call Utils.execSQLFireAndForget("use warehouse " & cbWarehouses)
    Utils.CustomRange(sgRangeRole) = cbRoles
    Utils.CustomRange(sgRangeWarehouse) = cbWarehouses
    Set SetRoleAndWarehouseForm = Nothing
    Exit Sub
    
ErrorHandlerSettingWarehouse:
    MsgBox ("Unable to use selected warehouse: " & cbWarehouses)
    Me.Show
End Sub

Private Sub cbRoles_Click()
    If Not bInitializing Then
        On Error GoTo ErrorHandlerSetRole
        Call Utils.execSQLFireAndForget("use role " & cbRoles)
        Call SetRoleAndWarehouseForm.getWarehouses(cbWarehouses)
        Call FormCommon.dropDBObjectsCache
    End If
    Exit Sub
ErrorHandlerSetRole:
    MsgBox ("Error updating role: " & err.Description)
End Sub

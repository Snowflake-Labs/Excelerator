Attribute VB_Name = "PublicMacros"
Sub OpenUploadDataFormFromRibbon(control As IRibbonControl)
    setupSnowflakeIntegration

    If Not worksheetBelongsToAddin Then
        If Utils.login Then
            Set StatusForm = Nothing
            Call StatusForm.execMethod("Load", "OpenUploadDataForm")
        Else
            StatusForm.Hide
        End If
    End If
End Sub

Sub SetUpConfigDataFromRibbon(control As IRibbonControl)
    setupSnowflakeIntegration
    ConfigForm.SetUpConfigData
End Sub

Sub AddDataTypeDropDown(control As IRibbonControl)
    setupSnowflakeIntegration
    Call Load.AddDataTypeDropDowns
End Sub

Sub ConnectFromRibbon(control As IRibbonControl)
    setupSnowflakeIntegration
    Call Utils.Connect
End Sub

Sub ExecuteSelectAllFromTableFromRibbon(control As IRibbonControl)
    setupSnowflakeIntegration
    If Utils.login Then
        Set StatusForm = Nothing
        Call StatusForm.execMethod("Query", "ExecuteSelectAllFromUploadTable")
    End If
End Sub
Sub OpenSQLFormFromRibbon(control As IRibbonControl)
    setupSnowflakeIntegration
    ' If this is one of the Addins worksheet don't allow because you don't want to overwrite it
    If Not worksheetBelongsToAddin Then
        If Utils.login Then
            Set StatusForm = Nothing
            Call StatusForm.execMethod("Query", "OpenSQLForm")
        End If
        StatusForm.Hide
    End If
End Sub

Sub setupSnowflakeIntegration()
    Call Utils.CopySnowflakeConfgWS
End Sub

Sub RollbackLastUpdateFromRibbon(control As IRibbonControl)
    If Utils.login Then
        Call Load.RollbackLastUpdateWithCheck
    End If
End Sub

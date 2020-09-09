Attribute VB_Name = "PublicMacros"
Sub OpenUploadDataFormFromRibbon(control As IRibbonControl)
    If Not setupSnowflakeIntegration Then
        Exit Sub
    End If

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
    If Not setupSnowflakeIntegration Then
        Exit Sub
    End If
    ConfigForm.SetUpConfigData
End Sub

Sub AddDataTypeDropDown(control As IRibbonControl)
    If Not setupSnowflakeIntegration Then
        Exit Sub
    End If
    Call Load.AddDataTypeDropDowns
End Sub

Sub ConnectFromRibbon(control As IRibbonControl)
    If Not setupSnowflakeIntegration Then
        Exit Sub
    End If
    Call Utils.Connect
End Sub

Sub ExecuteSelectAllFromTableFromRibbon(control As IRibbonControl)
    If Not setupSnowflakeIntegration Then
        Exit Sub
    End If
    If Utils.login Then
        Set StatusForm = Nothing
        Call StatusForm.execMethod("Query", "ExecuteSelectAllFromUploadTable")
    End If
End Sub
Sub OpenSQLFormFromRibbon(control As IRibbonControl)
    If Not setupSnowflakeIntegration Then
        Exit Sub
    End If
    ' If this is one of the Addins worksheet don't allow because you don't want to overwrite it
    If Not worksheetBelongsToAddin Then
        If Utils.login Then
            Set StatusForm = Nothing
            Call StatusForm.execMethod("Query", "OpenSQLForm")
        End If
        StatusForm.Hide
    End If
End Sub

Function setupSnowflakeIntegration()
    If doesWorksheetExist Then
        Call Utils.CopySnowflakeConfgWS
        setupSnowflakeIntegration = True
    Else
        setupSnowflakeIntegration = False
    End If
End Function


Sub RollbackLastUpdateFromRibbon(control As IRibbonControl)
    If Utils.login Then
        Call Load.RollbackLastUpdateWithCheck
    End If
End Sub

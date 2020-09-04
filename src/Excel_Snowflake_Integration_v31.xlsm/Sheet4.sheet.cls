Private Sub ConnectImage_Click()
    Call Utils.Connect
End Sub

Private Sub RunMainButton_Click()
    If Utils.login Then
        Call SQLForm.ExecuteButton_Click
    End If
End Sub

Private Sub SelectFromUploadTable_click()
    If Utils.login Then
        Set StatusForm = Nothing
        Call StatusForm.execMethod("Query", "ExecuteSelectAllFromUploadTable")
    End If
End Sub

Private Sub SetupImage_Click()
    Call ConfigForm.SetUpConfigData
End Sub

Private Sub UploadButton_Click()
    Call Load.OpenUploadDataForm
End Sub
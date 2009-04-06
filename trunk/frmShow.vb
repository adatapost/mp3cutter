Public Class frmShow

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()

    End Sub
    Private Sub RemoveItemFromList(Optional ByRef SelectedItem As Boolean = True, Optional ByRef All As Boolean = False)
        '---------------
        Dim i As Integer
        '--------------
        If All = True Then
            frmMain.lstPlayList.Items.Clear()
            frmMain.lstPlayListPath.Items.Clear()
            Exit Sub
        End If
        '--------------
Re_search:
        For i = 0 To frmMain.lstPlayList.Items.Count - 1
            '----------------
            If frmMain.lstPlayList.GetSelected(i) = SelectedItem Then
                frmMain.lstPlayList.Items.RemoveAt((i))
                frmMain.lstPlayListPath.Items.RemoveAt((i))
                GoTo Re_search
            End If
            '-----------
            System.Windows.Forms.Application.DoEvents()
        Next i
        '---------------
    End Sub
    Private Sub FixNumbers()
        '---------------
        Dim i As Integer
        '---------------
        With frmMain.lstPlayList
            For i = 0 To .Items.Count - 1
                VB6.SetItemString(frmMain.lstPlayList, i, i + 1 & ". " & Mid(VB6.GetItemString(frmMain.lstPlayList, i), InStr(VB6.GetItemString(frmMain.lstPlayList, i), ".") + 1))
                System.Windows.Forms.Application.DoEvents()
            Next i
        End With
        '---------------
    End Sub

    '==============================================================
   
    '==============================================================

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        RemoveItemFromList(True)
        FixNumbers()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        RemoveItemFromList(False)
        FixNumbers()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        RemoveItemFromList(, True)

    End Sub
End Class
Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        '身份认证


        'Dim start_kaishi As New SplashScreen1


        'start_kaishi.Show()
        'System.Threading.Thread.Sleep(1500)
        'start_kaishi.Close()
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub Application_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Microsoft.Office.Interop.Excel.Range, ByRef Cancel As Boolean) Handles Application.SheetBeforeDoubleClick
        Dim yongh_dlg As New yongh
        Dim gongy_dlg As New gongy
        Dim tools_dlg As New tools
        Dim tools_c As New truntool

        Dim my_col As Integer '列
        Dim my_row As Integer '行

        Dim sb_ty As String  '加工中心或者数控车控制启动 

        Dim myExcel As Excel.Application = Nothing  '定义进程
        myExcel = GetObject(, "Excel.Application")  '打开已经打开的excel程序
        myExcel.Visible = True
        sb_ty = Mid(myExcel.ActiveSheet.Cells(6, 6).Value, 1, 1)

        my_row = myExcel.ActiveCell.Row
        my_col = myExcel.ActiveCell.Column
        GlobalData.rowActivite = my_row
        GlobalData.colActivite = my_col


        If myExcel.ActiveSheet.Cells(1, 1).Value = "pengjj" And myExcel.ActiveSheet.Cells(1, 7).Value = "p" And myExcel.ActiveSheet.Cells(1, 9).Value = "e" Then

            '用户身份验证
            If (OperationTools.CheckUser()) Then
            Else
                MsgBox("权限错误，请联系系统管理员",, "错误提示")
                Return
            End If

            'If DateAndTime.Year(DateAndTime.Today) = 2013 And DateAndTime.Month(DateAndTime.Today) <= 10 And DateAndTime.Month(DateAndTime.Today) >= 6 Then
            ' my_col_row = Target.Address
            '
            '  target_egproject(my_col_row)

            If my_col = 3 And my_row = 3 Then yongh_dlg.ShowDialog()
                If my_col = 3 And my_row > 8 Then gongy_dlg.ShowDialog()
                If my_col >= 4 And my_col <= 5 And my_row > 8 Then

                    '判断是否点击的是第一格：即判定A列中是否有值
                    If (myExcel.ActiveSheet.Cells(my_row, my_col - 3).Value = Nothing) Then
                        MsgBox("请从刀具第一行开始填写")
                        Return
                    Else
                        Try
                            'MsgBox(myExcel.ActiveSheet.Cells(my_row, my_col - 3).Value)
                            '记录刀号
                            GlobalData.toolNum = myExcel.ActiveSheet.Cells(my_row, my_col - 3).Value
                            '计算此刀号对应的表格目前包含的行数
                            'myExcel.Cells.Find(myExcel.ActiveSheet.Cells(my_row, my_col - 3).Value.ToString.Trim).Activate()
                            myExcel.ActiveSheet.Cells(my_row, my_col - 3).Activate()

                            GlobalData.rowsNow = Integer.Parse(myExcel.ActiveCell.MergeArea.Address.Substring(9)) - Integer.Parse(myExcel.ActiveCell.MergeArea.Address.Substring(3, 2)) + 1
                            'MsgBox(GlobalData.rowsNow)
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try


                    End If

                    If sb_ty = "M" Or sb_ty = "m" Then
                        tools_dlg.ShowDialog()
                    End If
                    If sb_ty = "C" Or sb_ty = "c" Then
                        tools_c.ShowDialog()
                    End If
                End If
                'End If

            End If

    End Sub
End Class

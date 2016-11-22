Imports System.Windows.Forms
Imports System.Windows.Forms.Keys
Imports Excel = Microsoft.Office.Interop.Excel

Public Class yongh


    Friend WithEvents Application As Microsoft.Office.Interop.Excel.Application
    Dim eg_project_sys As String '机床类型识别

    'Public app As Excel.Application
    'Public wb As Excel.Workbook
    'Public ws As Excel.Worksheet



    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        Dim MyDate
        Dim myExcel As Excel.Application = Nothing  '定义进程

        Dim i As Integer
        Dim str As String
        Dim fs, f
        Dim yongh As String
        Dim yhbm As String         '材料类型

        yongh = "0"

        Me.DialogResult = System.Windows.Forms.DialogResult.OK


        If Dir("c:\Program Files\方案文件库\设备类型.lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\设备类型.lib", 1, 0)
            i = 1


            Do While i = 1
                str = f.readline
                If str = "END" Then
                    i = -1
                Else

                    If ComboBox1.Text = Mid(str, 1, Len(str) - 1) Then yongh = Mid(str, Len(str), 1)

                End If

            Loop
            f.Close()

        End If

        If Dir("c:\Program Files\方案文件库\材料类型.lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\材料类型.lib", 1, 0)
            i = 1


            Do While i = 1
                str = f.readline
                If str = "END" Then
                    i = -1
                Else

                    If ComboBox3.Text = Mid(str, 4, Len(str) - 3) Then
                        yhbm = Mid(str, 1, 3)
                        i = -1
                    End If

                End If

            Loop
            f.Close()

        End If
        'MsgBox(eg_project_sys)

        MyDate = MonthCalendar1.TodayDate

        myExcel = GetObject(, "Excel.Application")  '打开已经打开的excel程序
        myExcel.Visible = True
        ' myExcel.ActiveSheet.Cells(1, 1).Value = "Hello"

        myExcel.ActiveSheet.Cells(4, 4).Value = TextBox2.Text '工件名称
        myExcel.ActiveSheet.Cells(4, 5).Value = "材料：" + TextBox3.Text '材料
        myExcel.ActiveSheet.Cells(4, 9).Value = TextBox4.Text '加工内容
        myExcel.ActiveSheet.Cells(4, 14).Value = ComboBox1.Text '机床类型
        myExcel.ActiveSheet.Cells(4, 16).Value = TextBox5.Text '机床型号
        myExcel.ActiveSheet.Cells(6, 6).Value = ComboBox1.Text.Substring(0, 1) '车床和铣床判定
        myExcel.ActiveSheet.Cells(6, 5).Value = ComboBox4.Text '机床接口
        myExcel.ActiveSheet.Cells(4, 20).Value = TextBox6.Text + "rpm" '转速
        myExcel.ActiveSheet.Cells(5, 16).Value = TextBox7.Text '平均换刀时间
        myExcel.ActiveSheet.Cells(3, 19).Value = MonthCalendar1.TodayDate.Date '日期
        'myExcel.ActiveSheet.Cells(3, 3).Value = TextBox1.Text
        myExcel.ActiveSheet.Cells(3, 3).Value = ComboBox2.Text '客户名称
        myExcel.ActiveSheet.Cells(3, 8).Value = TextBox1.Text '项目编码

        myExcel.ActiveSheet.Cells(3, 5).Value = ComboBox5.Text '拉钉型号
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub yongh_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim myExcel As Excel.Application = Nothing  '定义进程
        Dim i As Integer
        Dim str As String
        Dim fs, f
        Dim str_tmp As String

        ComboBox1.Items.Clear()
        ComboBox3.Items.Clear()

        If Dir("c:\Program Files\方案文件库\设备类型.lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\设备类型.lib", 1, 0)
            i = 1

            Do While i = 1
                str = f.readline
                If str = "END" Then
                    i = -1
                Else

                    'ComboBox1.Items.Add(Mid(str, 1, Len(str) - 1))
                    ComboBox1.Items.Add(Mid(str, 1, Len(str) - 1))

                End If

            Loop
            f.Close()
            If ComboBox1.Items.Count <> 0 Then ComboBox1.SelectedIndex = 0

        End If

        '材料类型
        str_tmp = "0000"
        If Dir("c:\Program Files\方案文件库\材料类型.lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\材料类型.lib", 1, 0)
            i = 1


            Do While i = 1
                str = f.readline
                If str = "END" Then
                    i = -1
                Else
                    ComboBox3.Items.Add(Mid(str, 4, Len(str) - 3))
                    '   If Len(Cells(5, 6).Value) = 4 And Mid(Cells(5, 6).Value, 2, 3) = Mid(str, 1, 3) And str_tmp = "0000" Then str_tmp = str
                End If
            Loop
            f.Close()


            If ComboBox3.Items.Count <> 0 Then ComboBox3.SelectedIndex = 0
        End If




        myExcel = GetObject(, "Excel.Application")  '打开已经打开的excel程序
        myExcel.Visible = True
        ' myExcel.ActiveSheet.Cells(1, 1).Value = "Hello"

        TextBox2.Text = myExcel.ActiveSheet.Cells(4, 4).Value    '工件名称

        TextBox4.Text = myExcel.ActiveSheet.Cells(4, 9).Value '加工内容

        ComboBox1.Text = myExcel.ActiveSheet.Cells(4, 14).Value '机床类型

        TextBox5.Text = myExcel.ActiveSheet.Cells(4, 16).Value '机床型号

        ComboBox4.Text = myExcel.ActiveSheet.Cells(6, 5).Value '机床接口

        TextBox7.Text = myExcel.ActiveSheet.Cells(5, 16).Value  '平均换刀时间

        TextBox1.Text = myExcel.ActiveSheet.Cells(3, 8).Value '经销商编码

        ComboBox5.Text = myExcel.ActiveSheet.Cells(3, 5).Value '拉钉型号

        TextBox3.Text = Mid(myExcel.ActiveSheet.Cells(4, 5).Value, 4, Len(myExcel.ActiveSheet.Cells(4, 5).Value) - 3)        '材料

        TextBox6.Text = Mid(myExcel.ActiveSheet.Cells(4, 20).Value, 1, Len(myExcel.ActiveSheet.Cells(4, 20).Value) - 3)   '转速
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim i As Integer
        Dim str As String
        Dim fs, f
        Dim str_temp1 As String

        ComboBox4.Items.Clear()


        If Dir("c:\Program Files\方案文件库\设备类型.lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\设备类型.lib", 1, 0)
            i = 1

            Do While i = 1
                str = f.readline
                If str = "END" Then
                    i = -1
                Else

                    If ComboBox1.Text = Mid(str, 1, Len(str) - 1) Then str_temp1 = Mid(str, Len(str), 1)

                End If

            Loop
            f.Close()


        End If

        'MsgBox(str_temp1)
        eg_project_sys = str_temp1

        If Dir("c:\Program Files\方案文件库\mection.dat", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\mection.dat", 1, 0)
            i = 1

            Do While i = 1
                str = f.readline
                If str = "END" Then
                    i = -1
                Else
                    If (Mid(str, 1, 1)) = str_temp1 Then
                        ComboBox4.Items.Add(Mid(str, 2, Len(str) - 4))
                    End If
                End If


            Loop
            f.Close()
            If ComboBox4.Items.Count <> 0 Then ComboBox4.SelectedIndex = 0

        End If


    End Sub






    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged

        Dim i As Integer
        Dim str As String
        Dim fs, f
        Dim str_temp1 As String

        ComboBox5.Items.Clear()

        If Dir("c:\Program Files\方案文件库\mection.dat", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\mection.dat", 1, 0)
            i = 1

            Do While i = 1
                str = f.readline
                If str = "END" Then
                    i = -1
                Else
                    If ComboBox4.Text = (Mid(str, 2, Len(str) - 4)) Then str_temp1 = Mid(str, Len(str) - 2, 3)
                End If
            Loop
            f.Close()
        End If


        If Dir("c:\Program Files\方案文件库\me_din.dat", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\me_din.dat", 1, 0)
            i = 1

            Do While i = 1
                str = f.readline
                If str = "END" Then
                    i = -1
                Else

                    If (Mid(str, 1, 3)) = str_temp1 Then
                        ComboBox5.Items.Add(Mid(str, 4, Len(str) - 3))
                    End If

                End If

            Loop
            f.Close()
            If ComboBox5.Items.Count <> 0 Then ComboBox5.SelectedIndex = 0 Else ComboBox5.Text = ""

        End If

    End Sub






    Private Sub ComboBox4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.TextChanged

        Dim i As Integer
        Dim str As String
        Dim fs, f
        Dim str_temp1 As String

        ComboBox5.Items.Clear()

        If Dir("c:\Program Files\方案文件库\mection.dat", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\mection.dat", 1, 0)
            i = 1

            Do While i = 1
                str = f.readline
                If str = "END" Then
                    i = -1
                Else

                    If ComboBox4.Text = (Mid(str, 2, Len(str) - 4)) Then str_temp1 = Mid(str, Len(str) - 2, 3)

                End If

            Loop
            f.Close()


        End If


        If Dir("c:\Program Files\方案文件库\me_din.dat", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\me_din.dat", 1, 0)
            i = 1

            Do While i = 1
                str = f.readline
                If str = "END" Then
                    i = -1
                Else

                    If (Mid(str, 1, 3)) = str_temp1 Then
                        ComboBox5.Items.Add(Mid(str, 4, Len(str) - 3))
                    End If

                End If

            Loop
            f.Close()
            If ComboBox5.Items.Count <> 0 Then ComboBox5.SelectedIndex = 0 Else ComboBox5.Text = ""

        End If

    End Sub

End Class

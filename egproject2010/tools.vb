Imports System.Windows.Forms
Imports System.Windows.Forms.Keys
Imports Excel = Microsoft.Office.Interop.Excel

Public Class tools

    Dim stxf1 As String '刀具型号文件开头第1行
    Dim stxf2 As String '刀具型号文件开头第2行  刀柄信息
    Dim stxf3 As String '刀具型号文件开头第3行  附件信息

    Dim tooling_insert_k As String '刀片型号
    Dim t_max As String '刀具直径最大
    Dim t_min As String '刀具直径最小
    Dim t_ren As String '刀具切削刃数
    Dim t_inf As String '刀具
    Dim t_Tap As String '刀具接口尺寸
    Dim t_Tap_Type_cb As String '刀柄接口类型
    Dim t_long As String '刀具长度
    Friend Shared t_yx_long As String '刀具有效刃长，丝锥表示柄部方尺寸
    Dim t_num_ck_cb As String '刀柄接口尺寸




    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        GlobalData.rowsToExcel = 0

        Dim my_col As Integer '列
        Dim my_row As Integer '行

        Dim i As Integer
        Dim stxf As String
        Dim fs, f

        Dim myExcel As Excel.Application = Nothing  '定义进程
        myExcel = GetObject(, "Excel.Application")  '打开已经打开的excel程序
        myExcel.Visible = True
        'sb_ty = Mid(myExcel.ActiveSheet.Cells(6, 6).Value, 1, 1)

        my_row = myExcel.ActiveCell.Row
        my_col = myExcel.ActiveCell.Column

        '''''''''''''''''''''''计算需要写入的行数''''''''''''''''''''''''''''''''''
        '刀具型号
        GlobalData.rowsToExcel += 1
        '刀片
        If Val(Mid(stxf1, 2, 1)) = 1 Then
            GlobalData.rowsToExcel += 1
        End If
        '工具名称和型号
        If ComboBox3.Text <> "" Then
            GlobalData.rowsToExcel += 1
        End If
        '切削刀具附件名称和型号
        If CheckBox4.CheckState = 1 Then
            GlobalData.rowsToExcel += 1
        End If
        '工具系统附件名称和型号
        If CheckBox2.CheckState = 1 Then
            GlobalData.rowsToExcel += 1
        End If
        '其他附件一名称和型号
        If CheckBox5.CheckState = 1 Then
            GlobalData.rowsToExcel += 1
        End If
        '其他附件二名称和型号
        If CheckBox6.CheckState = 1 Then
            GlobalData.rowsToExcel += 1
        End If
        '工具系统附件补充名称和型号
        If CheckBox7.CheckState = 1 Then
            GlobalData.rowsToExcel += 1
        End If
        '拉钉型号
        GlobalData.rowsToExcel += 1
        ''''''''''''''''''''''''''''''''''调整行数''''''''''''''''''''''''''''''''''
        OperationTools.AdjustRows(GlobalData.rowsNow, GlobalData.rowsToExcel, GlobalData.rowActivite, myExcel)
        ''''''''''''''''''''''''''''''''''清理数据''''''''''''''''''''''''''''''''''
        Dim tmpRange As Excel.Range
        If GlobalData.rowsToExcel <= 4 Then
            tmpRange = myExcel.Range(Cell1:=myExcel.Cells(GlobalData.rowActivite, GlobalData.colActivite),
                              Cell2:=myExcel.Cells(GlobalData.rowActivite + 3, GlobalData.colActivite + 20))
        Else
            tmpRange = myExcel.Range(Cell1:=myExcel.Cells(GlobalData.rowActivite, GlobalData.colActivite),
                              Cell2:=myExcel.Cells(GlobalData.rowActivite + GlobalData.rowsToExcel - 1, GlobalData.colActivite + 20))
        End If
        tmpRange.ClearContents()
        '清除批注
        Try
            tmpRange.ClearComments()
        Catch ex As Exception

        End Try

        ''''''''''''''''''''''''''''''''''''''''''填写数据''''''''''''''''''''''''''''''''''
        '填写切削刀具信息

        myExcel.ActiveSheet.Cells(my_row, 4).Value = ListBox2.Text

        If Val(Mid(stxf1, 1, 1)) = 1 Then

            If Mid(myExcel.ActiveSheet.Cells(6, 6).Value, 2, 1) = "P" Or Mid(myExcel.ActiveSheet.Cells(6, 6).Value, 2, 1) = "M" Then
                myExcel.ActiveSheet.Cells(my_row, 5).Value = ComboBox1.Text & " MG18"
            Else
                myExcel.ActiveSheet.Cells(my_row, 5).Value = ComboBox1.Text & " ZK10UF"

            End If

        Else
            myExcel.ActiveSheet.Cells(my_row, 5).Value = ComboBox1.Text
        End If

        myExcel.ActiveSheet.Cells(my_row, 6).Value = TextBox1.Text
        myExcel.ActiveSheet.Cells(my_row, 7).Value = zh_long_k()
        myExcel.ActiveSheet.Cells(my_row, 8).Value = TextBox4.Text
        myExcel.ActiveSheet.Cells(my_row, 9).Value = "=H" & my_row & "*1000/3.14/" & "F" & my_row
        myExcel.ActiveSheet.Cells(my_row, 10).Value = TextBox5.Text
        myExcel.ActiveSheet.Cells(my_row, 11).Value = "=J" & my_row & "*I" & my_row & "*" & TextBox3.Text
        myExcel.ActiveSheet.Cells(my_row, 13).Value = 0
        myExcel.ActiveSheet.Cells(my_row, 14).Value = 1
        myExcel.ActiveSheet.Cells(my_row, 15).Value = "=M" & my_row & "*N" & my_row & "/K" & my_row
        myExcel.ActiveSheet.Cells(my_row, 16).Value = "=P5/60"
        myExcel.ActiveSheet.Cells(my_row, 17).Value = "0"
        myExcel.ActiveSheet.Cells(my_row, 18).Value = "=O" & my_row & "+P" & my_row & "+Q" & my_row

        myExcel.ActiveSheet.Cells(my_row + 1, 4).Value = ""
        myExcel.ActiveSheet.Cells(my_row + 1, 5).Value = ""
        myExcel.ActiveSheet.Cells(my_row + 2, 4).Value = ""
        myExcel.ActiveSheet.Cells(my_row + 2, 5).Value = ""
        myExcel.ActiveSheet.Cells(my_row + 3, 4).Value = ""
        myExcel.ActiveSheet.Cells(my_row + 3, 5).Value = ""

        my_row = my_row + 1

        '刀片
        If Val(Mid(stxf1, 2, 1)) = 1 Then
            myExcel.ActiveSheet.Cells(my_row, 4).Value = "刀片"
            myExcel.ActiveSheet.Cells(my_row, 5).Value = ComboBox2.Text
            myExcel.Rows(my_row).Hidden = False
            my_row = my_row + 1
        End If

        '工具名称和型号
        If ComboBox3.Text <> "" Then
            myExcel.ActiveSheet.Cells(my_row, 4).Value = ComboBox3.Text
            myExcel.ActiveSheet.Cells(my_row, 5).Value = ComboBox4.Text
            myExcel.Rows(my_row).Hidden = False
            my_row = my_row + 1
        End If

        '切削刀具附件名称和型号
        If CheckBox4.CheckState = 1 Then
            myExcel.ActiveSheet.Cells(my_row, 4).Value = ComboBox8.Text
            myExcel.ActiveSheet.Cells(my_row, 5).Value = ComboBox7.Text
            myExcel.Rows(my_row).Hidden = False
            my_row = my_row + 1
        End If

        '工具系统附件名称和型号
        If CheckBox2.CheckState = 1 Then
            myExcel.ActiveSheet.Cells(my_row, 4).Value = ComboBox6.Text
            myExcel.ActiveSheet.Cells(my_row, 5).Value = ComboBox5.Text
            myExcel.Rows(my_row).Hidden = False
            my_row = my_row + 1
        End If

        '其他附件一名称和型号
        If CheckBox5.CheckState = 1 Then
            myExcel.ActiveSheet.Cells(my_row, 4).Value = ComboBox10.Text
            myExcel.ActiveSheet.Cells(my_row, 5).Value = ComboBox9.Text
            myExcel.Rows(my_row).Hidden = False
            my_row = my_row + 1
        End If

        '其他附件二名称和型号
        If CheckBox6.CheckState = 1 Then
            myExcel.ActiveSheet.Cells(my_row, 4).Value = ComboBox12.Text
            myExcel.ActiveSheet.Cells(my_row, 5).Value = ComboBox11.Text
            myExcel.Rows(my_row).Hidden = False
            my_row = my_row + 1
        End If

        '工具系统附件补充名称和型号
        If CheckBox7.CheckState = 1 Then
            myExcel.ActiveSheet.Cells(my_row, 4).Value = ComboBox14.Text
            myExcel.ActiveSheet.Cells(my_row, 5).Value = ComboBox13.Text
            myExcel.Rows(my_row).Hidden = False
            my_row = my_row + 1
        End If
        '填写拉钉型号
        myExcel.ActiveSheet.Cells(my_row, 4).Value = "拉钉"
        myExcel.ActiveSheet.Cells(my_row, 5).Value = "=E3"
        'my_row = my_row + 1

        If Dir("c:\Program Files\方案文件库\me_din.lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\me_din.lib", 1, 0)
            i = 1

            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else
                    ' MsgBox(Len(myExcel.ActiveSheet.Cells(3, 5).Value) + 1)
                    If myExcel.ActiveSheet.Cells(3, 5).Value = Mid(stxf, 1, Len(myExcel.ActiveSheet.Cells(3, 5).Value)) Then
                        myExcel.ActiveSheet.Cells(my_row, 4).Value = Mid(stxf, Len(myExcel.ActiveSheet.Cells(3, 5).Value) + 1, Len(stxf) - Len(myExcel.ActiveSheet.Cells(3, 5).Value))
                        myExcel.Rows(my_row).Hidden = False
                        i = -1
                    End If
                End If
            Loop
            f.Close()

        End If

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub tools_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer
        Dim stxf As String
        Dim fs, f

        Dim sb_ty As String  '加工中心或者数控车控制启动 

        Dim myExcel As Excel.Application = Nothing  '定义进程
        myExcel = GetObject(, "Excel.Application")  '打开已经打开的excel程序
        myExcel.Visible = True

        sb_ty = Mid(myExcel.ActiveSheet.Cells(6, 6).Value, 1, 1)


        If sb_ty = "M" Or sb_ty = "m" Then

            If Dir("c:\Program Files\方案文件库\刀具类型.lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\刀具类型.lib", 1, 0)
                i = 1

                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then i = -1 Else ListBox1.Items.Add(stxf)
                Loop
                f.Close()
                If ListBox1.Items.Count <> 0 Then ListBox1.SelectedIndex = 0
            End If

        End If

        If sb_ty = "C" Or sb_ty = "c" Then
            If Dir("c:\Program Files\方案文件库\刀具类型_车.lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\刀具类型_车.lib", 1, 0)
                i = 1

                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then i = -1 Else ListBox1.Items.Add(stxf)
                Loop
                f.Close()
                If ListBox1.Items.Count <> 0 Then ListBox1.SelectedIndex = 0
            End If

        End If

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

        Dim i As Integer
        Dim stxf As String
        Dim fs, f
        ListBox2.Items.Clear()

        If Dir("c:\Program Files\方案文件库\" & ListBox1.Text & ".lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ListBox1.Text & ".lib", 1, 0)
            i = 1
            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then i = -1 Else ListBox2.Items.Add(stxf)
            Loop
            f.Close()

            If ListBox2.Items.Count <> 0 Then ListBox2.SelectedIndex = 0
        End If

        TextBox1.Focus()
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox2.SelectedIndexChanged
        Dim i As Integer
        Dim stxf As String
        Dim fs, f



        Dim insert_temp_kz As String

        Dim myExcel As Excel.Application = Nothing  '定义进程
        myExcel = GetObject(, "Excel.Application")  '打开已经打开的excel程序
        myExcel.Visible = True
        insert_temp_kz = myExcel.ActiveSheet.Cells(6, 5).Value




        ComboBox1.Items.Clear()

        If Dir("c:\Program Files\方案文件库\图片\" & ListBox2.Text & ".jpg", vbNormal) <> "" Then
            PictureBox1.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & ListBox2.Text & ".jpg")
        Else
            PictureBox1.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片.jpg")
        End If

        '  PictureBox2.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片.jpg")


        If Dir("c:\Program Files\方案文件库\" & ListBox2.Text & ".lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ListBox2.Text & ".lib", 1, 0)

            stxf1 = f.readline
            Tools_QK(stxf1)

            stxf2 = f.readline
            stxf3 = f.readline




            i = 1
            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else

                    tooling_insert_k = Mid(zh_biam_f(stxf), 1, 4)
                    t_min = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 24, 6)
                    t_max = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 18, 6)
                    t_ren = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 27, 3)
                    t_inf = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 28, 1)
                    t_Tap = Mid(zh_biam_f(stxf), 11, 6)
                    t_long = Mid(zh_biam_f(stxf), 17, 6)
                    t_yx_long = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 6)
                    t_Tap_Type_cb = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 2, 3)
                    t_num_ck_cb = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 6)


                    If Mid(stxf1, 8, 1) = "1" Then

                        ' MsgBox(TOOLING_num(t_Tap_Type_cb, t_num_ck_cb) & " hg " & insert_temp_kz)
                        If TOOLING_num(t_Tap_Type_cb, t_num_ck_cb) = insert_temp_kz Then ComboBox1.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 29))


                    Else

                        ComboBox1.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 29))
                    End If

                    ' ComboBox1.Items.Add(stxf)
                End If
            Loop
            f.Close()

            If ComboBox1.Items.Count <> 0 Then ComboBox1.SelectedIndex = 0
        End If
        TextBox1.Focus()

    End Sub

    Public Sub Tools_QK(ByVal str As String)  '第一行解码
        ' Dim ii As Integer
        ' MsgBox(str)
        'For ii = 1 To Len(str) Step 1
        'MsgBox(Mid(str, ii, 1))
        'Next ii
        ComboBox1.Items.Clear()
        ComboBox2.Items.Clear()
        TextBox1.Text = "0"
        TextBox2.Text = "0"
        TextBox3.Text = "0"
        TextBox4.Text = "0"
        TextBox5.Text = "0"
        TextBox6.Text = "0"

        Label9.Text = ""
        Label10.Text = ""
        Label19.Text = ""


        ComboBox3.Items.Clear()
        ComboBox4.Items.Clear()
        ComboBox5.Items.Clear()
        ComboBox6.Items.Clear()
        ComboBox7.Items.Clear()
        ComboBox8.Items.Clear()
        ComboBox9.Items.Clear()
        ComboBox10.Items.Clear()


        If Mid(str, 2, 1) = "1" Then ComboBox2.Enabled = True Else ComboBox2.Enabled = False
        If Mid(str, 2, 1) = "1" Then CheckBox1.Enabled = True Else CheckBox1.Enabled = False

        If Mid(str, 3, 1) = "1" Then CheckBox4.Enabled = True Else CheckBox4.Enabled = False
        If Mid(str, 3, 1) = "1" Then ComboBox7.Enabled = True Else ComboBox7.Enabled = False
        If Mid(str, 3, 1) = "1" Then ComboBox8.Enabled = True Else ComboBox8.Enabled = False

        If Mid(str, 7, 1) = "1" Then ComboBox3.Enabled = True Else ComboBox3.Enabled = False
        If Mid(str, 7, 1) = "1" Then ComboBox4.Enabled = True Else ComboBox4.Enabled = False
        If Mid(str, 7, 1) = "1" Then ComboBox5.Enabled = True Else ComboBox5.Enabled = False
        If Mid(str, 7, 1) = "1" Then ComboBox6.Enabled = True Else ComboBox6.Enabled = False
        If Mid(str, 7, 1) = "1" Then CheckBox2.Enabled = True Else CheckBox2.Enabled = False
        If Mid(str, 7, 1) = "1" Then CheckBox3.Enabled = True Else CheckBox3.Enabled = False

    End Sub

    ''' <summary>
    ''' 对配置文件进行伪解密操作
    ''' </summary>
    ''' <param name="xd"></param>
    ''' <returns></returns>
    Friend Shared Function zh_biam_f(ByVal xd As String) As String
        Dim ii As Integer
        Dim stxf_d As String
        Dim d As String

        For ii = 1 To Len(xd) Step 1
            'stxf_w = stxf_w + zh_biam_f(Mid(stxf, ii, 1))

            d = Mid(xd, ii, 1)

            If d = "し" Then d = "0"
            If d = "さ" Then d = "1"
            If d = "ゑ" Then d = "2"
            If d = "と" Then d = "3"
            If d = "か" Then d = "4"
            If d = "ふ" Then d = "5"
            If d = "ィ" Then d = "6"
            If d = "サ" Then d = "7"
            If d = "ヱ" Then d = "8"
            If d = "ぃ" Then d = "9"

            If d = "こ" Then d = "A"
            If d = "ぅ" Then d = "B"
            If d = "せ" Then d = "C"
            If d = "ナ" Then d = "D"
            If d = "ヌ" Then d = "E"
            If d = "ハ" Then d = "F"
            If d = "ˇ" Then d = "G"
            If d = "ㄚ" Then d = "H"
            If d = "ㄥ" Then d = "I"
            If d = "ч" Then d = "J"
            If d = "ш" Then d = "K"
            If d = "ы" Then d = "L"
            If d = "β" Then d = "M"
            If d = "γ" Then d = "N"
            If d = "ω" Then d = "O"
            If d = "υ" Then d = "P"
            If d = "。" Then d = "Q"
            If d = "〖" Then d = "R"
            If d = "【" Then d = "S"
            If d = "＇" Then d = "T"
            If d = "』" Then d = "U"
            If d = "『" Then d = "V"
            If d = "〔" Then d = "W"
            If d = "〈" Then d = "X"
            If d = "《" Then d = "Y"
            If d = "｝" Then d = "Z"

            If d = "ニ" Then d = "a"
            If d = "ノ" Then d = "b"
            If d = "ホ" Then d = "c"
            If d = "ㄅ" Then d = "d"
            If d = "ㄣ" Then d = "e"
            If d = "ㄈ" Then d = "f"
            If d = "ㄌ" Then d = "g"
            If d = "а" Then d = "h"
            If d = "л" Then d = "i"
            If d = "ц" Then d = "j"
            If d = "δ" Then d = "k"
            If d = "ζ" Then d = "l"
            If d = "κ" Then d = "m"
            If d = "〗" Then d = "n"
            If d = "】" Then d = "o"
            If d = "〕" Then d = "p"
            If d = "〉" Then d = "q"
            If d = "》" Then d = "r"
            If d = "‖" Then d = "s"
            If d = "☆" Then d = "t"
            If d = "→" Then d = "u"
            If d = "♀" Then d = "v"
            If d = "□" Then d = "w"
            If d = "＠" Then d = "x"
            If d = "●" Then d = "y"
            If d = "◇" Then d = "z"

            If d = "★" Then d = "-"
            If d = "♂" Then d = "*"
            If d = "◎" Then d = "/"


            stxf_d = stxf_d + d
        Next ii

        xd = stxf_d
        Return xd
    End Function


    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        Dim i As Integer
        Dim stxf As String
        Dim fs, f

        Dim stri As Integer
        Dim strk As Integer




        If Dir("c:\Program Files\方案文件库\" & ListBox2.Text & "信息.lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ListBox2.Text & "信息.lib", 1, 0)

            i = 1

            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else

                    '   MsgBox(Mid(stxf, 1, Len(ComboBox1.Text)) & "afasfg  " & ComboBox1.Text)

                    If Mid(stxf, 1, Len(ComboBox1.Text)) = ComboBox1.Text Then
                        Label10.Text = Mid(stxf, Len(ComboBox1.Text) + 1, Len(stxf) - Len(ComboBox1.Text))
                        '   MsgBox(Mid(stxf, Len(ComboBox1.Text) + 1, Len(stxf) - Len(ComboBox1.Text)))
                        i = -1
                    Else
                        Label10.Text = ""
                    End If
                End If
            Loop
            f.Close()
        Else
            Label9.Text = ""
        End If


        If Dir("c:\Program Files\方案文件库\" & ListBox2.Text & ".lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ListBox2.Text & ".lib", 1, 0)
            i = 1

            stxf = f.readline
            stxf = f.readline
            stxf = f.readline

            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else

                    If Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 29) = ComboBox1.Text Then
                        tooling_insert_k = Mid(zh_biam_f(stxf), 1, 4)
                        t_min = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 24, 6)
                        t_max = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 18, 6)
                        t_ren = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 27, 3)
                        t_inf = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 28, 1)
                        t_Tap = Mid(zh_biam_f(stxf), 11, 6)
                        t_long = Mid(zh_biam_f(stxf), 17, 6)
                        t_yx_long = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 6)
                        t_Tap_Type_cb = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 2, 3)
                        t_num_ck_cb = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 6)



                        '  MsgBox(t_Tap_Type)
                        ' MsgBox(tooling_insert_k)
                        '  MsgBox(t_min)
                        ' MsgBox(t_max)
                        ' MsgBox(t_ren)
                        ' MsgBox(t_inf)
                        'MsgBox(t_long)
                        ' MsgBox(t_yx_long)


                        If Val(TextBox1.Text) >= Val(t_min) And Val(TextBox1.Text) <= Val(t_max) Then
                            TextBox1.Text = TextBox1.Text
                        Else
                            TextBox1.Text = Val(Mid(zh_biam_f(stxf), 5, 6))
                        End If

                        TextBox3.Text = Val(t_ren)
                        TextBox2.Text = Val(t_Tap)

                        '显示控制变量  k=0不显示  k=1刃数    k=2调节范围    k=3刃数和调节范围同时显示
                        'If Val(t_inf) = 0 Then Label9.Text = "D"
                        If Val(t_inf) = 1 Then Label9.Text = " 刃数：" & TextBox3.Text
                        If Val(t_inf) = 2 Then Label9.Text = "调节范围：" & CStr(Val(t_min)) & "～" & CStr(Val(t_max))
                        If Val(t_inf) = 3 Then Label9.Text = "刃数：" & TextBox3.Text & "      " & "调节范围：" & CStr(Val(t_min)) & "～" & CStr(Val(t_max))

                        i = -1



                    End If
                End If

            Loop
            f.Close()
        End If

        If Mid(stxf1, 3, 1) = "0" Then   '刀具附件开关

            ComboBox7.Enabled = False
            ComboBox8.Enabled = False
            CheckBox4.Enabled = False

            CheckBox4.CheckState = 0
            ComboBox8.Items.Clear()
            ComboBox7.Items.Clear()
        Else
            ComboBox7.Enabled = True
            ComboBox8.Enabled = True
            CheckBox4.Enabled = True

            CheckBox4.CheckState = 1
            ComboBox8.Items.Clear()
            ComboBox7.Items.Clear()

            Call TOOLING_Accessories()
        End If



        If Mid(stxf1, 2, 1) = "0" Then '刀片开关
            ComboBox2.Items.Clear()
            ComboBox2.Text = ""
            ComboBox2.Enabled = False
            CheckBox1.Enabled = False

        Else
            ComboBox2.Enabled = True
            CheckBox1.Enabled = True


            tooling_insert(tooling_insert_k)
        End If


        If Mid(stxf1, 7, 1) = "0" Then    '工具系统选项开关
            ComboBox3.Enabled = False
            ComboBox4.Enabled = False
            CheckBox3.Enabled = True

        Else
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            CheckBox3.Enabled = True


            ComboBox3.Items.Clear()
            If Dir("c:\Program Files\方案文件库\工具系统.lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")

                stri = 1
                strk = 1
                Do While stri <= Len(stxf2) / 4
                    f = fs.OpenTextFile("c:\Program Files\方案文件库\工具系统.lib", 1, 0)
                    i = 1
                    Do While i = 1
                        stxf = f.readline
                        If stxf = "END" Then
                            i = -1
                        Else
                            If Mid(stxf, 1, 4) = Mid(stxf2, strk, 4) Then ComboBox3.Items.Add(Mid(stxf, 5, Len(stxf) - 5))
                        End If
                    Loop
                    f.Close()

                    stri = stri + 1
                    strk = strk + 4
                Loop
                If ComboBox3.Items.Count <> 0 Then ComboBox3.SelectedIndex = 0
            End If

            'tooling_shank()
        End If


        Call date_canshu()
        TextBox1.Focus()

    End Sub

    Public Sub TOOLING_Accessories()

        Dim i As Integer
        Dim stri As Integer
        Dim strk As Integer

        Dim stxf As String
        Dim fs, f

        If Dir("c:\Program Files\方案文件库\刀具附件.lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")

            stri = 1
            strk = 1
            Do While stri <= Len(stxf3) / 4
                f = fs.OpenTextFile("c:\Program Files\方案文件库\刀具附件.lib", 1, 0)
                i = 1
                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else
                        If Mid(stxf, 1, 4) = Mid(stxf3, strk, 4) Then ComboBox8.Items.Add(Mid(stxf, 5, Len(stxf) - 4))
                    End If
                Loop
                f.Close()

                stri = stri + 1
                strk = strk + 4
            Loop
            If ComboBox8.Items.Count <> 0 Then ComboBox8.SelectedIndex = 0
        End If
    End Sub

    Public Sub tooling_insert(ByVal insert_km As String)
        ComboBox2.Items.Clear()
        '刀片型号
        Dim i As Integer
        Dim stxf As String
        Dim fs, f
        Dim insert_temp_kz As String


        Dim myExcel As Excel.Application = Nothing  '定义进程
        myExcel = GetObject(, "Excel.Application")  '打开已经打开的excel程序
        myExcel.Visible = True
        insert_temp_kz = myExcel.ActiveSheet.Cells(6, 6).Value

        If CheckBox1.CheckState = 1 Then
            If Dir("c:\Program Files\方案文件库\刀片型号.lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\刀片型号.lib", 1, 0)
                i = 1
                Do While i = 1
                    stxf = f.readline

                    If stxf = "END" Then
                        i = -1
                    Else

                        If insert_km = Mid(zh_biam_f(stxf), 1, 4) Then

                            ComboBox2.Items.Add(Mid(zh_biam_f(stxf), 18, Len(zh_biam_f(stxf)) - 16))
                        End If
                    End If
                Loop
                f.Close()

            End If
        End If


        If CheckBox1.CheckState = 0 Then

            If Dir("c:\Program Files\方案文件库\刀片型号.lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\刀片型号.lib", 1, 0)
                i = 1
                Do While i = 1
                    stxf = f.readline

                    If stxf = "END" Then
                        i = -1
                    Else
                        ' MsgBox(Mid(zh_biam_f(stxf), 18, Len(zh_biam_f(stxf)) - 16))
                        'myExcel.ActiveSheet.Cells(6, 6).Value
                        If insert_km = Mid(zh_biam_f(stxf), 1, 4) Then

                            If insert_temp_kz = "" Or Len(insert_temp_kz) <> 4 Then
                                ComboBox2.Items.Add(Mid(zh_biam_f(stxf), 18, Len(zh_biam_f(stxf)) - 16))
                            Else
                                ' MsgBox(myExcel.ActiveSheet.Cells(6, 6).Value)
                                If Mid(insert_temp_kz, 2, 1) = Mid(zh_biam_f(stxf), 5, 1) Or Mid(insert_temp_kz, 2, 1) = Mid(zh_biam_f(stxf), 6, 1) Or Mid(insert_temp_kz, 2, 1) = Mid(zh_biam_f(stxf), 7, 1) Or Mid(insert_temp_kz, 2, 1) = Mid(zh_biam_f(stxf), 8, 1) Or Mid(insert_temp_kz, 2, 1) = Mid(zh_biam_f(stxf), 9, 1) Or Mid(insert_temp_kz, 2, 1) = Mid(zh_biam_f(stxf), 10, 1) Then
                                    If Mid(stxf1, 4, 1) = Mid(zh_biam_f(stxf), 17, 1) Or Mid(zh_biam_f(stxf), 17, 1) = "M" Or Mid(zh_biam_f(stxf), 17, 1) = "m" Then ComboBox2.Items.Add(Mid(zh_biam_f(stxf), 18, Len(zh_biam_f(stxf)) - 16))
                                End If
                            End If
                        End If
                    End If
                Loop
                f.Close()

            End If
        End If

        If ComboBox2.Items.Count <> 0 Then ComboBox2.SelectedIndex = 0
    End Sub


    Public Sub tooling_shank()
        ComboBox4.Items.Clear()
        '工具系统型号
        Dim i As Integer
        Dim stxf As String
        Dim stxf_ass As String
        Dim fs, f
        Dim ass As Integer '附件

        Dim qh_km As String '附件主柄

        Dim tap_ck As String '夹持部位类型
        Dim tap_type_ck As String '柄部类型（3位）
        Dim tap_num_ck As String '柄部尺寸（6位）
        Dim tap_typenum_km As String '柄部完整型号

        Dim insert_temp_kz As String

        Dim myExcel As Excel.Application = Nothing  '定义进程
        myExcel = GetObject(, "Excel.Application")  '打开已经打开的excel程序
        myExcel.Visible = True
        insert_temp_kz = myExcel.ActiveSheet.Cells(6, 5).Value


        If Dir("c:\Program Files\方案文件库\工具系统.lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\工具系统.lib", 1, 0)
            i = 1

            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else

                    If Mid(stxf, 5, Len(stxf) - 5) = ComboBox3.Text Then qh_km = Mid(stxf, Len(stxf), 1)

                End If
            Loop
            f.Close()

        End If

        ' MsgBox(qh_km)
        ass = 0
        If qh_km = "0" Then

            CheckBox5.Enabled = False
            CheckBox5.CheckState = 0

            If CheckBox3.CheckState = 1 Then

                If Dir("c:\Program Files\方案文件库\" & ComboBox3.Text & ".lib", vbNormal) <> "" Then

                    fs = CreateObject("Scripting.FileSystemObject")
                    f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox3.Text & ".lib", 1, 0)
                    i = 1

                    stxf_ass = f.readline

                    If stxf_ass <> "0000" Then
                        ComboBox5.Enabled = True
                        ComboBox6.Enabled = True

                        CheckBox2.Enabled = True
                        CheckBox2.CheckState = 1
                        ass = 1
                    Else
                        ComboBox5.Enabled = False
                        ComboBox6.Enabled = False
                        CheckBox2.Enabled = False
                        CheckBox2.CheckState = 0
                        ass = 0
                    End If


                    Do While i = 1
                        stxf = f.readline
                        If stxf = "END" Then
                            i = -1
                        Else

                            tap_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 11, 3)
                            tap_type_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 3)
                            tap_num_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 5, 6)
                            tap_typenum_km = TOOLING_num(tap_type_ck, tap_num_ck)


                            If tap_typenum_km = insert_temp_kz Then
                                ComboBox4.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))
                            End If
                        End If
                    Loop
                    f.Close()

                End If
            End If


            If CheckBox3.CheckState = 0 Then

                If Dir("c:\Program Files\方案文件库\" & ComboBox3.Text & ".lib", vbNormal) <> "" Then

                    fs = CreateObject("Scripting.FileSystemObject")
                    f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox3.Text & ".lib", 1, 0)
                    i = 1

                    stxf_ass = f.readline

                    If stxf_ass <> "0000" Then
                        ComboBox5.Enabled = True
                        ComboBox6.Enabled = True

                        CheckBox2.Enabled = True
                        CheckBox2.CheckState = 1
                        ass = 1
                    Else
                        ComboBox5.Enabled = False
                        ComboBox6.Enabled = False

                        CheckBox2.Enabled = False
                        CheckBox2.CheckState = 0
                        ass = 0
                    End If


                    Do While i = 1
                        stxf = f.readline
                        If stxf = "END" Then
                            i = -1
                        Else
                            ' Dim Tap_head As String
                            '    MsgBox(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 11, 3))
                            '    MsgBox(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 3))
                            '    MsgBox(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 5, 6))
                            '    MsgBox(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))


                            tap_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 11, 3)
                            tap_type_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 3)
                            tap_num_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 5, 6)
                            tap_typenum_km = TOOLING_num(tap_type_ck, tap_num_ck)


                            '  Dim t_Tap As String '刀具接口尺寸
                            '  Dim t_Tap_Type As String '接口类型ComboBox4.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))

                            'MsgBox(t_Tap_Type & "DGD" & tap_ck)

                            If t_Tap_Type_cb = tap_ck Then

                                If Val(TextBox2.Text) >= Val(Mid(zh_biam_f(stxf), 5, 6)) And Val(TextBox2.Text) <= Val(Mid(zh_biam_f(stxf), 11, 6)) And tap_typenum_km = insert_temp_kz Then

                                    ComboBox4.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))
                                End If

                            End If


                        End If
                    Loop
                    f.Close()

                End If
            End If
            If ComboBox4.Items.Count <> 0 Then ComboBox4.SelectedIndex = 0

        End If

        If qh_km <> "0" Then

            CheckBox5.Enabled = True
            CheckBox5.CheckState = 1
            ComboBox10.Items.Clear()

            If CheckBox5.CheckState = 1 Then
                ComboBox10.Enabled = True
                ComboBox9.Enabled = True

                If qh_km = "1" Then ComboBox10.Items.Add("自动换刀工具锥柄")
                If qh_km = "2" Then ComboBox10.Items.Add("有扁尾莫氏圆锥孔刀柄")
                If qh_km = "3" Then ComboBox10.Items.Add("可调刀柄主柄")
                If ComboBox10.Items.Count <> 0 Then ComboBox10.SelectedIndex = 0
            Else
                ComboBox10.Enabled = False
                ComboBox9.Enabled = False
                ComboBox10.Items.Clear()

            End If


            If CheckBox3.CheckState = 1 Then

                If Dir("c:\Program Files\方案文件库\" & ComboBox3.Text & ".lib", vbNormal) <> "" Then

                    fs = CreateObject("Scripting.FileSystemObject")
                    f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox3.Text & ".lib", 1, 0)
                    i = 1

                    stxf_ass = f.readline

                    If stxf_ass <> "0000" Then
                        ComboBox5.Enabled = True
                        ComboBox6.Enabled = True

                        CheckBox2.Enabled = True
                        CheckBox2.CheckState = 1
                        ass = 1
                    Else
                        ComboBox5.Enabled = False
                        ComboBox6.Enabled = False
                        CheckBox2.Enabled = False
                        CheckBox2.CheckState = 0
                        ass = 0
                    End If


                    Do While i = 1
                        stxf = f.readline
                        If stxf = "END" Then
                            i = -1
                        Else

                            tap_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 11, 3)
                            tap_type_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 3)
                            tap_num_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 5, 6)
                            tap_typenum_km = TOOLING_num(tap_type_ck, tap_num_ck)

                            ComboBox4.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))

                        End If
                    Loop
                    f.Close()

                End If
            End If


            If CheckBox3.CheckState = 0 Then

                If Dir("c:\Program Files\方案文件库\" & ComboBox3.Text & ".lib", vbNormal) <> "" Then

                    fs = CreateObject("Scripting.FileSystemObject")
                    f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox3.Text & ".lib", 1, 0)
                    i = 1

                    stxf_ass = f.readline

                    If stxf_ass <> "0000" Then
                        ComboBox5.Enabled = True
                        ComboBox6.Enabled = True

                        CheckBox2.Enabled = True
                        CheckBox2.CheckState = 1
                        ass = 1
                    Else
                        ComboBox5.Enabled = False
                        ComboBox6.Enabled = False

                        CheckBox2.Enabled = False
                        CheckBox2.CheckState = 0
                        ass = 0
                    End If


                    Do While i = 1
                        stxf = f.readline
                        If stxf = "END" Then
                            i = -1
                        Else
                            ' Dim Tap_head As String
                            '    MsgBox(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 11, 3))
                            '    MsgBox(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 3))
                            '    MsgBox(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 5, 6))
                            '    MsgBox(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))


                            tap_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 11, 3)
                            tap_type_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 3)
                            tap_num_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 5, 6)
                            tap_typenum_km = TOOLING_num(tap_type_ck, tap_num_ck)


                            '  Dim t_Tap As String '刀具接口尺寸
                            '  Dim t_Tap_Type As String '接口类型ComboBox4.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))

                            'MsgBox(t_Tap_Type & "DGD" & tap_ck)

                            If t_Tap_Type_cb = tap_ck Then

                                If Val(TextBox2.Text) >= Val(Mid(zh_biam_f(stxf), 5, 6)) And Val(TextBox2.Text) <= Val(Mid(zh_biam_f(stxf), 11, 6)) Then

                                    ComboBox4.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))
                                End If

                            End If


                        End If
                    Loop
                    f.Close()

                End If
            End If
            If ComboBox4.Items.Count <> 0 Then ComboBox4.SelectedIndex = 0


        End If

    End Sub


    ''' <summary>
    ''' 通过选择工具名称绑定工具型号至下拉框
    ''' </summary>
    ''' <param name="gjmc"></param>
    ''' <param name="gjxh"></param>
    Public Sub tooling_shank(gjmc As String, gjxh As ComboBox)
        gjxh.Items.Clear()
        '工具系统型号
        Dim i As Integer
        Dim stxf As String
        Dim stxf_ass As String
        Dim fs, f
        Dim ass As Integer '附件

        Dim qh_km As String '附件主柄

        Dim tap_ck As String '夹持部位类型
        Dim tap_type_ck As String '柄部类型（3位）
        Dim tap_num_ck As String '柄部尺寸（6位）
        Dim tap_typenum_km As String '柄部完整型号

        Dim insert_temp_kz As String

        Dim myExcel As Excel.Application = Nothing  '定义进程
        myExcel = GetObject(, "Excel.Application")  '打开已经打开的excel程序
        myExcel.Visible = True
        insert_temp_kz = myExcel.ActiveSheet.Cells(6, 5).Value


        If Dir("c:\Program Files\方案文件库\工具系统.lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\工具系统.lib", 1, 0)
            i = 1

            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else

                    If Mid(stxf, 5, Len(stxf) - 5) = gjmc Then qh_km = Mid(stxf, Len(stxf), 1)

                End If
            Loop
            f.Close()

        End If

        ' MsgBox(qh_km)
        ass = 0
        If qh_km = "0" Then

            CheckBox5.Enabled = False
            CheckBox5.CheckState = 0

            If CheckBox3.CheckState = 1 Then

                If Dir("c:\Program Files\方案文件库\" & gjmc & ".lib", vbNormal) <> "" Then

                    fs = CreateObject("Scripting.FileSystemObject")
                    f = fs.OpenTextFile("c:\Program Files\方案文件库\" & gjmc & ".lib", 1, 0)
                    i = 1

                    stxf_ass = f.readline

                    If stxf_ass <> "0000" Then
                        ComboBox5.Enabled = True
                        ComboBox6.Enabled = True

                        CheckBox2.Enabled = True
                        CheckBox2.CheckState = 1
                        ass = 1
                    Else
                        ComboBox5.Enabled = False
                        ComboBox6.Enabled = False
                        CheckBox2.Enabled = False
                        CheckBox2.CheckState = 0
                        ass = 0
                    End If


                    Do While i = 1
                        stxf = f.readline
                        If stxf = "END" Then
                            i = -1
                        Else

                            tap_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 11, 3)
                            tap_type_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 3)
                            tap_num_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 5, 6)
                            tap_typenum_km = TOOLING_num(tap_type_ck, tap_num_ck)


                            If tap_typenum_km = insert_temp_kz Then
                                gjxh.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))
                            End If
                        End If
                    Loop
                    f.Close()

                End If
            End If


            If CheckBox3.CheckState = 0 Then

                If Dir("c:\Program Files\方案文件库\" & gjmc & ".lib", vbNormal) <> "" Then

                    fs = CreateObject("Scripting.FileSystemObject")
                    f = fs.OpenTextFile("c:\Program Files\方案文件库\" & gjmc & ".lib", 1, 0)
                    i = 1

                    stxf_ass = f.readline

                    If stxf_ass <> "0000" Then
                        ComboBox5.Enabled = True
                        ComboBox6.Enabled = True

                        CheckBox2.Enabled = True
                        CheckBox2.CheckState = 1
                        ass = 1
                    Else
                        ComboBox5.Enabled = False
                        ComboBox6.Enabled = False

                        CheckBox2.Enabled = False
                        CheckBox2.CheckState = 0
                        ass = 0
                    End If


                    Do While i = 1
                        stxf = f.readline
                        If stxf = "END" Then
                            i = -1
                        Else
                            ' Dim Tap_head As String
                            '    MsgBox(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 11, 3))
                            '    MsgBox(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 3))
                            '    MsgBox(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 5, 6))
                            '    MsgBox(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))


                            tap_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 11, 3)
                            tap_type_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 3)
                            tap_num_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 5, 6)
                            tap_typenum_km = TOOLING_num(tap_type_ck, tap_num_ck)


                            '  Dim t_Tap As String '刀具接口尺寸
                            '  Dim t_Tap_Type As String '接口类型gjxh.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))

                            'MsgBox(t_Tap_Type & "DGD" & tap_ck)

                            If t_Tap_Type_cb = tap_ck Then

                                If Val(TextBox14.Text) >= Val(Mid(zh_biam_f(stxf), 5, 6)) And Val(TextBox14.Text) <= Val(Mid(zh_biam_f(stxf), 11, 6)) And tap_typenum_km = insert_temp_kz Then

                                    gjxh.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))
                                ElseIf TextBox14.Text = "" And tap_typenum_km = insert_temp_kz Then
                                    gjxh.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))
                                End If

                            End If


                        End If
                    Loop
                    f.Close()

                End If
            End If
            If gjxh.Items.Count <> 0 Then
                gjxh.SelectedIndex = 0
            Else
                gjxh.Items.Add("无匹配项")
                gjxh.Text = ""
            End If

        End If

        If qh_km <> "0" Then

            CheckBox5.Enabled = True
            CheckBox5.CheckState = 1
            ComboBox10.Items.Clear()

            If CheckBox5.CheckState = 1 Then
                ComboBox10.Enabled = True
                ComboBox9.Enabled = True

                If qh_km = "1" Then ComboBox10.Items.Add("自动换刀工具锥柄")
                If qh_km = "2" Then ComboBox10.Items.Add("有扁尾莫氏圆锥孔刀柄")
                If qh_km = "3" Then ComboBox10.Items.Add("可调刀柄主柄")
                If ComboBox10.Items.Count <> 0 Then ComboBox10.SelectedIndex = 0
            Else
                ComboBox10.Enabled = False
                ComboBox9.Enabled = False
                ComboBox10.Items.Clear()

            End If


            If CheckBox3.CheckState = 1 Then

                If Dir("c:\Program Files\方案文件库\" & gjmc & ".lib", vbNormal) <> "" Then

                    fs = CreateObject("Scripting.FileSystemObject")
                    f = fs.OpenTextFile("c:\Program Files\方案文件库\" & gjmc & ".lib", 1, 0)
                    i = 1

                    stxf_ass = f.readline

                    If stxf_ass <> "0000" Then
                        ComboBox5.Enabled = True
                        ComboBox6.Enabled = True

                        CheckBox2.Enabled = True
                        CheckBox2.CheckState = 1
                        ass = 1
                    Else
                        ComboBox5.Enabled = False
                        ComboBox6.Enabled = False
                        CheckBox2.Enabled = False
                        CheckBox2.CheckState = 0
                        ass = 0
                    End If


                    Do While i = 1
                        stxf = f.readline
                        If stxf = "END" Then
                            i = -1
                        Else

                            tap_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 11, 3)
                            tap_type_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 3)
                            tap_num_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 5, 6)
                            tap_typenum_km = TOOLING_num(tap_type_ck, tap_num_ck)

                            gjxh.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))

                        End If
                    Loop
                    f.Close()

                End If
            End If


            If CheckBox3.CheckState = 0 Then

                If Dir("c:\Program Files\方案文件库\" & gjmc & ".lib", vbNormal) <> "" Then

                    fs = CreateObject("Scripting.FileSystemObject")
                    f = fs.OpenTextFile("c:\Program Files\方案文件库\" & gjmc & ".lib", 1, 0)
                    i = 1

                    stxf_ass = f.readline

                    If stxf_ass <> "0000" Then
                        ComboBox5.Enabled = True
                        ComboBox6.Enabled = True

                        CheckBox2.Enabled = True
                        CheckBox2.CheckState = 1
                        ass = 1
                    Else
                        ComboBox5.Enabled = False
                        ComboBox6.Enabled = False

                        CheckBox2.Enabled = False
                        CheckBox2.CheckState = 0
                        ass = 0
                    End If


                    Do While i = 1
                        stxf = f.readline
                        If stxf = "END" Then
                            i = -1
                        Else
                            ' Dim Tap_head As String
                            '    MsgBox(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 11, 3))
                            '    MsgBox(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 3))
                            '    MsgBox(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 5, 6))
                            '    MsgBox(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))


                            tap_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 11, 3)
                            tap_type_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 3)
                            tap_num_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 5, 6)
                            tap_typenum_km = TOOLING_num(tap_type_ck, tap_num_ck)


                            '  Dim t_Tap As String '刀具接口尺寸
                            '  Dim t_Tap_Type As String '接口类型gjxh.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))

                            'MsgBox(t_Tap_Type & "DGD" & tap_ck)

                            If t_Tap_Type_cb = tap_ck Then

                                If Val(TextBox14.Text) >= Val(Mid(zh_biam_f(stxf), 5, 6)) And Val(TextBox14.Text) <= Val(Mid(zh_biam_f(stxf), 11, 6)) Then

                                    gjxh.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))
                                ElseIf TextBox14.Text = "" And tap_typenum_km = insert_temp_kz Then
                                    gjxh.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))
                                End If

                            End If


                        End If
                    Loop
                    f.Close()

                End If
            End If
            If gjxh.Items.Count <> 0 Then
                gjxh.SelectedIndex = 0
            Else
                gjxh.Items.Add("无匹配项")
                gjxh.Text = ""
            End If


        End If

    End Sub

    Public Function TOOLING_num(ByVal tap_type_ckls As String, ByVal tap_num_ckls As String) As String
        Dim sxtdf As Integer
        Dim stxf_ck1 As String

        sxtdf = Val(tap_type_ckls)

        Select Case sxtdf
            Case 1
                stxf_ck1 = "BT" & Val(tap_num_ckls)
            Case 2
                stxf_ck1 = "JT" & Val(tap_num_ckls)
            Case 3
                stxf_ck1 = "CT" & Val(tap_num_ckls)
            Case 4
                stxf_ck1 = "ST" & Val(tap_num_ckls)
            Case 5
                stxf_ck1 = "HSK" & Val(tap_num_ckls) & "A"
            Case 6
                stxf_ck1 = "HSK" & Val(tap_num_ckls) & "B"
            Case 7
                stxf_ck1 = "HSK" & Val(tap_num_ckls) & "C"
            Case 8
                stxf_ck1 = "HSK" & Val(tap_num_ckls) & "D"
            Case 9
                stxf_ck1 = "HSK" & Val(tap_num_ckls) & "E"
            Case 10
                stxf_ck1 = "HSK" & Val(tap_num_ckls) & "F"
            Case 11
                stxf_ck1 = "M" & Val(tap_num_ckls)
            Case 12
                stxf_ck1 = "ME" & Val(tap_num_ckls)
            Case 13
                stxf_ck1 = "C" & Val(tap_num_ckls)
            Case 14
                stxf_ck1 = "XP" & Val(tap_num_ckls)
            Case 15
                stxf_ck1 = "XPD" & Val(tap_num_ckls)
            Case 16
                stxf_ck1 = "BB" & Val(tap_num_ckls)
            Case 17
                stxf_ck1 = "21CD" & Val(tap_num_ckls)
            Case 18
                stxf_ck1 = "BT" & Val(tap_num_ckls) & "B"
            Case 19
                stxf_ck1 = "JT" & Val(tap_num_ckls) & "B"
            Case 20
                stxf_ck1 = "CT" & Val(tap_num_ckls) & "B"
            Case 21
                stxf_ck1 = "ER" & Val(tap_num_ckls)
            Case 22
                stxf_ck1 = "DC" & Val(tap_num_ckls)
            Case 23
                stxf_ck1 = "XM" & Val(tap_num_ckls)
            Case 24
                stxf_ck1 = "TDC" & Val(tap_num_ckls)
            Case 25
                stxf_ck1 = "XS" & Val(tap_num_ckls)
        End Select


        tap_type_ckls = stxf_ck1

        Return tap_type_ckls
    End Function

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        tooling_shank()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        tooling_insert(tooling_insert_k)
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim i As Integer
        Dim stxf21 As String
        Dim stxf As String
        Dim fs, f
        Dim skm As Integer
        Dim insert_temp_kz As String

        If Dir("c:\Program Files\方案文件库\" & ListBox2.Text & ".lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ListBox2.Text & ".lib", 1, 0)
            i = 1

            stxf21 = f.readline
            stxf = f.readline
            stxf = f.readline

            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else


                    If ComboBox1.Text = Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 29) And Val(TextBox1.Text) >= Val(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 24, 6)) And Val(TextBox1.Text) <= Val(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 18, 6)) Then
                        ComboBox1.Text = ComboBox1.Text
                        skm = 1
                        i = -1
                    Else

                    End If

                End If

            Loop
            f.Close()
        End If


        If skm = 0 Then

            Dim myExcel As Excel.Application = Nothing  '定义进程
            myExcel = GetObject(, "Excel.Application")  '打开已经打开的excel程序
            myExcel.Visible = True
            insert_temp_kz = myExcel.ActiveSheet.Cells(6, 5).Value

            If Dir("c:\Program Files\方案文件库\" & ListBox2.Text & ".lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ListBox2.Text & ".lib", 1, 0)
                i = 1

                stxf21 = f.readline
                stxf = f.readline
                stxf = f.readline

                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else




                        If Mid(stxf1, 7, 1) = "0" And Mid(stxf1, 8, 1) <> "0" Then

                            If insert_temp_kz = TOOLING_num(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 2, 3), Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 6)) And Val(TextBox1.Text) >= Val(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 24, 6)) And Val(TextBox1.Text) <= Val(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 18, 6)) Then

                                ComboBox1.Text = Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 29)
                                i = -1
                            Else
                                ComboBox1.Text = "没有该数据,请核对!"
                                TextBox2.Text = "无"
                            End If

                        End If

                        If Mid(stxf1, 7, 1) <> "0" And Mid(stxf1, 8, 1) = "0" Then

                            If Val(TextBox1.Text) >= Val(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 24, 6)) And Val(TextBox1.Text) <= Val(Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 18, 6)) Then
                                ComboBox1.Text = Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 29)
                                i = -1
                            Else

                                ComboBox1.Text = "没有该数据,请核对!"
                                TextBox2.Text = "无"
                            End If
                        End If




                    End If

                Loop
                f.Close()
            End If

        End If

        TextBox1.Focus()
    End Sub


    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.CheckState = 1 Then
            ComboBox5.Enabled = True
            ComboBox6.Enabled = True

        Else
            ComboBox5.Enabled = False
            ComboBox6.Enabled = False
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.CheckState = 0 Then '刀具附件开关

            ComboBox8.Enabled = False
            ComboBox7.Enabled = False

        Else
            ComboBox8.Enabled = True
            ComboBox7.Enabled = True

        End If

    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox8.SelectedIndexChanged

        Dim i As Integer
        Dim stxf As String
        Dim fs, f
        ' Dim tooling_insert_k As String '刀片型号

        ComboBox7.Items.Clear()


        If Dir("c:\Program Files\方案文件库\" & ComboBox8.Text & ".lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox8.Text & ".lib", 1, 0)

            i = 1

            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else


                    If Val(Mid(zh_biam_f(stxf), 1, 4)) = Val(TextBox2.Text) Then
                        ComboBox7.Items.Add(Mid(zh_biam_f(stxf), 9, Len(zh_biam_f(stxf)) - 8))
                    End If

                    If Val(Mid(zh_biam_f(stxf), 5, 4)) = Val(tooling_insert_k) And Mid(zh_biam_f(stxf), 3, 2) = "XG" Then
                        ComboBox7.Items.Add(Mid(zh_biam_f(stxf), 9, Len(zh_biam_f(stxf)) - 8))
                    End If




                End If
            Loop


        End If

        If ComboBox7.Items.Count <> 0 Then ComboBox7.SelectedIndex = 0

        TextBox1.Focus()
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox7.SelectedIndexChanged
        Dim i As Integer
        Dim stxf As String
        Dim fs, f
        If Dir("c:\Program Files\方案文件库\" & ComboBox8.Text & "信息.lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox8.Text & "信息.lib", 1, 0)

            i = 1

            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else

                    If Mid(stxf, 1, Len(ComboBox7.Text)) = ComboBox7.Text Then
                        Label18.Text = Mid(stxf, Len(ComboBox7.Text) + 1, Len(stxf) - Len(ComboBox7.Text))
                        i = -1
                    Else
                        Label18.Text = ""
                    End If
                End If
            Loop
        Else
            Label18.Text = "信息："
        End If
        TextBox1.Focus()

    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged

        Dim i As Integer
        Dim stxf As String
        Dim fs, f

        Dim qh_km As String '附件主柄

        If Dir("c:\Program Files\方案文件库\工具系统.lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\工具系统.lib", 1, 0)
            i = 1

            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else

                    If Mid(stxf, 5, Len(stxf) - 5) = ComboBox3.Text Then qh_km = Mid(stxf, Len(stxf), 1)

                End If
            Loop
            f.Close()

        End If


        ComboBox9.Items.Clear()

        If CheckBox5.CheckState = 1 Then
            ComboBox10.Enabled = True
            ComboBox9.Enabled = True

            If qh_km = "1" Then ComboBox10.Items.Add("自动换刀工具锥柄")
            If qh_km = "2" Then ComboBox10.Items.Add("有扁尾莫氏圆锥孔刀柄")
            If qh_km = "3" Then ComboBox10.Items.Add("可调刀柄主柄")
            If ComboBox10.Items.Count <> 0 Then ComboBox10.SelectedIndex = 0
        Else
            ComboBox10.Enabled = False
            ComboBox9.Enabled = False
            ComboBox10.Items.Clear()
            ComboBox9.Items.Clear()
        End If
    End Sub


    Public Function qh_km_kz() As String
        Dim i As Integer
        Dim stxf As String

        Dim fs, f

        Dim qh_km1 As String

        qh_km1 = "0"

        If Dir("c:\Program Files\方案文件库\工具系统.lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\工具系统.lib", 1, 0)
            i = 1

            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else

                    If Mid(stxf, 5, Len(stxf) - 5) = ComboBox3.Text Then qh_km1 = Mid(stxf, Len(stxf), 1)

                End If
            Loop
            f.Close()

        End If
        Return qh_km1
    End Function


    Private Sub ComboBox10_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox10.SelectedIndexChanged
        tooling_fjl()
    End Sub

    Public Sub tooling_fjl()
        Dim i As Integer
        Dim stxf As String
        Dim fs, f
        Dim insert_temp_kz As String

        Dim tap_type_ck As String '柄部类型（3位）
        Dim tap_ck As String '夹持部位类型（3位）
        Dim tap_num_ck As String '柄部尺寸（6位）
        Dim tap_typenum_km As String '柄部完整型号

        Dim tap_type_ck1 As String '柄部类型（3位）
        Dim tap_ck1 As String '夹持部位类型（3位）
        Dim tap_num_ck1 As String '柄部尺寸（6位）

        ComboBox9.Items.Clear()

        Dim myExcel As Excel.Application = Nothing  '定义进程
        myExcel = GetObject(, "Excel.Application")  '打开已经打开的excel程序
        myExcel.Visible = True
        insert_temp_kz = myExcel.ActiveSheet.Cells(6, 5).Value


        If Dir("c:\Program Files\方案文件库\" & ComboBox3.Text & ".lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox3.Text & ".lib", 1, 0)
            i = 1

            stxf = f.readline

            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else

                    If ComboBox4.Text = Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12) Then
                        tap_ck1 = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 11, 3)
                        tap_type_ck1 = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 3)
                        tap_num_ck1 = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 5, 6)

                    End If
                End If
            Loop
            f.Close()

        End If


        If Dir("c:\Program Files\方案文件库\" & ComboBox10.Text & ".lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox10.Text & ".lib", 1, 0)

            i = 1
            stxf = f.readline
            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else
                    tap_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 11, 3)
                    tap_type_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 8, 3)
                    tap_num_ck = Mid(zh_biam_f(stxf), Len(zh_biam_f(stxf)) - 5, 6)
                    tap_typenum_km = TOOLING_num(tap_type_ck, tap_num_ck)


                    If tap_typenum_km = insert_temp_kz Then
                        If Val(tap_num_ck1) >= Val(Mid(zh_biam_f(stxf), 5, 6)) And Val(tap_num_ck1) <= Val(Mid(zh_biam_f(stxf), 11, 6)) Then
                            ComboBox9.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))
                        End If
                    End If
                End If
            Loop
            f.Close()

        End If


        If ComboBox9.Items.Count <> 0 Then ComboBox9.SelectedIndex = 0

    End Sub


    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        Dim fs, f
        Dim i As Integer
        Dim stxf As String

        If CheckBox6.CheckState = 1 Then
            ComboBox12.Enabled = True
            ComboBox11.Enabled = True

            '附件名称
            If Dir("c:\Program Files\方案文件库\附件名称.lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\附件名称.lib", 1, 0)
                i = 1
                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else
                        ComboBox12.Items.Add(Mid(stxf, 2, Len(stxf) - 1))

                    End If
                Loop
                f.Close()
                If ComboBox12.Items.Count <> 0 Then ComboBox12.SelectedIndex = 0
            End If
        Else
            ComboBox12.Enabled = False
            ComboBox11.Enabled = False

            ComboBox12.Items.Clear()
            ComboBox11.Items.Clear()
        End If
    End Sub

    Private Sub ComboBox12_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox12.SelectedIndexChanged
        Dim i As Integer
        Dim stxf As String
        Dim fs, f
        Dim km As String

        ComboBox11.Items.Clear()

        If Dir("c:\Program Files\方案文件库\附件名称.lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\附件名称.lib", 1, 0)
            i = 1

            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else

                    If Mid(stxf, 2, Len(stxf) - 1) = ComboBox12.Text Then km = Mid(stxf, 1, 1)

                End If
            Loop
            f.Close()

        End If


        If km = "0" Then

            If Dir("c:\Program Files\方案文件库\" & ComboBox12.Text & ".lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox12.Text & ".lib", 1, 0)
                i = 1
                stxf = f.readline
                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else
                        ComboBox11.Items.Add(Mid(zh_biam_f(stxf), 17, Len(stxf) - 16))
                    End If
                Loop
                f.Close()

            End If
            If ComboBox11.Items.Count <> 0 Then ComboBox11.SelectedIndex = 0
        End If

        If km = "1" Then

            If Dir("c:\Program Files\方案文件库\" & ComboBox12.Text & ".lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox12.Text & ".lib", 1, 0)
                i = 1
                stxf = f.readline
                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else
                        ComboBox11.Items.Add(Mid(zh_biam_f(stxf), 23, Len(stxf) - 22))
                    End If
                Loop
                f.Close()

            End If
            If ComboBox11.Items.Count <> 0 Then ComboBox11.SelectedIndex = 0
        End If

    End Sub

    Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged
        Dim fs, f
        Dim i As Integer
        Dim stxf As String

        If CheckBox7.CheckState = 1 Then
            ComboBox13.Enabled = True
            ComboBox14.Enabled = True

            '附件名称
            If Dir("c:\Program Files\方案文件库\附件名称.lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\附件名称.lib", 1, 0)
                i = 1
                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else
                        ComboBox14.Items.Add(Mid(stxf, 2, Len(stxf) - 1))

                    End If
                Loop
                f.Close()
                If ComboBox14.Items.Count <> 0 Then ComboBox14.SelectedIndex = 0
            End If
        Else
            ComboBox14.Enabled = False
            ComboBox13.Enabled = False

            ComboBox14.Items.Clear()
            ComboBox13.Items.Clear()
        End If
    End Sub

    Private Sub ComboBox14_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox14.SelectedIndexChanged
        Dim i As Integer
        Dim stxf As String
        Dim fs, f
        Dim km As String

        ComboBox13.Items.Clear()

        If Dir("c:\Program Files\方案文件库\附件名称.lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\附件名称.lib", 1, 0)
            i = 1

            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else

                    If Mid(stxf, 2, Len(stxf) - 1) = ComboBox14.Text Then km = Mid(stxf, 1, 1)

                End If
            Loop
            f.Close()

        End If


        If km = "0" Then

            If Dir("c:\Program Files\方案文件库\" & ComboBox14.Text & ".lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox14.Text & ".lib", 1, 0)
                i = 1
                stxf = f.readline
                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else
                        ComboBox13.Items.Add(Mid(zh_biam_f(stxf), 17, Len(stxf) - 16))
                    End If
                Loop
                f.Close()

            End If
            If ComboBox13.Items.Count <> 0 Then ComboBox13.SelectedIndex = 0
        End If

        If km = "1" Then

            If Dir("c:\Program Files\方案文件库\" & ComboBox14.Text & ".lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox14.Text & ".lib", 1, 0)
                i = 1
                stxf = f.readline
                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else
                        ComboBox13.Items.Add(Mid(zh_biam_f(stxf), 23, Len(stxf) - 22))
                    End If
                Loop
                f.Close()

            End If
            If ComboBox13.Items.Count <> 0 Then ComboBox13.SelectedIndex = 0
        End If
    End Sub


    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        Label19.Text = ""
        tooling_shank(ComboBox3.Text, ComboBox4)
        TextBox1.Focus()
    End Sub


    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        ComboBox5.Items.Clear()
        ComboBox6.Items.Clear()

        Dim i As Integer
        Dim stxf As String
        Dim stxf_ass As String
        Dim fs, f
        Dim ass As Integer '附件
        Dim qh_km As String

        qh_km = qh_km_kz()
        If qh_km <> "0" Then tooling_fjl()
        Label19.Text = ""
        ass = 0

        If Dir("c:\Program Files\方案文件库\" & ComboBox3.Text & ".lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox3.Text & ".lib", 1, 0)
            i = 1

            stxf_ass = f.readline

            If stxf_ass <> "0000" Then
                ComboBox5.Enabled = True
                ComboBox6.Enabled = True

                CheckBox2.Enabled = True
                CheckBox2.CheckState = 1
                ass = 1
            Else
                ComboBox5.Enabled = False
                ComboBox6.Enabled = False

                CheckBox2.Enabled = False
                CheckBox2.CheckState = 0
                ass = 0
            End If


            f.Close()

        End If

        If ass = 1 Then tooling_ass(stxf_ass)


        If Dir("c:\Program Files\方案文件库\" & ComboBox3.Text & "信息.lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox3.Text & "信息.lib", 1, 0)

            i = 1

            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else

                    If Mid(stxf, 1, Len(ComboBox4.Text)) = ComboBox4.Text Then
                        Label19.Text = Mid(stxf, Len(ComboBox4.Text) + 1, Len(stxf) - Len(ComboBox4.Text))
                        i = -1
                    Else
                        Label19.Text = ""
                    End If
                End If
            Loop
        Else
            Label19.Text = ""
        End If
        TextBox1.Focus()
    End Sub


    Public Sub tooling_ass(ByVal Tool_ass_name As String) '工具系统附件
        ComboBox5.Items.Clear()
        ComboBox6.Items.Clear()

        Dim i As Integer
        Dim stxf As String
        Dim fs, f
        'Dim Tool_ass_name As String '附件名称代码
        ' Dim Tool_ass As String '附件型号代码
        'Dim sizhui_bin As String '丝锥柄部尺寸

        Dim stri As Integer
        Dim strk As Integer


        If Dir("c:\Program Files\方案文件库\工具系统附件.lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")

            stri = 1
            strk = 1
            Do While stri <= Len(Tool_ass_name) / 4
                f = fs.OpenTextFile("c:\Program Files\方案文件库\工具系统附件.lib", 1, 0)
                i = 1
                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else
                        If Mid(stxf, 1, 4) = Mid(Tool_ass_name, strk, 4) Then ComboBox6.Items.Add(Mid(stxf, 5, Len(stxf) - 4))
                    End If
                Loop
                f.Close()

                stri = stri + 1
                strk = strk + 4
            Loop
            If ComboBox6.Items.Count <> 0 Then ComboBox6.SelectedIndex = 0
        End If


    End Sub


    Public Sub date_canshu()

        Dim i As Integer
        Dim stxf_xp As String
        ' Dim stxf_xp1 As String
        '  Dim stxf_xp2 As String
        ' Dim stxf_xp3 As String
        ' Dim insert_sb As String
        Dim fs, f


        Dim insert_temp_kz As String

        Dim myExcel As Excel.Application = Nothing  '定义进程
        myExcel = GetObject(, "Excel.Application")  '打开已经打开的excel程序
        myExcel.Visible = True
        insert_temp_kz = myExcel.ActiveSheet.Cells(6, 6).Value


        If Len(ListBox2.Text) <> 0 Then

            If Dir("c:\Program Files\方案文件库\" & insert_temp_kz & ".lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\" & insert_temp_kz & ".lib", 1, 0)

                i = 1

                Do While i = 1
                    stxf_xp = f.readline
                    If stxf_xp = "END" Then
                        i = -1
                    Else

                        If Mid(stxf_xp, 1, Len(ListBox2.Text)) = ListBox2.Text Then

                            If Val(TextBox1.Text) > Val(Mid(stxf_xp, Len(ListBox2.Text) + 7, 10)) And Val(TextBox1.Text) <= Val(Mid(stxf_xp, Len(ListBox2.Text) + 17, 10)) Then

                                TextBox4.Text = Val(Mid(stxf_xp, Len(ListBox2.Text) + 27, 10))
                                TextBox5.Text = Val(Mid(stxf_xp, Len(ListBox2.Text) + 37, 10))
                                i = -1
                            Else

                                TextBox4.Text = 0
                                TextBox5.Text = 0

                            End If
                        Else

                            TextBox4.Text = 0
                            TextBox5.Text = 0

                        End If
                    End If
                Loop
            Else

            End If
        End If
    End Sub





    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TextBox6.Text = zh_long_k()
    End Sub

    Function zh_long_k() As String

        Dim i As Integer
        Dim stxf As String
        Dim fs, f
        Dim ix_p As Integer
        Dim Long_k As Double

        Dim L1 As Double
        Dim L2 As Double
        Dim L3 As Double
        Dim L4 As Double
        Dim L5 As Double



        L1 = 0
        L2 = 0
        L3 = 0
        L4 = 0
        L5 = 0
        Long_k = 0


        If ComboBox1.Text <> "" Then Long_k = t_long

        If ComboBox4.Text <> "" Then
            If Dir("c:\Program Files\方案文件库\" & ComboBox3.Text & ".lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox3.Text & ".lib", 1, 0)
                i = 1

                stxf = f.readline

                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else
                        If ComboBox4.Text = (Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12)) Then
                            L1 = Val(Mid(zh_biam_f(stxf), 17, 6))
                            i = -1
                        End If
                    End If
                Loop
                f.Close()
            End If

            Long_k = Long_k + L1
        End If

        If ComboBox5.Text <> "" Then


            If Dir("c:\Program Files\方案文件库\工具系统附件长度控制.lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\工具系统附件长度控制.lib", 1, 0)
                i = 1

                stxf = f.readline

                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else


                        If Len(ComboBox5.Text) < Len(stxf) Then
                            If ComboBox5.Text = Mid(stxf, Len(stxf) - Len(ComboBox5.Text) + 1, Len(ComboBox5.Text)) Then
                                L2 = Val(Mid(stxf, 1, Len(stxf) - Len(ComboBox5.Text)))
                                i = -1
                            End If
                        End If



                    End If
                Loop
                f.Close()
            End If


            If Mid(stxf1, 5, 1) = "B" Then Long_k = Long_k - L2
            If Mid(stxf1, 5, 1) = "F" Then Long_k = Long_k
        End If

        If ComboBox9.Text <> "" Then
            If Dir("c:\Program Files\方案文件库\" & ComboBox10.Text & ".lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox10.Text & ".lib", 1, 0)
                i = 1

                stxf = f.readline

                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else
                        If ComboBox9.Text = Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12) Then
                            L3 = Val(Mid(zh_biam_f(stxf), 17, 6))
                            i = -1
                        End If
                    End If
                Loop
                f.Close()
            End If
            Long_k = Long_k + L3
        End If


        If ComboBox11.Text <> "" Then

            If Dir("c:\Program Files\方案文件库\附件名称.lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\附件名称.lib", 1, 0)
                i = 1


                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else

                        If ComboBox12.Text = Mid(stxf, 2, Len(ComboBox12.Text)) Then
                            ix_p = Val(Mid(stxf, 1, 1))

                            i = -1
                        End If
                    End If
                Loop
                f.Close()
            End If


            If ix_p = 0 Then
                If Dir("c:\Program Files\方案文件库\工具系统附件长度控制.lib", vbNormal) <> "" Then

                    fs = CreateObject("Scripting.FileSystemObject")
                    f = fs.OpenTextFile("c:\Program Files\方案文件库\工具系统附件长度控制.lib", 1, 0)
                    i = 1

                    stxf = f.readline

                    Do While i = 1
                        stxf = f.readline
                        If stxf = "END" Then
                            i = -1
                        Else

                            If ComboBox11.Text = Mid(stxf, Len(stxf) - Len(ComboBox11.Text) + 1, Len(ComboBox11.Text)) Then
                                L4 = Val(Mid(stxf, 1, Len(stxf) - Len(ComboBox11.Text)))
                                i = -1
                            End If
                        End If
                    Loop
                    f.Close()
                End If
                Long_k = Long_k - L4
            Else

                If Dir("c:\Program Files\方案文件库\" & ComboBox12.Text & ".lib", vbNormal) <> "" Then

                    fs = CreateObject("Scripting.FileSystemObject")
                    f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox12.Text & ".lib", 1, 0)
                    i = 1

                    stxf = f.readline

                    Do While i = 1
                        stxf = f.readline
                        If stxf = "END" Then
                            i = -1
                        Else
                            If ComboBox11.Text = Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12) Then
                                L4 = Val(Mid(zh_biam_f(stxf), 17, 6))
                                i = -1
                            End If
                        End If
                    Loop
                    f.Close()
                End If
                Long_k = Long_k + L4
            End If

        End If


        If ComboBox13.Text <> "" Then

            If Dir("c:\Program Files\方案文件库\附件名称.lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\附件名称.lib", 1, 0)
                i = 1


                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else

                        If ComboBox14.Text = Mid(stxf, 2, Len(ComboBox14.Text)) Then
                            ix_p = Val(Mid(stxf, 1, 1))

                            i = -1
                        End If
                    End If
                Loop
                f.Close()
            End If


            If ix_p = 0 Then
                If Dir("c:\Program Files\方案文件库\工具系统附件长度控制.lib", vbNormal) <> "" Then

                    fs = CreateObject("Scripting.FileSystemObject")
                    f = fs.OpenTextFile("c:\Program Files\方案文件库\工具系统附件长度控制.lib", 1, 0)
                    i = 1

                    stxf = f.readline

                    Do While i = 1
                        stxf = f.readline
                        If stxf = "END" Then
                            i = -1
                        Else

                            If ComboBox13.Text = Mid(stxf, Len(stxf) - Len(ComboBox13.Text) + 1, Len(ComboBox13.Text)) Then
                                L5 = Val(Mid(stxf, 1, Len(stxf) - Len(ComboBox13.Text)))
                                i = -1
                            End If
                        End If
                    Loop
                    f.Close()
                End If
                Long_k = Long_k - L5
            Else

                If Dir("c:\Program Files\方案文件库\" & ComboBox14.Text & ".lib", vbNormal) <> "" Then

                    fs = CreateObject("Scripting.FileSystemObject")
                    f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox14.Text & ".lib", 1, 0)
                    i = 1

                    stxf = f.readline

                    Do While i = 1
                        stxf = f.readline
                        If stxf = "END" Then
                            i = -1
                        Else
                            If ComboBox13.Text = Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12) Then
                                L5 = Val(Mid(zh_biam_f(stxf), 17, 6))
                                i = -1
                            End If
                        End If
                    Loop
                    f.Close()
                End If
                Long_k = Long_k + L5
            End If
        End If

        Return Long_k

    End Function

    Private Sub ComboBox6_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox6.SelectedIndexChanged
        Dim i As Integer
        Dim stxf As String
        Dim fs, f

        Dim gs_km_cb As Integer

        Dim Tool_ass As String

        ' Dim tooling_insert_k As String '刀片型号
        'ComboBox4.Items.Add(Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12))

        If Dir("c:\Program Files\方案文件库\" & ComboBox3.Text & ".lib", vbNormal) <> "" Then
            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox3.Text & ".lib", 1, 0)
            i = 1
            stxf = f.readline
            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else


                    If Mid(zh_biam_f(stxf), 23, Len(zh_biam_f(stxf)) - 22 - 12) = ComboBox4.Text Then
                        Tool_ass = Mid(zh_biam_f(stxf), 1, 4)
                        i = -1
                    End If
                End If
            Loop
            f.Close()
        End If

        ComboBox5.Items.Clear()


        gs_km_cb = 0
        If ComboBox6.Text = "攻丝卡簧" Then gs_km_cb = 1

        If gs_km_cb = 0 Then

            '通用夹持
            If Dir("c:\Program Files\方案文件库\" & ComboBox6.Text & ".lib", vbNormal) <> "" Then
                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox6.Text & ".lib", 1, 0)
                i = 1

                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else


                        If Val(Tool_ass) = Val(Mid(zh_biam_f(stxf), 1, 4)) And Val(TextBox2.Text) >= Val(Mid(zh_biam_f(stxf), 5, 6)) And Val(TextBox2.Text) <= Val(Mid(zh_biam_f(stxf), 11, 6)) Then
                            ComboBox5.Items.Add(Mid(zh_biam_f(stxf), 17, Len(zh_biam_f(stxf)) - 16))
                        End If
                    End If
                Loop
                f.Close()
            End If


            If ComboBox5.Items.Count <> 0 Then ComboBox5.SelectedIndex = 0
        End If


        If gs_km_cb = 1 Then
            '攻丝卡簧夹持
            If Dir("c:\Program Files\方案文件库\" & ComboBox6.Text & ".lib", vbNormal) <> "" Then
                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox6.Text & ".lib", 1, 0)
                i = 1

                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else
                        ' MsgBox(Mid(zh_biam_f(stxf), 5, 6))
                        ' MsgBox(Mid(zh_biam_f(stxf), 11, 6))

                        If Val(Tool_ass) = Val(Mid(zh_biam_f(stxf), 1, 4)) And Val(TextBox2.Text) = Val(Mid(zh_biam_f(stxf), 5, 6)) And t_yx_long = Val(Mid(zh_biam_f(stxf), 11, 6)) Then
                            ComboBox5.Items.Add(Mid(zh_biam_f(stxf), 17, Len(zh_biam_f(stxf)) - 16))
                        End If
                    End If
                Loop
                f.Close()
            End If


            If ComboBox5.Items.Count <> 0 Then ComboBox5.SelectedIndex = 0
        End If

        TextBox1.Focus()

    End Sub


    Private Sub ComboBox5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged

        Dim i As Integer
        Dim stxf As String
        Dim fs As New Object
        Dim f As New Object
        If Dir("c:\Program Files\方案文件库\" & ComboBox6.Text & "信息.lib", vbNormal) <> "" Then
            Try
                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox6.Text & "信息.lib", 1, 0)

                i = 1

                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else

                        If Mid(stxf, 1, Len(ComboBox5.Text)) = ComboBox5.Text Then
                            Label18.Text = Mid(stxf, Len(ComboBox5.Text) + 1, Len(stxf) - Len(ComboBox5.Text))
                            i = -1
                        Else
                            Label19.Text = ""
                        End If
                    End If
                Loop
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        Else

        End If
        'f.Close()
        TextBox1.Focus()
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        TextBox1.Focus()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        TextBox1.Focus()
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        TextBox1.Focus()
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        TextBox1.Focus()
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        TextBox1.Focus()
    End Sub

    ''' <summary>
    ''' 非标OK按钮
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        GlobalData.rowsToExcel = 0

        Dim i As Integer
        Dim my_col As Integer '列
        Dim my_row As Integer '行

        Dim myExcel As Excel.Application = Nothing  '定义进程
        myExcel = GetObject(, "Excel.Application")  '打开已经打开的excel程序
        myExcel.Visible = True

        my_row = myExcel.ActiveCell.Row
        my_col = myExcel.ActiveCell.Column

        '确定刀具序号为几：my_row/11的商，再乘11再加1即可
        Dim tmprow As Integer
        Try
            '刀具型号是否为空
            If TextBox13.Text.Trim = "" Then
                MessageBox.Show("请输入刀具型号。" + Application.ProductVersion)
                Return
            Else
                '''''''''''''''''''''''''''判断需要几行EXCEL'''''''''''''''''''''''''''''''''
                '刀具型号
                GlobalData.rowsToExcel += 1
                '刀片1
                If ComboBox15.Text <> "" Then GlobalData.rowsToExcel += 1
                '刀片2
                If ComboBox16.Text <> "" Then GlobalData.rowsToExcel += 1
                '刀片3
                If ComboBox17.Text <> "" Then GlobalData.rowsToExcel += 1
                '刀片附件1
                If TextBox12.Text <> "" Then GlobalData.rowsToExcel += 1
                '刀片附件2
                If TextBox16.Text <> "" Then GlobalData.rowsToExcel += 1
                '刀片附件3
                If TextBox22.Text <> "" Then GlobalData.rowsToExcel += 1
                '工具系统
                If ComboBox24.Text <> "" Then GlobalData.rowsToExcel += 1
                '工具系统附件1
                If CheckBox11.Checked = True Then GlobalData.rowsToExcel += 1
                '工具系统附件2
                If CheckBox10.Checked = True Then GlobalData.rowsToExcel += 1
                '工具系统附件3
                If CheckBox15.Checked = True Then GlobalData.rowsToExcel += 1
                '增加一栏添加拉钉型号
                GlobalData.rowsToExcel += 1

                'MsgBox("一共需要：" + GlobalData.rowsToExcel.ToString)

                ''''''''''''''''''''''''''''''''调整行数'''''''''''''''''''''''''''''
                OperationTools.AdjustRows(GlobalData.rowsNow, GlobalData.rowsToExcel, GlobalData.rowActivite, myExcel)
                '''''''''''''''''''''''''''''''清除已有数据''''''''''''''''''''''''''''''''
                Dim tmpRange As Excel.Range
                If GlobalData.rowsToExcel <= 4 Then
                    tmpRange = myExcel.Range(Cell1:=myExcel.Cells(GlobalData.rowActivite, GlobalData.colActivite),
                              Cell2:=myExcel.Cells(GlobalData.rowActivite + 3, GlobalData.colActivite + 20))
                Else
                    tmpRange = myExcel.Range(Cell1:=myExcel.Cells(GlobalData.rowActivite, GlobalData.colActivite),
                              Cell2:=myExcel.Cells(GlobalData.rowActivite + GlobalData.rowsToExcel - 1, GlobalData.colActivite + 20))
                End If
                'tmpRange.Value2 = Nothing
                tmpRange.ClearContents()
                'myExcel.Range(Cell1:=myExcel.Cells(GlobalData.rowActivite, GlobalData.colActivite),
                '              Cell2:=myExcel.Cells(GlobalData.rowActivite + GlobalData.rowsToExcel, GlobalData.colActivite + 20)).Value2 = ""
                '清除批注
                Try
                    tmpRange.ClearComments()
                Catch ex As Exception

                End Try
                '''''''''''''''''''''''''''''''写入数据''''''''''''''''''''''''''''''''
                tmprow = GlobalData.rowActivite
                my_col = GlobalData.colActivite
                ''''''刀具型号
                myExcel.ActiveSheet.Cells(tmprow, my_col).value = “非标刀具”
                myExcel.ActiveSheet.Cells(tmprow, my_col + 1).value = TextBox13.Text
                ''刀具参数
                myExcel.ActiveSheet.Cells(tmprow, 6).Value = TextBox9.Text  '刀具直径
                myExcel.ActiveSheet.Cells(tmprow, 7).Value = TextBox11.Text '刀具长度
                myExcel.ActiveSheet.Cells(tmprow, 8).Value = TextBox10.Text '线速度
                myExcel.ActiveSheet.Cells(tmprow, 9).Value = "=H" & tmprow & "*1000/3.14/" & "F" & tmprow
                myExcel.ActiveSheet.Cells(tmprow, 10).Value = TextBox8.Text '每齿进给
                'myExcel.ActiveSheet.Cells(tmprow, 11).Value = "=J" & tmprow & "*I" & tmprow & "*" & TextBox3.Text
                myExcel.ActiveSheet.Cells(tmprow, 11).Value = ""
                myExcel.ActiveSheet.Cells(tmprow, 13).Value = 0
                myExcel.ActiveSheet.Cells(tmprow, 14).Value = 1
                myExcel.ActiveSheet.Cells(tmprow, 15).Value = "=M" & tmprow & "*N" & tmprow & "/K" & tmprow
                myExcel.ActiveSheet.Cells(tmprow, 16).Value = "=P5/60"
                myExcel.ActiveSheet.Cells(tmprow, 17).Value = "0"
                myExcel.ActiveSheet.Cells(tmprow, 18).Value = "=O" & tmprow & "+P" & tmprow & "+Q" & tmprow

                ''''''刀片1:rowActivite+1
                If ComboBox15.Text.Trim <> "" Then
                    tmprow = tmprow + 1
                    '写入刀片型号
                    myExcel.ActiveSheet.Cells(tmprow, my_col).value = "刀片"
                    myExcel.ActiveSheet.Cells(tmprow, my_col + 1).value = ComboBox15.Text
                    '添加刀片数量批注
                    If TextBox18.Text.Trim <> "" Then
                        myExcel.ActiveSheet.Cells(tmprow, my_col + 1).AddComment(Text:="数量：" + TextBox18.Text + "片")
                    End If
                End If
                ''''''刀片2:rowActivite+2
                If ComboBox16.Text.Trim <> "" Then
                    tmprow = tmprow + 1
                    '写入刀片型号
                    myExcel.ActiveSheet.Cells(tmprow, my_col).value = "刀片"
                    myExcel.ActiveSheet.Cells(tmprow, my_col + 1).value = ComboBox16.Text
                    '添加刀片数量批注
                    If TextBox19.Text.Trim <> "" Then
                        myExcel.ActiveSheet.Cells(tmprow, my_col + 1).AddComment(Text:="数量：" + TextBox19.Text + "片")
                    End If
                End If
                ''''''刀片3:rowActivite+3
                If ComboBox17.Text.Trim <> "" Then
                    tmprow = tmprow + 1
                    '写入刀片型号
                    myExcel.ActiveSheet.Cells(tmprow, my_col).value = "刀片"
                    myExcel.ActiveSheet.Cells(tmprow, my_col + 1).value = ComboBox17.Text
                    '添加刀片数量批注
                    If TextBox21.Text.Trim <> "" Then
                        myExcel.ActiveSheet.Cells(tmprow, my_col + 1).AddComment(Text:="数量：" + TextBox21.Text + "片")
                    End If
                End If
                ''''''刀片附件1:rowActivite+4
                If TextBox12.Text.Trim <> "" Then
                    tmprow = tmprow + 1
                    myExcel.ActiveSheet.Cells(tmprow, my_col).value = TextBox12.Text
                End If
                ''''''刀片附件2:my_row+5
                If TextBox16.Text.Trim <> "" Then
                    tmprow = tmprow + 1
                    myExcel.ActiveSheet.Cells(tmprow, my_col).value = TextBox16.Text
                End If
                ''''''刀片附件3:my_row+6
                If TextBox22.Text.Trim <> "" Then
                    tmprow = tmprow + 1
                    myExcel.ActiveSheet.Cells(tmprow, my_col).value = TextBox22.Text
                End If
                ''''''工具系统：my_row+7
                If ComboBox24.Text.Trim <> "" Then
                    tmprow = tmprow + 1
                    '名称
                    myExcel.ActiveSheet.Cells(tmprow, my_col).value = ComboBox24.Text
                    '型号
                    myExcel.ActiveSheet.Cells(tmprow, my_col + 1).value = ComboBox23.Text
                End If
                ''''''工具系统附件1：my_row+8
                If CheckBox11.Checked = True Then
                    tmprow = tmprow + 1
                    myExcel.ActiveSheet.Cells(tmprow, my_col).value = ComboBox22.Text
                    myExcel.ActiveSheet.Cells(tmprow, my_col + 1).value = ComboBox21.Text
                End If

                ''''''工具系统附件2：my_row+9
                If CheckBox10.Checked = True Then
                    tmprow = tmprow + 1
                    myExcel.ActiveSheet.Cells(tmprow, my_col).value = ComboBox19.Text
                    myExcel.ActiveSheet.Cells(tmprow, my_col + 1).value = ComboBox20.Text
                End If
                ''''''工具系统附件3：my_row+10
                If CheckBox15.Checked = True Then
                    tmprow = tmprow + 1
                    myExcel.ActiveSheet.Cells(tmprow, my_col).value = ComboBox29.Text
                    myExcel.ActiveSheet.Cells(tmprow, my_col + 1).value = ComboBox30.Text
                End If

                ''''''插入拉钉型号
                tmprow = tmprow + 1
                myExcel.ActiveSheet.Cells(tmprow, my_col).value = “拉钉”
                myExcel.ActiveSheet.Cells(tmprow, my_col + 1).value = "=E3"




                '''''''''''''''''''''''''''''''插入图片''''''''''''''''''''''''''''''''
                If PictureBox9.Image IsNot Nothing Then
                    Try
                        myExcel.ActiveSheet.Cells(GlobalData.rowActivite, 22).Activate()
                        Dim pic = myExcel.ActiveSheet.Pictures.Insert(GlobalData.picToExcel)
                        'pic.select.ShapeRange.LockAspectRatio = False
                        'pic.LockAspectRatio = False
                        Dim tmpRange2 As Excel.Range
                        tmpRange2 = myExcel.ActiveSheet.Range("V12:X15")
                        pic.Height = tmpRange2.Height
                        'MsgBox(pic.Height.ToString)
                        'pic.Width = tmpRange2.Width
                        'MsgBox(pic.Width.ToString)
                    Catch ex As Exception
                        MsgBox("插入图片出错：Error-" + ex.Message)
                    End Try

                End If
            End If
        Catch ex As Exception
            MsgBox("error:" + ex.Message)
        End Try

        Me.Close()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub ComboBox15_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox15.TextUpdate
        '刀片1
        '显示图片到对应图片框
        OperationTools.ShowImage(ComboBox15.Text.Trim, ConstData.bladeImageDir, PictureBox8)

        '执行模糊查询
        OperationTools.FuzzyQuery(sender, GlobalData.bladeTypeShort)
        Cursor = Cursors.Default
    End Sub
    Private Sub ComboBox17_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox17.TextUpdate
        '刀片3
        '显示图片到对应图片框
        OperationTools.ShowImage(ComboBox17.Text.Trim, ConstData.bladeImageDir, PictureBox4)

        '执行模糊查询
        OperationTools.FuzzyQuery(sender, GlobalData.bladeTypeShort)
        Cursor = Cursors.Default
    End Sub
    Private Sub ComboBox16_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox16.TextUpdate
        '刀片2
        '显示图片到对应图片框
        OperationTools.ShowImage(ComboBox16.Text.Trim, ConstData.bladeImageDir, PictureBox3)

        '执行模糊查询
        OperationTools.FuzzyQuery(sender, GlobalData.bladeTypeShort)
        Cursor = Cursors.Default
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        ComboBox21.Enabled = CheckBox11.Checked
        ComboBox22.Enabled = CheckBox11.Checked
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        ComboBox19.Enabled = CheckBox10.Checked
        ComboBox20.Enabled = CheckBox10.Checked
    End Sub

    Private Sub CheckBox15_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox15.CheckedChanged
        ComboBox29.Enabled = CheckBox15.Checked
        ComboBox30.Enabled = CheckBox15.Checked
    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        '刀具附件1
        '显示图片到对应图片框
        OperationTools.ShowImage(TextBox12.Text.Trim, ConstData.bladeAttachImageDir, PictureBox5)
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged
        '刀具附件2
        '显示图片到对应图片框
        OperationTools.ShowImage(TextBox16.Text.Trim, ConstData.bladeAttachImageDir, PictureBox6)
    End Sub

    Private Sub TextBox22_TextChanged(sender As Object, e As EventArgs) Handles TextBox22.TextChanged
        '刀具附件3
        '显示图片到对应图片框
        OperationTools.ShowImage(TextBox22.Text.Trim, ConstData.bladeAttachImageDir, PictureBox7)
    End Sub

    ''' <summary>
    ''' 激活第二个选项卡事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TabControl1_Selected(sender As Object, e As TabControlEventArgs) Handles TabControl1.Selected
        If TabControl1.SelectedIndex = 1 Then
            '激活第二个选项卡时
            Dim myExcel As Excel.Application = GetObject(, "Excel.Application")  '打开已经打开的excel程序
            myExcel.Visible = True
            TextBox13.Text = myExcel.ActiveSheet.Cells(GlobalData.rowActivite, 5).Value
            '读取刀片型号并保存在集合中
            '读取全部刀具
            OperationTools.GetLibToCollection("刀片型号.lib", GlobalData.bladeType)
            OperationTools.GetLibToCollection("刀片型号.lib", GlobalData.bladeTypeShort, 16, 0)
            OperationTools.GetLibToComboBox("刀片型号.lib", ComboBox15, 16, 0)
            OperationTools.GetLibToComboBox("刀片型号.lib", ComboBox16, 16, 0)
            OperationTools.GetLibToComboBox("刀片型号.lib", ComboBox17, 16, 0)

            '读取符合材料的刀具

            '读取全部工具系统
            OperationTools.GetLibToCollection("工具系统.lib", GlobalData.holderType)
            OperationTools.GetLibToCollection("工具系统.lib", GlobalData.holderTypeShort, 4, 0)
            OperationTools.GetLibToComboBox("工具系统.lib", ComboBox24, 4, 0)
            '读取符合材料工具系统


            '读取工具系统附件
            OperationTools.GetLibToComboBox("工具系统附件.lib", ComboBox22, 4, 0)
            OperationTools.GetLibToComboBox("工具系统附件.lib", ComboBox19, 4, 0)
            OperationTools.GetLibToComboBox("工具系统附件.lib", ComboBox29, 4, 0)



            'MsgBox("选择1")
        End If
        If TabControl1.SelectedIndex = 0 Then
            'MsgBox("选择0")
        End If
    End Sub

    Private Sub ComboBox24_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox24.TextUpdate
        OperationTools.FuzzyQuery(ComboBox24, GlobalData.holderTypeShort)
        Cursor = Cursors.Default
    End Sub

    Private Sub ComboBox15_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox15.SelectedIndexChanged
        '显示图片到对应图片框
        OperationTools.ShowImage(ComboBox15.Text.Trim, ConstData.bladeImageDir, PictureBox8)
    End Sub

    Private Sub ComboBox16_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox16.SelectedIndexChanged
        '显示图片到对应图片框
        OperationTools.ShowImage(ComboBox16.Text.Trim, ConstData.bladeImageDir, PictureBox3)
    End Sub

    Private Sub ComboBox17_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox17.SelectedIndexChanged
        '显示图片到对应图片框
        OperationTools.ShowImage(ComboBox17.Text.Trim, ConstData.bladeImageDir, PictureBox4)
    End Sub

    ''' <summary>
    ''' 工具系统名称选定事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ComboBox24_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox24.SelectedIndexChanged

        tooling_shank(ComboBox24.Text.Substring(0, ComboBox24.Text.Length - 1), ComboBox23)

        'ComboBox3.SelectedText = ComboBox24.Text.Substring(0, ComboBox24.Text.Length - 1)
        'tooling_shank()
        'Dim i As Integer = 0
        'For i = 0 To ComboBox4.Items.Count
        '    ComboBox23.Items.Add(ComboBox4.Items(i))
        'Next
        '通过工具系统名称获取对应配置文件中的型号
        'MsgBox(ComboBox24.Text.Substring(0, ComboBox24.Text.Length - 1))
        'OperationTools.GetLibToComboBox(ComboBox24.Text.Substring(0, ComboBox24.Text.Length - 1) & ".lib", ComboBox23, 22, 12)


    End Sub

    ''' <summary>
    ''' 工具系统附件一
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ComboBox22_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox22.SelectedIndexChanged
        'OperationTools.GetHolderAttachToCombBox(ComboBox24, ComboBox23, ComboBox21, ComboBox22, TextBox14)
        OperationTools.GetLibToComboBox(ComboBox22.Text + ".lib", ComboBox21, 16, 0)

    End Sub

    ''' <summary>
    ''' 工具系统附件二
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ComboBox19_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox19.SelectedIndexChanged
        OperationTools.GetLibToComboBox(ComboBox19.Text + ".lib", ComboBox20, 16, 0)
    End Sub

    ''' <summary>
    ''' 工具系统附件三
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ComboBox29_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox29.SelectedIndexChanged
        OperationTools.GetLibToComboBox(ComboBox29.Text + ".lib", ComboBox30, 16, 0)
    End Sub
    ''' <summary>
    ''' 刀具型号
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        '刀具型号
        '显示图片到对应图片框
        Dim toolTypeRef As String
        toolTypeRef = OperationTools.GetToolTypeRef(TextBox13.Text.Trim)
        If toolTypeRef <> "" Then
            GlobalData.picToExcel = ConstData.toolTypeRefImageDir + toolTypeRef + ".jpg"
            Label27.Text = toolTypeRef + ".jpg"
            OperationTools.ShowImage(toolTypeRef.Trim, ConstData.toolTypeRefImageDir, PictureBox9)
        Else
            GlobalData.picToExcel = ""
        End If

    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        tooling_shank(ComboBox24.Text.Substring(0, ComboBox24.Text.Length - 1), ComboBox23)
    End Sub

End Class










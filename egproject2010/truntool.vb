Imports System.Windows.Forms

Public Class truntool

    Dim tools_name As String

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim my_col As Integer '列
        Dim my_row As Integer '行

        Dim myExcel As Excel.Application = Nothing  '定义进程
        myExcel = GetObject(, "Excel.Application")  '打开已经打开的excel程序
        myExcel.Visible = True
        'sb_ty = Mid(myExcel.ActiveSheet.Cells(6, 6).Value, 1, 1)

        my_row = myExcel.ActiveCell.Row
        my_col = myExcel.ActiveCell.Column

        myExcel.ActiveSheet.Cells(my_row, 4).Value = TextBox1.Text
        myExcel.ActiveSheet.Cells(my_row, 5).Value = ComboBox1.Text

        my_row = my_row + 1

        myExcel.ActiveSheet.Cells(my_row, 4).Value = "刀片"
        myExcel.ActiveSheet.Cells(my_row, 5).Value = ComboBox2.Text

        my_row = my_row + 1


        If CheckBox2.CheckState = 1 Then

            myExcel.ActiveSheet.Cells(my_row, 4).Value = ComboBox4.Text
            myExcel.ActiveSheet.Cells(my_row, 5).Value = ComboBox3.Text

            my_row = my_row + 1

        End If

        If CheckBox3.CheckState = 1 Then

            myExcel.ActiveSheet.Cells(my_row, 4).Value = ComboBox5.Text
            myExcel.ActiveSheet.Cells(my_row, 5).Value = ComboBox6.Text

            my_row = my_row + 1

        End If

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub truntool_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CheckBox2.CheckState = 0
        CheckBox3.CheckState = 0

        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        ComboBox5.Enabled = False
        ComboBox6.Enabled = False

        PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
       
    End Sub

    Public Sub tools_xuanzhe(ByVal tools_xuanzhe As String)
        Dim i As Integer
        Dim stxf As String
        Dim stxf1 As String
        Dim stxf2 As String
        Dim fs, f
        Dim insert_temp_kz As String


        Dim myExcel As Excel.Application = Nothing  '定义进程
        myExcel = GetObject(, "Excel.Application")  '打开已经打开的excel程序
        myExcel.Visible = True

        insert_temp_kz = myExcel.ActiveSheet.Cells(6, 5).Value

        TextBox1.Text = ""
        ComboBox1.Items.Clear()
        ComboBox2.Items.Clear()
        ComboBox3.Items.Clear()
        ComboBox4.Items.Clear()
        ComboBox5.Items.Clear()
        ComboBox6.Items.Clear()

        tools_name = tools_xuanzhe



        If Dir("c:\Program Files\方案文件库\" & tools_xuanzhe & ".lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & tools_xuanzhe & ".lib", 1, 0)

            stxf1 = f.readline
            stxf2 = f.readline

            If Val(stxf1) = 1 Then TextBox1.Text = "复合压紧式外圆车刀"
            If Val(stxf1) = 2 Then TextBox1.Text = "复合压紧式内孔车刀"
            If Val(stxf1) = 3 Then TextBox1.Text = "螺钉压紧式外圆车刀"
            If Val(stxf1) = 4 Then TextBox1.Text = "螺钉压紧式内孔车刀"
            If Val(stxf1) = 5 Then TextBox1.Text = "外螺纹车刀"
            If Val(stxf1) = 6 Then TextBox1.Text = "内螺纹车刀"
            If Val(stxf1) = 7 Then TextBox1.Text = "QA切槽(断)刀"
            If Val(stxf1) = 8 Then TextBox1.Text = "GRI内切槽刀"
            If Val(stxf1) = 9 Then TextBox1.Text = "GRV内切槽刀"
            If Val(stxf1) = 10 Then TextBox1.Text = "QD切槽(断)刀"
            If Val(stxf1) = 11 Then TextBox1.Text = "QD内孔切槽刀"
            If Val(stxf1) = 12 Then TextBox1.Text = "GRE外圆切槽刀"
            If Val(stxf1) = 13 Then TextBox1.Text = "QD端面切槽刀"

            i = 1

            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else

                    If insert_temp_kz <> "" Then

                        If Val(stxf2) = 1 Then
                            If insert_temp_kz = Mid(zh_biam_f(stxf), 18, 4) Or insert_temp_kz = Mid(zh_biam_f(stxf), 16, 4) Or insert_temp_kz = Mid(zh_biam_f(stxf), 15, 4) Or insert_temp_kz = Mid(zh_biam_f(stxf), 18, 4) Then ComboBox1.Items.Add(Mid(zh_biam_f(stxf), 13, Len(zh_biam_f(stxf)) - 12))
                        Else

                            ComboBox1.Items.Add(Mid(zh_biam_f(stxf), 13, Len(zh_biam_f(stxf)) - 12))
                        End If
                    Else

                        ComboBox1.Items.Add(Mid(zh_biam_f(stxf), 13, Len(zh_biam_f(stxf)) - 12))


                    End If


                End If

            Loop

            f.Close()

        End If

        If ComboBox1.Items.Count <> 0 <> 0 Then
            ComboBox1.SelectedIndex = 0
        Else
            ComboBox1.Text = "数据不存在,请核对"
        End If

    End Sub


    Function zh_biam_f(ByVal xd As String) As String
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


    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "MCLN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "MCLN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("95°MCLN")
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "MCBN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "MCBN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("75°MCBN")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        Dim i As Integer
        Dim stxf As String
        Dim stxf1 As String
        Dim stxf2 As String
        Dim fs, f

        Dim tooling_insert_k As String '刀片型号

        If Dir("c:\Program Files\方案文件库\" & tools_name & ".lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & tools_name & ".lib", 1, 0)

            stxf1 = f.readline
            stxf2 = f.readline

            i = 1
            Do While i = 1
                stxf = f.readline

                If stxf = "END" Then
                    i = -1
                Else

                    If ComboBox1.Text = Mid(zh_biam_f(stxf), 13, Len(ComboBox1.Text)) Then
                        tooling_insert_k = Mid(zh_biam_f(stxf), 1, 4)
                        tooling_insert(tooling_insert_k)
                    End If


                End If
            Loop
            f.Close()

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
                        
                        If insert_km = Mid(zh_biam_f(stxf), 1, 4) Then

                            If insert_temp_kz = "" Or Len(insert_temp_kz) <> 4 Then
                                ComboBox2.Items.Add(Mid(zh_biam_f(stxf), 18, Len(zh_biam_f(stxf)) - 16))
                            Else

                                If Mid(insert_temp_kz, 2, 1) = Mid(zh_biam_f(stxf), 5, 1) Or Mid(insert_temp_kz, 2, 1) = Mid(zh_biam_f(stxf), 6, 1) Or Mid(insert_temp_kz, 2, 1) = Mid(zh_biam_f(stxf), 7, 1) Or Mid(insert_temp_kz, 2, 1) = Mid(zh_biam_f(stxf), 8, 1) Or Mid(insert_temp_kz, 2, 1) = Mid(zh_biam_f(stxf), 9, 1) Or Mid(insert_temp_kz, 2, 1) = Mid(zh_biam_f(stxf), 10, 1) Then
                                    ComboBox2.Items.Add(Mid(zh_biam_f(stxf), 18, Len(zh_biam_f(stxf)) - 16))
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

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        Dim fs, f
        Dim i As Integer
        Dim stxf As String



        If CheckBox2.CheckState = 1 Then
            ComboBox4.Enabled = True
            ComboBox3.Enabled = True


            If Dir("c:\Program Files\方案文件库\车刀附件.lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\车刀附件.lib", 1, 0)
                i = 1
                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else
                        ComboBox4.Items.Add(stxf)

                    End If
                Loop
                f.Close()
                If ComboBox4.Items.Count <> 0 Then ComboBox4.SelectedIndex = 0
            End If

        Else
            ComboBox4.Items.Clear()
            ComboBox5.Items.Clear()
            ComboBox5.Enabled = False
            ComboBox4.Enabled = False
        End If

    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        Dim fs, f
        Dim i As Integer
        Dim stxf As String



        If CheckBox3.CheckState = 1 Then
            ComboBox5.Enabled = True
            ComboBox6.Enabled = True


            If Dir("c:\Program Files\方案文件库\车刀附件.lib", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\车刀附件.lib", 1, 0)
                i = 1
                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else
                        ComboBox5.Items.Add(stxf)

                    End If
                Loop
                f.Close()
                If ComboBox5.Items.Count <> 0 Then ComboBox5.SelectedIndex = 0
            End If

        Else
            ComboBox5.Items.Clear()
            ComboBox6.Items.Clear()
            ComboBox5.Enabled = False
            ComboBox6.Enabled = False
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        Dim i As Integer
        Dim stxf As String
        Dim fs, f

        ComboBox3.Items.Clear()

        If Dir("c:\Program Files\方案文件库\" & ComboBox4.Text & ".lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox4.Text & ".lib", 1, 0)
            i = 1

            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else
                    ComboBox3.Items.Add(stxf)
                End If
            Loop
            f.Close()

        End If
        If ComboBox3.Items.Count <> 0 Then ComboBox3.SelectedIndex = 0

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged
        Dim i As Integer
        Dim stxf As String
        Dim fs, f

        ComboBox6.Items.Clear()

        If Dir("c:\Program Files\方案文件库\" & ComboBox5.Text & ".lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & ComboBox5.Text & ".lib", 1, 0)
            i = 1

            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else
                    ComboBox6.Items.Add(stxf)
                End If
            Loop
            f.Close()

        End If
        If ComboBox6.Items.Count <> 0 Then ComboBox6.SelectedIndex = 0

    End Sub

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "MDJN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "MDJN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("93°MDJN")
    End Sub

    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "MDNN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "MDNN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("62.5°MDNNN")
    End Sub

    Private Sub PictureBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox5.Click

        If Dir("c:\Program Files\方案文件库\图片\" & "MDHN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "MDHN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("107.5°MDHN")

    End Sub

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "MSDN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "MSDN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("45°MSDN")
    End Sub

    Private Sub PictureBox7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox7.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "MSKN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "MSKN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("75°MSKN")
    End Sub

    Private Sub PictureBox11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox11.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "MSRN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "MSRN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("75°MSRN")
    End Sub

    Private Sub PictureBox13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox13.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "MSBN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "MSBN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("75°MSBN")
    End Sub

    Private Sub PictureBox15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox15.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "MSSN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "MSSN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("45°MSSN")
    End Sub

    Private Sub PictureBox17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox17.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "MTFN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "MTFN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("90°MTFN")
    End Sub

    Private Sub PictureBox19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox19.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "MTGN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "MTGN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("90°MTGN")
    End Sub

    Private Sub PictureBox8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox8.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "MTJN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "MTJN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("93°MTJN")
    End Sub

    Private Sub PictureBox10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox10.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "MVJN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "MVJN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("93°MVJN")
    End Sub

    Private Sub PictureBox12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox12.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "MVVN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "MVVN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("72.5°MVVNN")
    End Sub

    Private Sub PictureBox14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox14.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "MWLN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "MWLN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("95°MWLN")
    End Sub



    Private Sub PictureBox48_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox48.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "SCFC" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "SCFC" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("90°SCFC")
    End Sub

    Private Sub PictureBox47_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox47.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "SCGC" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "SCGC" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("90°SCGC")
    End Sub

    Private Sub PictureBox46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox46.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "SCLC" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "SCLC" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("95°SCLC")
    End Sub

    Private Sub PictureBox45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox45.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "SDJC" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "SDJC" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("95°SDJC")
    End Sub

    Private Sub PictureBox44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox44.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "SDNCN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "SDNCN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("62.5°SDNC")
    End Sub

    Private Sub PictureBox43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox43.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "SRAC" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "SRAC" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("SRAC")
    End Sub

    Private Sub PictureBox42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox42.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "SRDC" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "SRDC" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("SRDC")
    End Sub

    Private Sub PictureBox40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox40.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "SRGC" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "SRGC" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("SRGC")
    End Sub

    Private Sub PictureBox38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox38.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "SSBC" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "SSBC" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("75°SSBC")
    End Sub

    Private Sub PictureBox36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox36.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "SSDC" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "SSDC" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("45°SSDCN")
    End Sub

    Private Sub PictureBox34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox34.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "SSKC" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "SSKC" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("75°SSKC")
    End Sub

    Private Sub PictureBox32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox32.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "SSSC" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "SSSC" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("45°SSSC")
    End Sub

    Private Sub PictureBox41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox41.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "STFC" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "STFC" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("90°STFC")
    End Sub

    Private Sub PictureBox39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox39.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "STGC" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "STGC" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("90°STGC")
    End Sub

    Private Sub PictureBox37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox37.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "SVJC" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "SVJC" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("93°SVJC")
    End Sub

    Private Sub PictureBox35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox35.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "SVVC" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "SVVC" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("72.5°SVVCN")
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Dim i As Integer
        Dim stxf As String
        Dim stxf1 As String
        Dim stxf2 As String
        Dim fs, f

        Dim tooling_insert_k As String '刀片型号

        If Dir("c:\Program Files\方案文件库\" & tools_name & ".lib", vbNormal) <> "" Then

            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & tools_name & ".lib", 1, 0)

            stxf1 = f.readline
            stxf2 = f.readline

            i = 1
            Do While i = 1
                stxf = f.readline

                If stxf = "END" Then
                    i = -1
                Else

                    If ComboBox1.Text = Mid(zh_biam_f(stxf), 13, Len(ComboBox1.Text)) Then
                        tooling_insert_k = Mid(zh_biam_f(stxf), 1, 4)
                        tooling_insert(tooling_insert_k)
                    End If


                End If
            Loop
            f.Close()

        End If
    End Sub

    Private Sub PictureBox31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox31.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MCFN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MCFN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("90°S..MCFN")
    End Sub

    Private Sub PictureBox30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox30.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MCLN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MCLN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("95°S..MCLN")
    End Sub

    Private Sub PictureBox29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox29.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MDQN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MDQN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("107.5°S..MDQN")
    End Sub

    Private Sub PictureBox28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox28.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MDUN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MDUN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("93°S..MDUN")
    End Sub

    Private Sub PictureBox27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox27.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MSKN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MSKN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("75°S..MSKN")
    End Sub

    Private Sub PictureBox26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox26.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MTFN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MTFN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("90°S..MTFN")
    End Sub

    Private Sub PictureBox25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox25.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MCKN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MCKN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("75°S..MCKN")
    End Sub

    Private Sub PictureBox24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox24.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MDNN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MDNN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("62.5°S..MDNN")
    End Sub

    Private Sub PictureBox23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox23.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MVLN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MVLN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("95°S..MVLN")
    End Sub

    Private Sub PictureBox22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox22.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MVUN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MVUN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("93°S..MVUN")
    End Sub

    Private Sub PictureBox21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox21.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MWLN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MWLN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("95°S..MWLN")
    End Sub

    'hyhfhhdfhd


    Private Sub PictureBox59_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox59.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MCFN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MCFN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("90°S..SCFC")
    End Sub

    Private Sub PictureBox58_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox58.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MCLN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MCLN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("95°S..SCLC")
    End Sub

    Private Sub PictureBox57_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox57.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MDQN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MDQN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("107.5°S..SDQC")
    End Sub

    Private Sub PictureBox56_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox56.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MDUN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MDUN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("93°S..SDUC")
    End Sub

    Private Sub PictureBox55_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox55.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MSKN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MSKN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("75°S..SSKC")
    End Sub

    Private Sub PictureBox54_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox54.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MTFN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MTFN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("90°S..STFC")
    End Sub

    Private Sub PictureBox52_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox52.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..SVJC" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..SVJC" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("93°S..SVJC")
    End Sub

    Private Sub PictureBox51_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox51.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MVUN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MVUN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("93°S..SVUC")
    End Sub

    Private Sub PictureBox50_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox50.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..SVQC" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..SVQC" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("107.5°S..SVQC")
    End Sub

    Private Sub PictureBox49_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox49.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "S..MWLN" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "S..MWLN" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("95°S..SWLC")
    End Sub

    Private Sub PictureBox18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox18.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "SER-L" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "SER-L" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("SE")
    End Sub

    Private Sub PictureBox16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox16.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "SNR-L" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "SNR-L" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("SN")
    End Sub

    Private Sub PictureBox20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox20.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "QA" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "QA" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("QA")
    End Sub

    Private Sub PictureBox33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox33.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "QD" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "QD" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("QD")
    End Sub

    Private Sub PictureBox53_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox53.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "GRE" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "GRE" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("GRE")
    End Sub

    Private Sub PictureBox61_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox61.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "GRI" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "GRI" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("GRI")
    End Sub

    Private Sub PictureBox62_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox62.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "GRV" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "GRV" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("GRV")
    End Sub

    Private Sub PictureBox63_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox63.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "QD_D" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "QD_D" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("QD_D")
    End Sub

    Private Sub PictureBox60_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox60.Click
        If Dir("c:\Program Files\方案文件库\图片\" & "QD_N" & ".jpg", vbNormal) <> "" Then
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\" & "QD_N" & ".jpg")
        Else
            PictureBox9.Image = System.Drawing.Image.FromFile("c:\Program Files\方案文件库\图片\公司图片2.jpg")
        End If
        tools_xuanzhe("QD_N")
    End Sub
End Class

Imports System.Windows.Forms
Imports System.Windows.Forms.Keys
Imports Excel = Microsoft.Office.Interop.Excel


Public Class gongy

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim myExcel As Excel.Application = Nothing  '定义进程

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        myExcel = GetObject(, "Excel.Application")  '打开已经打开的excel程序
        myExcel.Visible = True

        'MsgBox(myExcel.ActiveCell.Row)


        myExcel.ActiveSheet.Cells(myExcel.ActiveCell.Row, 3).Value = RichTextBox1.Text '工件名称


        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub gongy_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim i As Integer
        Dim tfg As String
        Dim fs, f

        Dim sb_ty As String  '加工中心或者数控车控制启动 

        Dim myExcel As Excel.Application = Nothing  '定义进程
        myExcel = GetObject(, "Excel.Application")  '打开已经打开的excel程序
        myExcel.Visible = True

        sb_ty = Mid(myExcel.ActiveSheet.Cells(6, 6).Value, 1, 1)

        ' MsgBox(sb_ty)

        If sb_ty = "M" Or sb_ty = "m" Then
            If Dir("c:\Program Files\方案文件库\常用术语输入.mjj", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\常用术语输入.mjj", 1, 0)
                i = 1

                Do While i = 1
                    tfg = f.readline
                    If tfg = "END" Then i = -1 Else ListBox1.Items.Add(tfg)
                Loop
                f.Close()
            End If
        End If

        If sb_ty = "C" Or sb_ty = "c" Then
            If Dir("c:\Program Files\方案文件库\常用术语输入.cjj", vbNormal) <> "" Then

                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\常用术语输入.cjj", 1, 0)
                i = 1

                Do While i = 1
                    tfg = f.readline
                    If tfg = "END" Then i = -1 Else ListBox1.Items.Add(tfg)
                Loop
                f.Close()
            End If
        End If


        RichTextBox1.Text = myExcel.ActiveSheet.Cells(myExcel.ActiveCell.Row, 3).Value

    End Sub

    Private Sub ListBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.DoubleClick

        Dim MyText1 As String    '存放光标前的文本
        Dim Mytext2 As String    '存放光标后的文本
        MyText1 = RichTextBox1.Text.Substring(0, RichTextBox1.SelectionStart)
        Mytext2 = RichTextBox1.Text.Substring(RichTextBox1.SelectionStart)
        RichTextBox1.Text = MyText1 & ListBox1.SelectedItem.ToString & Mytext2

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, Button9.Click, Button8.Click, Button7.Click, Button6.Click, Button5.Click, Button49.Click, Button48.Click, Button47.Click, Button46.Click, Button45.Click, Button44.Click, Button43.Click, Button42.Click, Button41.Click, Button40.Click, Button4.Click, Button39.Click, Button38.Click, Button37.Click, Button36.Click, Button35.Click, Button34.Click, Button33.Click, Button32.Click, Button31.Click, Button30.Click, Button3.Click, Button29.Click, Button28.Click, Button27.Click, Button26.Click, Button25.Click, Button24.Click, Button23.Click, Button22.Click, Button21.Click, Button20.Click, Button2.Click, Button19.Click, Button18.Click, Button17.Click, Button16.Click, Button15.Click, Button14.Click, Button13.Click, Button12.Click, Button11.Click, Button10.Click
        InsertText(sender)
    End Sub

    Private Sub InsertText(button As Button)
        Dim MyText1 As String    '存放光标前的文本
        Dim Mytext2 As String    '存放光标后的文本
        MyText1 = RichTextBox1.Text.Substring(0, RichTextBox1.SelectionStart)
        Mytext2 = RichTextBox1.Text.Substring(RichTextBox1.SelectionStart)
        RichTextBox1.Text = MyText1 & button.Text & Mytext2
    End Sub

End Class

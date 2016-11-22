Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports System.Management

Public Class OperationTools

    ''' <summary>
    ''' 储存原始列表-用于模糊查询
    ''' </summary>
    Private Shared iniItems As New Collection
    ''' <summary>
    ''' 储存新的列表-用于模糊查询
    ''' </summary>
    Private Shared newItems As New Collection



    Friend Shared Function CheckUser() As Boolean
        '判定配置文件是否合法
        Dim libFullName As String
        Dim i, numOfMac As Integer
        Dim mac, strEnd As String

        Dim fs, f
        numOfMac = 0
        libFullName = ConstData.libPath + "EGNC.lib"
        If Dir(libFullName, vbNormal) <> "" Then
            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile(libFullName, 1, 0)
            i = 1
            Do While i = 1
                mac = f.readline
                If mac.StartsWith("END") Then
                    strEnd = mac.Substring(3, 1)
                    'MsgBox(strEnd)
                    i = -1
                Else
                    numOfMac += 1
                    GlobalData.macs.Add(mac)
                    'MsgBox(mac)
                End If
            Loop
            f.Close()

            'MsgBox(Str(numOfMac + 64) + "::" + Asc(strEnd).ToString)

            If ((numOfMac + 64) = Asc(strEnd)) Then
                'MsgBox("right")
                '获取本地MAC
                Dim localMac As New Collection
                Dim Wmi As New System.Management.ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapterConfiguration")
                For Each WmiObj As Management.ManagementObject In Wmi.Get
                    If CBool(WmiObj("IPEnabled")) Then
                        localMac.Add(WmiObj("MACAddress"))
                    End If
                Next
                For i = 1 To localMac.Count
                    For j = 1 To GlobalData.macs.Count
                        'MsgBox("GlobalData.macs" + j.ToString + GlobalData.macs(j) + "--localMac" + i.ToString + localMac(i))
                        If localMac(i).Equals(GlobalData.macs(j)) Then
                            'MsgBox("MAC匹配")
                            Return True
                        End If
                    Next
                Next
                'MsgBox("MAC错误")
                Return False
            Else
                MsgBox("error")
                Return False
            End If

        End If



            '读取当前MAC

            '检查许可


            Return True
    End Function



    ''' <summary>
    ''' 获取LIB文件中内容并写入下拉框--整行写入
    ''' </summary>
    ''' <param name="libname">存放信息的LIB文件名称及后缀</param>
    ''' <param name="comboBox">要写入的下拉框</param>
    Friend Shared Sub GetLibToComboBox(libName As String, comboBox As ComboBox)
        comboBox.Items.Clear()
        Dim libFullName As String
        Dim i As Integer
        Dim str As String
        Dim fs, f
        libFullName = ConstData.libPath + libName
        If Dir(libFullName, vbNormal) <> "" Then
            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile(libFullName, 1, 0)
            i = 1
            Do While i = 1
                str = f.readline
                If str = "END" Then
                    i = -1
                Else
                    comboBox.Items.Add(tools.zh_biam_f(str.ToString().Trim))
                End If
            Loop
            f.Close()
            'If comboBox.Items.Count <> 0 Then
            '    comboBox.SelectedIndex = 0
            'End If
        End If
    End Sub

    ''' <summary>
    ''' 读取LIB文件的内容并写入集合中--整行写入
    ''' </summary>
    ''' <param name="libName"></param>
    ''' <param name="list"></param>
    Friend Shared Sub GetLibToCollection(libName As String, list As Collection)
        list.Clear()
        Dim libFullName As String
        Dim i As Integer
        Dim str As String
        Dim fs, f
        libFullName = ConstData.libPath + libName
        If Dir(libFullName, vbNormal) <> "" Then
            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile(libFullName, 1, 0)
            i = 1
            Do While i = 1
                str = f.readline
                If str = "END" Then
                    i = -1
                Else
                    list.Add(tools.zh_biam_f(str.ToString().Trim))
                End If
            Loop
            f.Close()
        End If
    End Sub

    ''' <summary>
    ''' 获取LIB文件中内容并写入下拉框--部分写入
    ''' </summary>
    ''' <param name="libName"></param>
    ''' <param name="comboBox"></param>
    ''' <param name="startIndex">读取该行的起始位置</param>
    ''' <param name="length">读取该行的长度，为0则读取到最后</param>
    Friend Shared Sub GetLibToComboBox(libName As String, comboBox As ComboBox, startIndex As Integer, length As Integer)
        Try
            comboBox.Items.Clear()
            Dim libFullName As String
            Dim i As Integer
            Dim str As String
            Dim fs, f
            libFullName = ConstData.libPath + libName
            If Dir(libFullName, vbNormal) <> "" Then
                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile(libFullName, 1, 0)
                i = 1
                Do While i = 1
                    str = f.readline
                    If str = "END" Then
                        i = -1
                    Else
                        If length = 0 Then
                            If str.Length > startIndex Then
                                comboBox.Items.Add(tools.zh_biam_f(str.Substring(startIndex).Trim))
                            End If
                        Else
                            If str.Length > startIndex Then
                                comboBox.Items.Add(tools.zh_biam_f(str.Substring(startIndex, length).Trim))
                            End If
                        End If
                    End If
                Loop
                f.Close()
                'If comboBox.Items.Count <> 0 Then
                '    comboBox.SelectedIndex = 0
                'End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' 获取LIB文件中内容并写入集合--部分写入
    ''' </summary>
    ''' <param name="libName"></param>
    ''' <param name="list"></param>
    ''' <param name="startIndex"></param>
    ''' <param name="length"></param>
    Friend Shared Sub GetLibToCollection(libName As String, list As Collection, startIndex As Integer, length As Integer)
        list.Clear()
        Dim libFullName As String
        Dim i As Integer
        Dim str As String
        Dim fs, f
        libFullName = ConstData.libPath + libName
        If Dir(libFullName, vbNormal) <> "" Then
            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile(libFullName, 1, 0)
            i = 1
            Do While i = 1
                str = f.readline
                If str = "END" Then
                    i = -1
                Else
                    If length = 0 Then
                        If str.Length > startIndex Then
                            list.Add(tools.zh_biam_f(str.Substring(startIndex).Trim))
                        End If
                    Else
                        If str.Length > startIndex Then
                            list.Add(tools.zh_biam_f(str.Substring(startIndex, length).Trim))
                        End If
                    End If
                End If
            Loop
            f.Close()
        End If
    End Sub


    ''' <summary>
    ''' 获取数据库中客户的名称和编码信息，并写入下拉框
    ''' </summary>
    ''' <param name="comboBox1"></param>
    ''' <param name="comboBox2"></param>
    Friend Shared Sub GetCustomInfo(comboBox1 As ComboBox)
        'Try
        '    '定义连接变量,连接字符串也很简单吧,扔个数据库的文件名就行了
        '    Dim conn As Data.SQLite.SQLiteConnection = New Data.SQLite.SQLiteConnection()
        '    '设置连接语句
        '    conn.ConnectionString = "Data Source=C:\VS\基于Excel快速方案选型设计系统\Code\EGNC_ExcelDoc\EGNC_ExcelDoc2\DB\EGNC_ExcelDoc.db" ';Password=egnc123"
        '    'conn.SetPassword(Globals.Sheet1.dbpassword)
        '    '打开连接
        '    conn.Open()
        '    '定义一个执行命令的对象
        '    Dim cmd As SQLite.SQLiteCommand = New SQLite.SQLiteCommand(conn)
        '    '设置SQL命令内容1
        '    cmd.CommandText = "select * from CustomInfo"
        '    '定义一个读取数据的reader对象
        '    Dim reader As SQLite.SQLiteDataReader = cmd.ExecuteReader()
        '    '写入下拉框
        '    While reader.Read
        '        comboBox1.Items.Add(reader("cusname").ToString)
        '        GlobalData.cusNames.Add(reader("cusname").ToString)
        '        'mboBox2.Items.Add(reader("cusid").ToString)
        '        GlobalData.projectID.Add(reader("cusid").ToString)
        '    End While
        '    '关闭reader
        '    reader.Close()
        '    '关闭连接
        '    conn.Close()
        'Catch
        '    MessageBox.Show("资源连接失败，请联系管理员解决！")
        '    Globals.ThisWorkbook.Close()
        'End Try
    End Sub

    ''' <summary>
    ''' 获取机床类型信息并写入下拉框
    ''' </summary>
    ''' <param name="comboBox">机床类型下拉框</param>
    Friend Shared Sub GetMachType(comboBox As ComboBox)
        'Try
        '    '定义连接变量,连接字符串也很简单吧,扔个数据库的文件名就行了
        '    Dim conn As Data.SQLite.SQLiteConnection = New Data.SQLite.SQLiteConnection()
        '    '设置连接语句
        '    conn.ConnectionString = "Data Source=C:\VS\基于Excel快速方案选型设计系统\Code\EGNC_ExcelDoc\EGNC_ExcelDoc2\DB\EGNC_ExcelDoc.db" ';Password=egnc123"
        '    'conn.SetPassword(Globals.Sheet1.dbpassword)
        '    '打开连接
        '    conn.Open()
        '    '定义一个执行命令的对象
        '    Dim cmd As SQLite.SQLiteCommand = New SQLite.SQLiteCommand(conn)
        '    '设置SQL命令内容1
        '    cmd.CommandText = "select * from ToolType"
        '    '定义一个读取数据的reader对象
        '    Dim reader As SQLite.SQLiteDataReader = cmd.ExecuteReader()
        '    '写入下拉框
        '    While reader.Read
        '        comboBox.Items.Add(reader("tooltype").ToString)
        '    End While
        '    '关闭reader
        '    reader.Close()
        '    '关闭连接
        '    conn.Close()
        'Catch
        '    MessageBox.Show("资源连接失败，请联系管理员解决！")
        'End Try
    End Sub

    ''' <summary>
    ''' 执行模糊查询的函数
    ''' </summary>
    ''' <param name="comboBox"></param>
    Friend Shared Sub FuzzyQuery(comboBox As ComboBox, iList As Collection)
        Try
            newItems.Clear()
            iniItems.Clear()
            comboBox.Items.Clear()
            '将已有项目写入集合
            For i = 1 To iList.Count
                iniItems.Add(iList(i).ToString)
            Next
            '验证输入的和已有的是否匹配
            For i = 1 To iniItems.Count
                If comboBox.Text.Equals("") Then
                    newItems.Add(iniItems(i))
                ElseIf iniItems(i).ToString.Contains(comboBox.Text) Then
                    newItems.Add(iniItems(i))
                Else

                End If
            Next
            '将匹配的写入下拉框
            If newItems.Count >= 1 Then
                For i = 1 To newItems.Count
                    comboBox.Items.Add(newItems(i))
                Next
                comboBox.SelectionStart = comboBox.Text.Length
                comboBox.DroppedDown = True
            Else
                comboBox.Items.Add("无匹配项")
                comboBox.SelectionStart = comboBox.Text.Length
                comboBox.DroppedDown = True
            End If


            'Me.Cursor = Cursors.Default
        Catch ex As Exception
            MsgBox(“模糊查询出错：ERROR-” + ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' 按照给定名字查找指定文件夹的图片，并放入PIXBOX容器
    ''' </summary>
    ''' <param name="imageName">图片名字</param>
    ''' <param name="imageDir">存放图片的路径</param>
    ''' <param name="picBox">要放入的图片容器</param>
    Friend Shared Sub ShowImage(imageName As String, imageDir As String, picBox As PictureBox)
        Try
            If imageName <> "" Then
                If Dir(imageDir & imageName & ".jpg", vbNormal) <> "" Then
                    picBox.Image = System.Drawing.Image.FromFile(imageDir & imageName & ".jpg")
                Else
                    'picBox.Image = System.Drawing.Image.FromFile(ConstData.logoImagePath)
                    picBox.Image = Nothing
                End If
            Else
                picBox.Image = Nothing
            End If
        Catch ex As Exception
            MsgBox("显示图片出错：ERROR-" + ex.Message)
        End Try
    End Sub


    ''' <summary>
    ''' 根据工具系统名称和接口等读取工具系统附件型号并写入下拉框
    ''' </summary>
    ''' <param name="gjmc">工具名称</param>
    ''' <param name="gjxh">工具型号</param>
    ''' <param name="fjxh">附件型号</param>
    ''' <param name="fjmc">附件名称</param>
    ''' <param name="jkcc">接口尺寸</param>
    Friend Shared Sub GetHolderAttachToCombBox(gjmc As ComboBox, gjxh As ComboBox, fjxh As ComboBox, fjmc As ComboBox, jkcc As TextBox)
        Dim i As Integer
        Dim stxf As String
        Dim fs, f

        Dim gs_km_cb As Integer

        Dim Tool_ass As String

        ' Dim tooling_insert_k As String '刀片型号
        'gjxh.Items.Add(Mid(tools.zh_biam_f(stxf), 23, Len(tools.zh_biam_f(stxf)) - 22 - 12))

        If Dir("c:\Program Files\方案文件库\" & gjmc.Text & ".lib", vbNormal) <> "" Then
            fs = CreateObject("Scripting.FileSystemObject")
            f = fs.OpenTextFile("c:\Program Files\方案文件库\" & gjmc.Text & ".lib", 1, 0)
            i = 1
            stxf = f.readline
            Do While i = 1
                stxf = f.readline
                If stxf = "END" Then
                    i = -1
                Else


                    If Mid(tools.zh_biam_f(stxf), 23, Len(tools.zh_biam_f(stxf)) - 22 - 12) = gjxh.Text Then
                        Tool_ass = Mid(tools.zh_biam_f(stxf), 1, 4)
                        i = -1
                    End If
                End If
            Loop
            f.Close()
        End If

        fjxh.Items.Clear()


        gs_km_cb = 0
        If fjmc.Text = "攻丝卡簧" Then gs_km_cb = 1

        If gs_km_cb = 0 Then

            '通用夹持
            If Dir("c:\Program Files\方案文件库\" & fjmc.Text & ".lib", vbNormal) <> "" Then
                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\" & fjmc.Text & ".lib", 1, 0)
                i = 1

                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else


                        If Val(Tool_ass) = Val(Mid(tools.zh_biam_f(stxf), 1, 4)) And Val(jkcc.Text) >= Val(Mid(tools.zh_biam_f(stxf), 5, 6)) And Val(jkcc.Text) <= Val(Mid(tools.zh_biam_f(stxf), 11, 6)) Then
                            fjxh.Items.Add(Mid(tools.zh_biam_f(stxf), 17, Len(tools.zh_biam_f(stxf)) - 16))
                        End If
                    End If
                Loop
                f.Close()
            End If


            If fjxh.Items.Count <> 0 Then fjxh.SelectedIndex = 0
        End If


        If gs_km_cb = 1 Then
            '攻丝卡簧夹持
            If Dir("c:\Program Files\方案文件库\" & fjmc.Text & ".lib", vbNormal) <> "" Then
                fs = CreateObject("Scripting.FileSystemObject")
                f = fs.OpenTextFile("c:\Program Files\方案文件库\" & fjmc.Text & ".lib", 1, 0)
                i = 1

                Do While i = 1
                    stxf = f.readline
                    If stxf = "END" Then
                        i = -1
                    Else
                        ' MsgBox(Mid(tools.zh_biam_f(stxf), 5, 6))
                        ' MsgBox(Mid(tools.zh_biam_f(stxf), 11, 6))

                        If Val(Tool_ass) = Val(Mid(tools.zh_biam_f(stxf), 1, 4)) And Val(jkcc.Text) = Val(Mid(tools.zh_biam_f(stxf), 5, 6)) And tools.t_yx_long = Val(Mid(tools.zh_biam_f(stxf), 11, 6)) Then
                            fjxh.Items.Add(Mid(tools.zh_biam_f(stxf), 17, Len(tools.zh_biam_f(stxf)) - 16))
                        End If
                    End If
                Loop
                f.Close()
            End If


            If fjxh.Items.Count <> 0 Then fjxh.SelectedIndex = 0
        End If
    End Sub

    ''' <summary>
    ''' 根据输入的非标刀具型号提取需要的刀具型号码
    ''' </summary>
    ''' <param name="toolType">输入的非标刀具型号</param>
    Friend Shared Function GetToolTypeRef(toolType As String) As String
        'a 对应的 ascii 码 为97     z 对应的 ascii 码 为122
        'A 对应的 ascii 码 为65     Z 对应的 ascii 码 为90
        Dim c As String
        Dim toolTypeRef As String = ""
        If toolType <> "" Then
            For i = 1 To Len(toolType)
                c = Mid(toolType, i, 1)
                If Asc(c) > 65 And Asc(c) < 90 Or Asc(c) > 97 And Asc(c) < 122 Then
                    ''是字母，则输出显示
                    toolTypeRef = toolTypeRef + c
                Else

                End If
            Next
        End If
        GetToolTypeRef = toolTypeRef
    End Function

    ''' <summary>
    ''' 调整行数
    ''' </summary>
    ''' <param name="rowsNow">现有行数</param>
    ''' <param name="rowsWant">需要的行数</param>
    ''' <param name="rowActivite">目前选中的行</param>
    ''' <param name="myExcel">表格对象</param>
    Friend Shared Sub AdjustRows(rowsNow As Integer, rowsWant As Integer, rowActivite As Integer, myExcel As Excel.Application)
        If rowsWant > 4 And rowsNow < rowsWant Then
            '插入行
            For i = 1 To rowsWant - rowsNow
                myExcel.Rows(rowActivite + 1).Insert
                'myExcel.Rows(rowActivite + 1).Insert(, CopyOrigin:=myExcel.xlFormatFromLeftOrAbove)
                'myExcel.ActiveSheet.Cells(rowActivite + 1, GlobalData.colActivite).EntireRow.Insert(, CopyOrigin:=myExcel.xlFormatFromLeftOrAbove)
                'MsgBox("插入行")
            Next
        End If
        If rowsNow > 4 And rowsWant < rowsNow Then
            '删除行，保留4行
            If rowsWant <= 4 Then
                For i = 1 To rowsNow - 4
                    myExcel.Rows(rowActivite + 1).Delete
                    'myExcel.ActiveSheet.Cells(rowActivite + i, GlobalData.colActivite).EntireRow.Delete
                    'MsgBox("A删除行")
                Next
            Else
                For i = 1 To rowsNow - rowsWant
                    myExcel.Rows(rowActivite + 1).Delete
                    'myExcel.ActiveSheet.Cells(rowActivite + i, GlobalData.colActivite).EntireRow.Delete
                    'MsgBox("B删除行")
                Next
            End If
        End If
    End Sub
End Class

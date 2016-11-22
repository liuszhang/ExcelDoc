
''' <summary>
''' 存放从配置文件或者系统读取的全局数据类
''' </summary>
Public Class GlobalData

    ''' <summary>
    ''' 要写入表格的行数
    ''' </summary>
    Friend Shared rowsToExcel As Integer
    ''' <summary>
    ''' 此刀号目前已有的行数
    ''' </summary>
    Friend Shared rowsNow As Integer
    ''' <summary>
    ''' 刀号（T01,T02,T03等）
    ''' </summary>
    Friend Shared toolNum As String
    ''' <summary>
    ''' 当前双击选中的行和列
    ''' </summary>
    Friend Shared rowActivite As Integer
    Friend Shared colActivite As Integer
    ''' <summary>
    ''' 需要插入的图片
    ''' </summary>
    Friend Shared picToExcel As String



    ''' <summary>
    ''' 客户名称信息
    ''' </summary>
    Friend Shared cusNames As New Collection
    ''' <summary>
    ''' 项目编号-暂时不用
    ''' </summary>
    Friend Shared projectID As New Collection
    ''' <summary>
    ''' 选定的设备类型
    ''' </summary>
    Friend Shared machType As String
    ''' <summary>
    ''' 是否内冷
    ''' </summary>
    Friend Shared isInCold As Boolean
    ''' <summary>
    ''' 拉钉型号
    ''' </summary>
    Friend Shared nailType As New Collection
    ''' <summary>
    ''' 选定的拉钉型号
    ''' </summary>
    Friend Shared nail As String
    ''' <summary>
    ''' 材料类型
    ''' </summary>
    Friend Shared materialType As New Collection
    ''' <summary>
    ''' 选定的材料类型
    ''' </summary>
    Friend Shared material As String


    ''' <summary>
    ''' 刀片型号
    ''' </summary>
    Friend Shared bladeType As New Collection
    ''' <summary>
    ''' 刀片型号（精简）
    ''' </summary>
    Friend Shared bladeTypeShort As New Collection
    ''' <summary>
    ''' 特定材料刀片型号
    ''' </summary>
    Friend Shared bladeTypeForMaterial As New Collection


    ''' <summary>
    ''' 工具系统型号
    ''' </summary>
    Friend Shared holderType As New Collection
    ''' <summary>
    ''' 工具系统型号(精简)
    ''' </summary>
    Friend Shared holderTypeShort As New Collection
    ''' <summary>
    ''' 特定材料工具系统型号
    ''' </summary>
    Friend Shared holderTypeForMaterial As New Collection


    ''' <summary>
    ''' 绑定的MAC地址
    ''' </summary>
    Friend Shared macs As New Collection


End Class

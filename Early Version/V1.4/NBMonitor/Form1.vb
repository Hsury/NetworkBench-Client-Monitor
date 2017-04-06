Public Class Form1

    Dim WshShell As Object
    Dim TimeNow As Date
    Dim Log As String
    Dim NBStartTime As String
    Dim NBRunTime As Long
    Dim NBYear, NBMonth, NBDay, NBHour, NBMinute, NBSecond As Integer
    Dim NBClient, NBRT, NBET, NBDT, NBWT, NBUT, NBUFT As Integer
    Dim NBsConns, NBsCPU, NBsDisk, NBsMem, NBsPageFile, NBsPing As Integer
    Dim NBCity, NBISP, NBSpeed As Integer
    Dim NBAutoLogin, NBAutoStart As Integer
    Dim ErrTime As Integer
    Dim Initial, S1, S2, S3, S4, S5, S6, S7, S8, S9, S10, S11, S12, S13 As Integer

    Private Sub WriteLog()
        On Error Resume Next
        TimeNow = Now
        If TextBox1.Text = "" Then
            TextBox1.Text = Format(TimeNow, "HH:mm:ss") & " " & Log & TextBox1.Text
        Else
            TextBox1.Text = Format(TimeNow, "HH:mm:ss") & " " & Log & vbCrLf & TextBox1.Text
        End If
        IO.File.AppendAllText("NBMonitor_Logs\" & Format(TimeNow, "yyyy年M月d日") & ".txt", Format(TimeNow, "HH:mm:ss") & " " & Log & vbCrLf)
    End Sub

    Private Sub GetNBService()
        On Error Resume Next
        If NBStartTime <> "" Then
            NBYear = Strings.Left(NBStartTime, 4)
            NBMonth = Strings.Right(Strings.Left(NBStartTime, 6), 2)
            NBDay = Strings.Right(Strings.Left(NBStartTime, 8), 2)
            NBHour = Strings.Right(Strings.Left(NBStartTime, 10), 2)
            NBMinute = Strings.Right(Strings.Left(NBStartTime, 12), 2)
            NBSecond = Strings.Right(NBStartTime, 2)
            NBRunTime = DateDiff("s", CDate(NBYear & "/" & NBMonth & "/" & NBDay & " " & NBHour & ":" & NBMinute & ":" & NBSecond), Now)
            Label1.ForeColor = Color.LimeGreen
            Label1.Text = "服务已启动"
            Label2.Text = "启动日期:" & NBYear & "年" & NBMonth & "月" & NBDay & "日 "
            Label3.Text = "启动时间:" & NBHour & "时" & NBMinute & "分" & NBSecond & "秒"
            Label4.Text = "运行时长:" & Int(NBRunTime / 86400) & "天" & Int((NBRunTime - Int(NBRunTime / 86400) * 86400) / 3600) & "小时" & Int((NBRunTime - Int(NBRunTime / 3600) * 3600) / 60) & "分"
            If S1 <> 1 Then
                Log = "[服务] 检测到服务已启动"
                WriteLog()
                Log = "[服务] 启动时间:" & NBYear & "年" & NBMonth & "月" & NBDay & "日 " & NBHour & "时" & NBMinute & "分" & NBSecond & "秒"
                WriteLog()
                Log = "[服务] 已运行时长:" & Int(NBRunTime / 86400) & "天" & Int((NBRunTime - Int(NBRunTime / 86400) * 86400) / 3600) & "小时" & Int((NBRunTime - Int(NBRunTime / 3600) * 3600) / 60) & "分"
                WriteLog()
                S1 = 1
            End If
        Else
            For Each Wmi In GetObject("winmgmts:\\.\root\cimv2:win32_process").instances_
                If Wmi.Name = "NBService.exe" Then NBStartTime = Strings.Left(Wmi.CreationDate, 14)
            Next
            If NBStartTime = "" Then
                NBRunTime = 0
                Label1.ForeColor = Color.OrangeRed
                Label1.Text = "服务未启动"
                Label2.Text = "启动日期:未启动"
                Label3.Text = "启动时间:未启动"
                Label4.Text = "运行时长:未启动"
                If S1 <> 2 Then
                    Log = "[服务] 检测到服务未启动"
                    WriteLog()
                    S1 = 2
                End If
            Else
                GetNBService()
            End If
        End If
    End Sub

    Private Sub GetNBTask()
        On Error Resume Next
        If WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\status") = 5 Then
            If S2 <> 1 Then
                NBStartTime = ""
                NBRT = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\rt")
                NBDT = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\dt")
                NBUT = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\ut")
                NBWT = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\wt")
                NBET = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\et")
                NBUFT = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\uft")
                Label5.ForeColor = Color.LimeGreen
                If NBET <> 0 Then
                    Label10.ForeColor = Color.DodgerBlue
                Else
                    Label10.ForeColor = SystemColors.ControlText
                End If
                If NBUFT <> 0 Then
                    Label11.ForeColor = Color.DeepSkyBlue
                Else
                    Label11.ForeColor = SystemColors.ControlText
                End If
                Label5.Text = "正在执行任务"
                Label6.Text = "请求任务数:" & NBRT
                Label7.Text = "完成任务数:" & NBDT
                Label8.Text = "上传任务数:" & NBUT
                Label9.Text = "等待任务数:" & NBWT
                Label10.Text = "超时任务数:" & NBET
                Label11.Text = "上传失败任务数:" & NBUFT
                Log = "[任务] 正在执行任务"
                WriteLog()
                Log = "[任务] 请求:" & NBRT & " 超时:" & NBET & " 完成:" & NBDT & " 等待:" & NBWT & " 上传:" & NBUT & " 上传失败:" & NBUFT
                WriteLog()
                S2 = 1
            Else
                If WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\rt") > NBRT Then
                    Log = "[任务] 请求任务 增加了" & WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\rt") - NBRT & "个"
                    WriteLog()
                    NBRT = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\rt")
                    Label6.Text = "请求任务数:" & NBRT
                    S3 = 1
                End If
                If WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\dt") > NBDT Then
                    Log = "[任务] 完成任务 增加了" & WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\dt") - NBDT & "个"
                    WriteLog()
                    NBDT = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\dt")
                    Label7.Text = "完成任务数:" & NBDT
                    S3 = 1
                End If
                If WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\ut") > NBUT Then
                    Log = "[任务] 上传任务 增加了" & WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\ut") - NBUT & "个"
                    WriteLog()
                    NBUT = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\ut")
                    Label8.Text = "上传任务数:" & NBUT
                    S3 = 1
                End If
                If WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\wt") > NBWT Then
                    Log = "[任务] 等待任务 增加了" & WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\wt") - NBWT & "个"
                    WriteLog()
                    NBWT = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\wt")
                    Label9.Text = "等待任务数:" & NBWT
                    S3 = 1
                End If
                If WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\wt") < NBWT Then
                    Log = "[任务] 等待任务 减少了" & NBWT - WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\wt") & "个"
                    WriteLog()
                    NBWT = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\wt")
                    Label9.Text = "等待任务数:" & NBWT
                    S3 = 1
                End If
                If WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\et") > NBET Then
                    Log = "[任务] 超时任务 增加了" & WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\et") - NBET & "个"
                    WriteLog()
                    NBET = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\et")
                    Label10.Text = "超时任务数:" & NBET
                    S3 = 1
                End If
                If WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\uft") > NBUFT Then
                    Log = "[任务] 上传失败任务 增加了" & WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\uft") - NBUFT & "个"
                    WriteLog()
                    NBUFT = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\uft")
                    Label11.Text = "上传失败任务数:" & NBUFT
                    S3 = 1
                End If
                If WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\uft") < NBUFT Then
                    Log = "[任务] 上传失败任务 减少了" & NBUFT - WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\uft") & "个"
                    WriteLog()
                    NBUFT = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\uft")
                    Label11.Text = "上传失败任务数:" & NBUFT
                    S3 = 1
                End If
                If NBET <> 0 Then
                    Label10.ForeColor = Color.DodgerBlue
                Else
                    Label10.ForeColor = SystemColors.ControlText
                End If
                If NBUFT <> 0 Then
                    Label11.ForeColor = Color.DeepSkyBlue
                Else
                    Label11.ForeColor = SystemColors.ControlText
                End If
                If S3 = 1 Then
                    Log = "[任务] 请求:" & NBRT & " 完成:" & NBDT & " 上传:" & NBUT & " 等待:" & NBWT & " 超时:" & NBET & " 上传失败:" & NBUFT
                    WriteLog()
                    S3 = 0
                End If
                If WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\rt") < NBRT Then
                    Timer2.Enabled = False
                    Log = "[服务] 服务状态出现异常,正在重启服务"
                    WriteLog()
                    ServiceController1.Stop()
                    ServiceController1.Start()
                    Log = "[服务] 服务已重启"
                    WriteLog()
                    NBStartTime = ""
                    S1 = 0
                    S2 = 0
                    Timer2.Enabled = True
                End If
            End If
        Else
            If S2 <> 2 Then
                NBRT = 0
                NBDT = 0
                NBUT = 0
                NBWT = 0
                NBET = 0
                NBUFT = 0
                Label5.ForeColor = Color.OrangeRed
                Label10.ForeColor = SystemColors.ControlText
                Label11.ForeColor = SystemColors.ControlText
                Label5.Text = "未在执行任务"
                Label6.Text = "请求任务数:" & NBRT
                Label7.Text = "完成任务数:" & NBDT
                Label8.Text = "上传任务数:" & NBUT
                Label9.Text = "等待任务数:" & NBWT
                Label10.Text = "超时任务数:" & NBET
                Label11.Text = "上传失败任务数:" & NBUFT
                Log = "[任务] 未在执行任务"
                WriteLog()
                S2 = 2
            End If
            NBStartTime = ""
        End If
    End Sub

    Private Sub GetNBStatus()
        On Error Resume Next
        NBsCPU = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\sCPU")
        NBsMem = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\sMem")
        NBsPing = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\sPing")
        NBsConns = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\sConns")
        NBsPageFile = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\sPageFile")
        NBsDisk = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\sDisk")
        If NBsCPU Then
            Label13.ForeColor = Color.Orange
            Label13.Text = "CPU使用率:异常"
            S4 = 0
            If S5 = 0 Then
                Log = "[系统] CPU使用率出现异常"
                WriteLog()
                S5 = 1
                ErrTime = ErrTime + 1
            End If
        Else
            Label13.ForeColor = SystemColors.ControlText
            Label13.Text = "CPU使用率:正常"
            If S5 = 1 Then
                Log = "[系统] CPU使用率恢复正常"
                WriteLog()
                S5 = 0
            End If
        End If
        If NBsMem Then
            Label14.ForeColor = Color.Orange
            Label14.Text = "内存使用率:异常"
            S4 = 0
            If S6 = 0 Then
                Log = "[系统] 内存使用率出现异常"
                WriteLog()
                S6 = 1
                ErrTime = ErrTime + 1
            End If
        Else
            Label14.ForeColor = SystemColors.ControlText
            Label14.Text = "内存使用率:正常"
            If S6 = 1 Then
                Log = "[系统] 内存使用率恢复正常"
                WriteLog()
                S6 = 0
            End If
        End If
        If NBsPing Then
            Label15.ForeColor = Color.Orange
            Label15.Text = "网络延迟值:异常"
            S4 = 0
            If S7 = 0 Then
                Log = "[系统] 网络延迟值出现异常"
                WriteLog()
                S7 = 1
                ErrTime = ErrTime + 1
            End If
        Else
            Label15.ForeColor = SystemColors.ControlText
            Label15.Text = "网络延迟值:正常"
            If S7 = 1 Then
                Log = "[系统] 网络延迟值恢复正常"
                WriteLog()
                S7 = 0
            End If
        End If
        If NBsConns Then
            Label16.ForeColor = Color.Orange
            Label16.Text = "并发连接数:异常"
            S4 = 0
            If S8 = 0 Then
                Log = "[系统] 并发连接数出现异常"
                WriteLog()
                S8 = 1
                ErrTime = ErrTime + 1
            End If
        Else
            Label16.ForeColor = SystemColors.ControlText
            Label16.Text = "并发连接数:正常"
            If S8 = 1 Then
                Log = "[系统] 并发连接数恢复正常"
                WriteLog()
                S8 = 0
            End If
        End If
        If NBsPageFile Then
            Label17.ForeColor = Color.Orange
            Label17.Text = "系统虚拟内存:异常"
            S4 = 0
            If S9 = 0 Then
                Log = "[系统] 系统虚拟内存出现异常"
                WriteLog()
                S9 = 1
                ErrTime = ErrTime + 1
            End If
        Else
            Label17.ForeColor = SystemColors.ControlText
            Label17.Text = "系统虚拟内存:正常"
            If S9 = 1 Then
                Log = "[系统] 系统虚拟内存恢复正常"
                WriteLog()
                S9 = 0
            End If
        End If
        If NBsDisk Then
            Label18.ForeColor = Color.Orange
            Label18.Text = "硬盘可用空间:异常"
            S4 = 0
            If S10 = 0 Then
                Log = "[系统] 硬盘可用空间出现异常"
                WriteLog()
                S10 = 1
                ErrTime = ErrTime + 1
            End If
        Else
            Label18.ForeColor = SystemColors.ControlText
            Label18.Text = "硬盘可用空间:正常"
            If S10 = 1 Then
                Log = "[系统] 硬盘可用空间恢复正常"
                WriteLog()
                S10 = 0
            End If
        End If
        If NBsCPU Or NBsMem Or NBsPing Or NBsConns Or NBsPageFile Or NBsDisk Then
            Label12.ForeColor = Color.OrangeRed
            Label12.Text = "系统状态不佳"
        Else
            Label12.ForeColor = Color.LimeGreen
            Label12.Text = "系统状态良好"
            If S4 = 0 Then
                Log = "[系统] 系统状态各方面正常"
                WriteLog()
                S4 = 1
            End If
        End If
    End Sub

    Private Sub GetNBNetwork()
        On Error Resume Next
        If (NBCity <> WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\City")) Or (NBISP <> WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\Isp")) Or (NBSpeed <> WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\Speed")) Then
            S11 = 1
        End If
        NBCity = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\City")
        NBISP = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\Isp")
        NBSpeed = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\Speed")
        Select Case Int(NBCity / 100)
            Case 4811 : Label20.Text = "省份:北京"
            Case 4812 : Label20.Text = "省份:天津"
            Case 4813 : Label20.Text = "省份:河北省"
            Case 4814 : Label20.Text = "省份:山西省"
            Case 4815 : Label20.Text = "省份:内蒙古自治区"
            Case 4821 : Label20.Text = "省份:辽宁省"
            Case 4822 : Label20.Text = "省份:吉林省"
            Case 4823 : Label20.Text = "省份:黑龙江省"
            Case 4831 : Label20.Text = "省份:上海"
            Case 4832 : Label20.Text = "省份:江苏省"
            Case 4833 : Label20.Text = "省份:浙江省"
            Case 4834 : Label20.Text = "省份:安徽省"
            Case 4835 : Label20.Text = "省份:福建省"
            Case 4836 : Label20.Text = "省份:江西省"
            Case 4837 : Label20.Text = "省份:山东省"
            Case 4841 : Label20.Text = "省份:河南省"
            Case 4842 : Label20.Text = "省份:湖北省"
            Case 4843 : Label20.Text = "省份:湖南省"
            Case 4844 : Label20.Text = "省份:广东省"
            Case 4845 : Label20.Text = "省份:广西壮族自治区"
            Case 4846 : Label20.Text = "省份:海南省"
            Case 4850 : Label20.Text = "省份:重庆"
            Case 4851 : Label20.Text = "省份:四川省"
            Case 4852 : Label20.Text = "省份:贵州省"
            Case 4853 : Label20.Text = "省份:云南省"
            Case 4854 : Label20.Text = "省份:西藏自治区"
            Case 4861 : Label20.Text = "省份:陕西省"
            Case 4862 : Label20.Text = "省份:甘肃"
            Case 4863 : Label20.Text = "省份:青海"
            Case 4864 : Label20.Text = "省份:宁夏回族自治区"
            Case 4865 : Label20.Text = "省份:新疆维吾尔自治区"
            Case 4871 : Label20.Text = "省份:台湾省"
            Case 4881 : Label20.Text = "省份:香港特别行政区"
            Case 4882 : Label20.Text = "省份:澳门特别行政区"
            Case Else : Label20.Text = "省份:未知"
        End Select
        Select Case NBCity
            Case 481101 : Label21.Text = "城市:北京市"
            Case 481201 : Label21.Text = "城市:天津市"
            Case 481301 : Label21.Text = "城市:石家庄市"
            Case 481302 : Label21.Text = "城市:唐山市"
            Case 481303 : Label21.Text = "城市:秦皇岛市"
            Case 481304 : Label21.Text = "城市:邯郸市"
            Case 481305 : Label21.Text = "城市:邢台市"
            Case 481306 : Label21.Text = "城市:保定市"
            Case 481307 : Label21.Text = "城市:张家口市"
            Case 481308 : Label21.Text = "城市:承德市"
            Case 481309 : Label21.Text = "城市:沧州市"
            Case 481310 : Label21.Text = "城市:廊坊市"
            Case 481311 : Label21.Text = "城市:衡水市"
            Case 481401 : Label21.Text = "城市:太原市"
            Case 481402 : Label21.Text = "城市:大同市"
            Case 481403 : Label21.Text = "城市:阳泉市"
            Case 481404 : Label21.Text = "城市:长治市"
            Case 481405 : Label21.Text = "城市:晋城市"
            Case 481406 : Label21.Text = "城市:朔州市"
            Case 481407 : Label21.Text = "城市:晋中市"
            Case 481408 : Label21.Text = "城市:运城市"
            Case 481409 : Label21.Text = "城市:忻州市"
            Case 481410 : Label21.Text = "城市:临汾市"
            Case 481411 : Label21.Text = "城市:吕梁市"
            Case 481501 : Label21.Text = "城市:呼和浩特市"
            Case 481502 : Label21.Text = "城市:包头市"
            Case 481503 : Label21.Text = "城市:乌海市"
            Case 481504 : Label21.Text = "城市:赤峰市"
            Case 481505 : Label21.Text = "城市:通辽市"
            Case 481506 : Label21.Text = "城市:鄂尔多斯市"
            Case 481507 : Label21.Text = "城市:呼伦贝尔市"
            Case 481508 : Label21.Text = "城市:巴彦淖尔市"
            Case 481509 : Label21.Text = "城市:乌兰察布市"
            Case 481522 : Label21.Text = "城市:兴安盟"
            Case 481525 : Label21.Text = "城市:锡林郭勒盟"
            Case 481529 : Label21.Text = "城市:阿拉善盟"
            Case 482101 : Label21.Text = "城市:沈阳市"
            Case 482102 : Label21.Text = "城市:大连市"
            Case 482103 : Label21.Text = "城市:鞍山市"
            Case 482104 : Label21.Text = "城市:抚顺市"
            Case 482105 : Label21.Text = "城市:本溪市"
            Case 482106 : Label21.Text = "城市:丹东市"
            Case 482107 : Label21.Text = "城市:锦州市"
            Case 482108 : Label21.Text = "城市:营口市"
            Case 482109 : Label21.Text = "城市:阜新市"
            Case 482110 : Label21.Text = "城市:辽阳市"
            Case 482111 : Label21.Text = "城市:盘锦市"
            Case 482112 : Label21.Text = "城市:铁岭市"
            Case 482113 : Label21.Text = "城市:朝阳市"
            Case 482114 : Label21.Text = "城市:葫芦岛市"
            Case 482201 : Label21.Text = "城市:长春市"
            Case 482202 : Label21.Text = "城市:吉林市"
            Case 482203 : Label21.Text = "城市:四平市"
            Case 482204 : Label21.Text = "城市:辽源市"
            Case 482205 : Label21.Text = "城市:通化市"
            Case 482206 : Label21.Text = "城市:白山市"
            Case 482207 : Label21.Text = "城市:松原市"
            Case 482208 : Label21.Text = "城市:白城市"
            Case 482224 : Label21.Text = "城市:延边朝鲜族自治州"
            Case 482301 : Label21.Text = "城市:哈尔滨市"
            Case 482302 : Label21.Text = "城市:齐齐哈尔市"
            Case 482303 : Label21.Text = "城市:鸡西市"
            Case 482304 : Label21.Text = "城市:鹤岗市"
            Case 482305 : Label21.Text = "城市:双鸭山市"
            Case 482306 : Label21.Text = "城市:大庆市"
            Case 482307 : Label21.Text = "城市:伊春市"
            Case 482308 : Label21.Text = "城市:佳木斯市"
            Case 482309 : Label21.Text = "城市:七台河市"
            Case 482310 : Label21.Text = "城市:牡丹江市"
            Case 482311 : Label21.Text = "城市:黑河市"
            Case 482312 : Label21.Text = "城市:绥化市"
            Case 482327 : Label21.Text = "城市:大兴安岭地区"
            Case 483101 : Label21.Text = "城市:上海市"
            Case 483201 : Label21.Text = "城市:南京市"
            Case 483202 : Label21.Text = "城市:无锡市"
            Case 483203 : Label21.Text = "城市:徐州市"
            Case 483204 : Label21.Text = "城市:常州市"
            Case 483205 : Label21.Text = "城市:苏州市"
            Case 483206 : Label21.Text = "城市:南通市"
            Case 483207 : Label21.Text = "城市:连云港市"
            Case 483208 : Label21.Text = "城市:淮安市"
            Case 483209 : Label21.Text = "城市:盐城市"
            Case 483210 : Label21.Text = "城市:扬州市"
            Case 483211 : Label21.Text = "城市:镇江市"
            Case 483212 : Label21.Text = "城市:泰州市"
            Case 483213 : Label21.Text = "城市:宿迁市"
            Case 483301 : Label21.Text = "城市:杭州市"
            Case 483302 : Label21.Text = "城市:宁波市"
            Case 483303 : Label21.Text = "城市:温州市"
            Case 483304 : Label21.Text = "城市:嘉兴市"
            Case 483305 : Label21.Text = "城市:湖州市"
            Case 483306 : Label21.Text = "城市:绍兴市"
            Case 483307 : Label21.Text = "城市:金华市"
            Case 483308 : Label21.Text = "城市:衢州市"
            Case 483309 : Label21.Text = "城市:舟山市"
            Case 483310 : Label21.Text = "城市:台州市"
            Case 483311 : Label21.Text = "城市:丽水市"
            Case 483401 : Label21.Text = "城市:合肥市"
            Case 483402 : Label21.Text = "城市:芜湖市"
            Case 483403 : Label21.Text = "城市:蚌埠市"
            Case 483404 : Label21.Text = "城市:淮南市"
            Case 483405 : Label21.Text = "城市:马鞍山市"
            Case 483406 : Label21.Text = "城市:淮北市"
            Case 483407 : Label21.Text = "城市:铜陵市"
            Case 483408 : Label21.Text = "城市:安庆市"
            Case 483410 : Label21.Text = "城市:黄山市"
            Case 483411 : Label21.Text = "城市:滁州市"
            Case 483412 : Label21.Text = "城市:阜阳市"
            Case 483413 : Label21.Text = "城市:宿州市"
            Case 483414 : Label21.Text = "城市:巢湖市"
            Case 483415 : Label21.Text = "城市:六安市"
            Case 483416 : Label21.Text = "城市:亳州市"
            Case 483417 : Label21.Text = "城市:池州市"
            Case 483418 : Label21.Text = "城市:宣城市"
            Case 483501 : Label21.Text = "城市:福州市"
            Case 483502 : Label21.Text = "城市:厦门市"
            Case 483503 : Label21.Text = "城市:莆田市"
            Case 483504 : Label21.Text = "城市:三明市"
            Case 483505 : Label21.Text = "城市:泉州市"
            Case 483506 : Label21.Text = "城市:漳州市"
            Case 483507 : Label21.Text = "城市:南平市"
            Case 483508 : Label21.Text = "城市:龙岩市"
            Case 483509 : Label21.Text = "城市:宁德市"
            Case 483601 : Label21.Text = "城市:南昌市"
            Case 483602 : Label21.Text = "城市:景德镇市"
            Case 483603 : Label21.Text = "城市:萍乡市"
            Case 483604 : Label21.Text = "城市:九江市"
            Case 483605 : Label21.Text = "城市:新余市"
            Case 483606 : Label21.Text = "城市:鹰潭市"
            Case 483607 : Label21.Text = "城市:赣州市"
            Case 483608 : Label21.Text = "城市:吉安市"
            Case 483609 : Label21.Text = "城市:宜春市"
            Case 483610 : Label21.Text = "城市:抚州市"
            Case 483611 : Label21.Text = "城市:上饶市"
            Case 483701 : Label21.Text = "城市:济南市"
            Case 483702 : Label21.Text = "城市:青岛市"
            Case 483703 : Label21.Text = "城市:淄博市"
            Case 483704 : Label21.Text = "城市:枣庄市"
            Case 483705 : Label21.Text = "城市:东营市"
            Case 483706 : Label21.Text = "城市:烟台市"
            Case 483707 : Label21.Text = "城市:潍坊市"
            Case 483708 : Label21.Text = "城市:济宁市"
            Case 483709 : Label21.Text = "城市:泰安市"
            Case 483710 : Label21.Text = "城市:威海市"
            Case 483711 : Label21.Text = "城市:日照市"
            Case 483712 : Label21.Text = "城市:莱芜市"
            Case 483713 : Label21.Text = "城市:临沂市"
            Case 483714 : Label21.Text = "城市:德州市"
            Case 483715 : Label21.Text = "城市:聊城市"
            Case 483716 : Label21.Text = "城市:滨州市"
            Case 483717 : Label21.Text = "城市:菏泽市"
            Case 484101 : Label21.Text = "城市:郑州市"
            Case 484102 : Label21.Text = "城市:开封市"
            Case 484103 : Label21.Text = "城市:洛阳市"
            Case 484104 : Label21.Text = "城市:平顶山市"
            Case 484105 : Label21.Text = "城市:安阳市"
            Case 484106 : Label21.Text = "城市:鹤壁市"
            Case 484107 : Label21.Text = "城市:新乡市"
            Case 484108 : Label21.Text = "城市:焦作市"
            Case 484109 : Label21.Text = "城市:濮阳市"
            Case 484110 : Label21.Text = "城市:许昌市"
            Case 484111 : Label21.Text = "城市:漯河市"
            Case 484112 : Label21.Text = "城市:三门峡市"
            Case 484113 : Label21.Text = "城市:南阳市"
            Case 484114 : Label21.Text = "城市:商丘市"
            Case 484115 : Label21.Text = "城市:信阳市"
            Case 484116 : Label21.Text = "城市:周口市"
            Case 484117 : Label21.Text = "城市:驻马店市"
            Case 484118 : Label21.Text = "城市:济源市"
            Case 484201 : Label21.Text = "城市:武汉市"
            Case 484202 : Label21.Text = "城市:黄石市"
            Case 484203 : Label21.Text = "城市:十堰市"
            Case 484205 : Label21.Text = "城市:宜昌市"
            Case 484206 : Label21.Text = "城市:襄樊市"
            Case 484207 : Label21.Text = "城市:鄂州市"
            Case 484208 : Label21.Text = "城市:荆门市"
            Case 484209 : Label21.Text = "城市:孝感市"
            Case 484210 : Label21.Text = "城市:荆州市"
            Case 484211 : Label21.Text = "城市:黄冈市"
            Case 484212 : Label21.Text = "城市:咸宁市"
            Case 484213 : Label21.Text = "城市:随州市"
            Case 484228 : Label21.Text = "城市:恩施土家族苗族州"
            Case 484291 : Label21.Text = "城市:神农架林区"
            Case 484294 : Label21.Text = "城市:仙桃市"
            Case 484295 : Label21.Text = "城市:潜江市"
            Case 484296 : Label21.Text = "城市:天门市"
            Case 484301 : Label21.Text = "城市:长沙市"
            Case 484302 : Label21.Text = "城市:株洲市"
            Case 484303 : Label21.Text = "城市:湘潭市"
            Case 484304 : Label21.Text = "城市:衡阳市"
            Case 484305 : Label21.Text = "城市:邵阳市"
            Case 484306 : Label21.Text = "城市:岳阳市"
            Case 484307 : Label21.Text = "城市:常德市"
            Case 484308 : Label21.Text = "城市:张家界市"
            Case 484309 : Label21.Text = "城市:益阳市"
            Case 484310 : Label21.Text = "城市:郴州市"
            Case 484311 : Label21.Text = "城市:永州市"
            Case 484312 : Label21.Text = "城市:怀化市"
            Case 484313 : Label21.Text = "城市:娄底市"
            Case 484331 : Label21.Text = "城市:湘西土家族苗族州"
            Case 484401 : Label21.Text = "城市:广州市"
            Case 484402 : Label21.Text = "城市:韶关市"
            Case 484403 : Label21.Text = "城市:深圳市"
            Case 484404 : Label21.Text = "城市:珠海市"
            Case 484405 : Label21.Text = "城市:汕头市"
            Case 484406 : Label21.Text = "城市:佛山市"
            Case 484407 : Label21.Text = "城市:江门市"
            Case 484408 : Label21.Text = "城市:湛江市"
            Case 484409 : Label21.Text = "城市:茂名市"
            Case 484412 : Label21.Text = "城市:肇庆市"
            Case 484413 : Label21.Text = "城市:惠州市"
            Case 484414 : Label21.Text = "城市:梅州市"
            Case 484415 : Label21.Text = "城市:汕尾市"
            Case 484416 : Label21.Text = "城市:河源市"
            Case 484417 : Label21.Text = "城市:阳江市"
            Case 484418 : Label21.Text = "城市:清远市"
            Case 484419 : Label21.Text = "城市:东莞市"
            Case 484420 : Label21.Text = "城市:中山市"
            Case 484451 : Label21.Text = "城市:潮州市"
            Case 484452 : Label21.Text = "城市:揭阳市"
            Case 484453 : Label21.Text = "城市:云浮市"
            Case 484501 : Label21.Text = "城市:南宁市"
            Case 484502 : Label21.Text = "城市:柳州市"
            Case 484503 : Label21.Text = "城市:桂林市"
            Case 484504 : Label21.Text = "城市:梧州市"
            Case 484505 : Label21.Text = "城市:北海市"
            Case 484506 : Label21.Text = "城市:防城港市"
            Case 484507 : Label21.Text = "城市:钦州市"
            Case 484508 : Label21.Text = "城市:贵港市"
            Case 484509 : Label21.Text = "城市:玉林市"
            Case 484510 : Label21.Text = "城市:百色市"
            Case 484511 : Label21.Text = "城市:贺州市"
            Case 484512 : Label21.Text = "城市:河池市"
            Case 484513 : Label21.Text = "城市:来宾市"
            Case 484514 : Label21.Text = "城市:崇左市"
            Case 484601 : Label21.Text = "城市:海口市"
            Case 484602 : Label21.Text = "城市:三亚市"
            Case 485001 : Label21.Text = "城市:重庆市"
            Case 485101 : Label21.Text = "城市:成都市"
            Case 485103 : Label21.Text = "城市:自贡市"
            Case 485104 : Label21.Text = "城市:攀枝花市"
            Case 485105 : Label21.Text = "城市:泸州市"
            Case 485106 : Label21.Text = "城市:德阳市"
            Case 485107 : Label21.Text = "城市:绵阳市"
            Case 485108 : Label21.Text = "城市:广元市"
            Case 485109 : Label21.Text = "城市:遂宁市"
            Case 485110 : Label21.Text = "城市:内江市"
            Case 485111 : Label21.Text = "城市:乐山市"
            Case 485113 : Label21.Text = "城市:南充市"
            Case 485114 : Label21.Text = "城市:眉山市"
            Case 485115 : Label21.Text = "城市:宜宾市"
            Case 485116 : Label21.Text = "城市:广安市"
            Case 485117 : Label21.Text = "城市:达州市"
            Case 485118 : Label21.Text = "城市:雅安市"
            Case 485119 : Label21.Text = "城市:巴中市"
            Case 485120 : Label21.Text = "城市:资阳市"
            Case 485132 : Label21.Text = "城市:阿坝藏族羌族州"
            Case 485133 : Label21.Text = "城市:甘孜藏族自治州"
            Case 485134 : Label21.Text = "城市:凉山彝族自治州"
            Case 485201 : Label21.Text = "城市:贵阳市"
            Case 485202 : Label21.Text = "城市:六盘水市"
            Case 485203 : Label21.Text = "城市:遵义市"
            Case 485204 : Label21.Text = "城市:安顺市"
            Case 485222 : Label21.Text = "城市:铜仁地区"
            Case 485223 : Label21.Text = "城市:黔西南布依族苗族州"
            Case 485224 : Label21.Text = "城市:毕节地区"
            Case 485226 : Label21.Text = "城市:黔东南苗族侗族州"
            Case 485227 : Label21.Text = "城市:黔南布依族苗族州"
            Case 485301 : Label21.Text = "城市:昆明市"
            Case 485303 : Label21.Text = "城市:曲靖市"
            Case 485304 : Label21.Text = "城市:玉溪市"
            Case 485305 : Label21.Text = "城市:保山市"
            Case 485306 : Label21.Text = "城市:昭通市"
            Case 485307 : Label21.Text = "城市:丽江市"
            Case 485308 : Label21.Text = "城市:思茅市"
            Case 485309 : Label21.Text = "城市:临沧市"
            Case 485323 : Label21.Text = "城市:楚雄彝族自治州"
            Case 485325 : Label21.Text = "城市:红河哈尼族彝族州"
            Case 485326 : Label21.Text = "城市:文山壮族苗族州"
            Case 485328 : Label21.Text = "城市:西双版纳傣族州"
            Case 485329 : Label21.Text = "城市:大理白族自治州"
            Case 485331 : Label21.Text = "城市:德宏傣族景颇族州"
            Case 485333 : Label21.Text = "城市:怒江傈僳族自治州"
            Case 485334 : Label21.Text = "城市:迪庆藏族自治州"
            Case 485401 : Label21.Text = "城市:拉萨市"
            Case 485421 : Label21.Text = "城市:昌都地区"
            Case 485422 : Label21.Text = "城市:山南地区"
            Case 485423 : Label21.Text = "城市:日喀则地区"
            Case 485424 : Label21.Text = "城市:那曲地区"
            Case 485425 : Label21.Text = "城市:阿里地区"
            Case 485426 : Label21.Text = "城市:林芝地区"
            Case 486101 : Label21.Text = "城市:西安市"
            Case 486102 : Label21.Text = "城市:铜川市"
            Case 486103 : Label21.Text = "城市:宝鸡市"
            Case 486104 : Label21.Text = "城市:咸阳市"
            Case 486105 : Label21.Text = "城市:渭南市"
            Case 486106 : Label21.Text = "城市:延安市"
            Case 486107 : Label21.Text = "城市:汉中市"
            Case 486108 : Label21.Text = "城市:榆林市"
            Case 486109 : Label21.Text = "城市:安康市"
            Case 486110 : Label21.Text = "城市:商洛市"
            Case 486201 : Label21.Text = "城市:兰州市"
            Case 486202 : Label21.Text = "城市:嘉峪关市"
            Case 486203 : Label21.Text = "城市:金昌市"
            Case 486204 : Label21.Text = "城市:白银市"
            Case 486205 : Label21.Text = "城市:天水市"
            Case 486206 : Label21.Text = "城市:武威市"
            Case 486207 : Label21.Text = "城市:张掖市"
            Case 486208 : Label21.Text = "城市:平凉市"
            Case 486209 : Label21.Text = "城市:酒泉市"
            Case 486210 : Label21.Text = "城市:庆阳市"
            Case 486211 : Label21.Text = "城市:定西市"
            Case 486212 : Label21.Text = "城市:陇南市"
            Case 486229 : Label21.Text = "城市:临夏回族自治州"
            Case 486230 : Label21.Text = "城市:甘南藏族自治州"
            Case 486301 : Label21.Text = "城市:西宁市"
            Case 486321 : Label21.Text = "城市:海东地区"
            Case 486322 : Label21.Text = "城市:海北藏族自治州"
            Case 486323 : Label21.Text = "城市:黄南藏族自治州"
            Case 486325 : Label21.Text = "城市:海南藏族自治州"
            Case 486326 : Label21.Text = "城市:果洛藏族自治州"
            Case 486327 : Label21.Text = "城市:玉树藏族自治州"
            Case 486328 : Label21.Text = "城市:海西蒙古族藏族州"
            Case 486401 : Label21.Text = "城市:银川市"
            Case 486402 : Label21.Text = "城市:石嘴山市"
            Case 486403 : Label21.Text = "城市:吴忠市"
            Case 486404 : Label21.Text = "城市:固原市"
            Case 486405 : Label21.Text = "城市:中卫市"
            Case 486501 : Label21.Text = "城市:乌鲁木齐市"
            Case 486502 : Label21.Text = "城市:克拉玛依市"
            Case 486521 : Label21.Text = "城市:吐鲁番地区"
            Case 486522 : Label21.Text = "城市:哈密地区"
            Case 486523 : Label21.Text = "城市:昌吉回族自治州"
            Case 486527 : Label21.Text = "城市:博尔塔拉蒙古州"
            Case 486528 : Label21.Text = "城市:巴音郭楞蒙古州"
            Case 486529 : Label21.Text = "城市:阿克苏地区"
            Case 486530 : Label21.Text = "城市:克孜勒苏柯尔克孜州"
            Case 486531 : Label21.Text = "城市:喀什地区"
            Case 486532 : Label21.Text = "城市:和田地区"
            Case 486540 : Label21.Text = "城市:伊犁哈萨克州"
            Case 486542 : Label21.Text = "城市:塔城地区"
            Case 486543 : Label21.Text = "城市:阿勒泰地区"
            Case 486591 : Label21.Text = "城市:石河子市"
            Case 486592 : Label21.Text = "城市:阿拉尔市"
            Case 486593 : Label21.Text = "城市:图木舒克市"
            Case 486594 : Label21.Text = "城市:五家渠市"
            Case 487101 : Label21.Text = "城市:台北市"
            Case 487102 : Label21.Text = "城市:新竹市"
            Case 487103 : Label21.Text = "城市:台中市"
            Case 487104 : Label21.Text = "城市:台南市"
            Case 487105 : Label21.Text = "城市:高雄市"
            Case 488101 : Label21.Text = "城市:香港"
            Case 488201 : Label21.Text = "城市:澳门"
            Case Else : Label21.Text = "城市:未知"
        End Select
        Select Case NBISP
            Case 16 : Label22.Text = "ISP:中国联通"
            Case 17 : Label22.Text = "ISP:中国电信"
            Case 18 : Label22.Text = "ISP:电信通"
            Case 19 : Label22.Text = "ISP:中国吉通"
            Case 20 : Label22.Text = "ISP:中国铁通"
            Case 21 : Label22.Text = "ISP:世纪互联"
            Case 22 : Label22.Text = "ISP:中电华通"
            Case 23 : Label22.Text = "ISP:有线通"
            Case 24 : Label22.Text = "ISP:电通"
            Case 25 : Label22.Text = "ISP:中国移动"
            Case 26 : Label22.Text = "ISP:联通(将弃用)"
            Case 27 : Label22.Text = "ISP:教育网"
            Case 28 : Label22.Text = "ISP:中电飞华"
            Case 29 : Label22.Text = "ISP:恒敦通信"
            Case 30 : Label22.Text = "ISP:石油通信"
            Case 31 : Label22.Text = "ISP:同辉通信"
            Case 32 : Label22.Text = "ISP:润迅通信"
            Case 33 : Label22.Text = "ISP:图像数据通信"
            Case 34 : Label22.Text = "ISP:长丰通信"
            Case 35 : Label22.Text = "ISP:比林通信"
            Case 36 : Label22.Text = "ISP:航天通信"
            Case 37 : Label22.Text = "ISP:国安创想通信"
            Case 38 : Label22.Text = "ISP:神州通信"
            Case 39 : Label22.Text = "ISP:畅捷通信"
            Case 40 : Label22.Text = "ISP:中关村通信"
            Case 41 : Label22.Text = "ISP:光通通信"
            Case 42 : Label22.Text = "ISP:辽河油田电信"
            Case 43 : Label22.Text = "ISP:中吉电信"
            Case 44 : Label22.Text = "ISP:东风通信"
            Case 45 : Label22.Text = "ISP:和记环讯(香港)"
            Case 46 : Label22.Text = "ISP:蝉凌通信"
            Case 47 : Label22.Text = "ISP:Concord通信"
            Case 48 : Label22.Text = "ISP:天地通电信"
            Case 49 : Label22.Text = "ISP:互联通"
            Case 50 : Label22.Text = "ISP:大唐电信"
            Case 51 : Label22.Text = "ISP:博路电信"
            Case 52 : Label22.Text = "ISP:信天通信"
            Case 53 : Label22.Text = "ISP:油田电信"
            Case 54 : Label22.Text = "ISP:恒汇通信"
            Case 55 : Label22.Text = "ISP:大庆油田通信"
            Case 56 : Label22.Text = "ISP:中华通信"
            Case 57 : Label22.Text = "ISP:北京CBD电信"
            Case 58 : Label22.Text = "ISP:安莱信息通信"
            Case 59 : Label22.Text = "ISP:中基电信"
            Case 60 : Label22.Text = "ISP:润科通信"
            Case 61 : Label22.Text = "ISP:长城宽带"
            Case 63 : Label22.Text = "ISP:方正宽带"
            Case 64 : Label22.Text = "ISP:广电宽带"
            Case 65 : Label22.Text = "ISP:平安保险"
            Case 66 : Label22.Text = "ISP:长鸿宽带"
            Case 67 : Label22.Text = "ISP:信息港宽带"
            Case 68 : Label22.Text = "ISP:万通宽带"
            Case 69 : Label22.Text = "ISP:远传电信"
            Case 70 : Label22.Text = "ISP:CPCNet"
            Case 71 : Label22.Text = "ISP:帝联科技"
            Case 72 : Label22.Text = "ISP:E家宽带"
            Case 74 : Label22.Text = "ISP:蓝波宽带"
            Case 75 : Label22.Text = "ISP:蓝讯通信技术"
            Case 76 : Label22.Text = "ISP:天威宽带"
            Case 77 : Label22.Text = "ISP:第一线(香港)"
            Case 78 : Label22.Text = "ISP:油田宽带"
            Case 79 : Label22.Text = "ISP:视讯宽带"
            Case 80 : Label22.Text = "ISP:歌华有线宽带"
            Case 81 : Label22.Text = "ISP:楹联宽带"
            Case 82 : Label22.Text = "ISP:长丰宽带"
            Case 83 : Label22.Text = "ISP:视通宽带"
            Case 84 : Label22.Text = "ISP:中华宽带"
            Case 85 : Label22.Text = "ISP:百灵宽带"
            Case 86 : Label22.Text = "ISP:华宇宽带"
            Case 87 : Label22.Text = "ISP:中海宽带"
            Case 88 : Label22.Text = "ISP:中华电信"
            Case 89 : Label22.Text = "ISP:光环新网"
            Case 90 : Label22.Text = "ISP:慧聪网络"
            Case 91 : Label22.Text = "ISP:景安网络"
            Case 92 : Label22.Text = "ISP:263网络"
            Case 93 : Label22.Text = "ISP:网宿科技"
            Case 94 : Label22.Text = "ISP:HKNet"
            Case 95 : Label22.Text = "ISP:城市电讯(香港)"
            Case 96 : Label22.Text = "ISP:电讯盈科"
            Case 97 : Label22.Text = "ISP:索尼SONET网络"
            Case 98 : Label22.Text = "ISP:阿里巴巴"
            Case 99 : Label22.Text = "ISP:百度网络"
            Case 100 : Label22.Text = "ISP:谷歌网络"
            Case 101 : Label22.Text = "ISP:东方网景"
            Case 102 : Label22.Text = "ISP:飞华领航"
            Case 103 : Label22.Text = "ISP:国研网"
            Case 104 : Label22.Text = "ISP:金桥网"
            Case 106 : Label22.Text = "ISP:神州在线"
            Case 107 : Label22.Text = "ISP:世导信息"
            Case 108 : Label22.Text = "ISP:首创网络"
            Case 109 : Label22.Text = "ISP:首信网"
            Case 110 : Label22.Text = "ISP:统计信息网"
            Case 111 : Label22.Text = "ISP:网联无限"
            Case 112 : Label22.Text = "ISP:新一代数据中心"
            Case 113 : Label22.Text = "ISP:中国工程技术信息网"
            Case 114 : Label22.Text = "ISP:中国科技网"
            Case 116 : Label22.Text = "ISP:中经网"
            Case 117 : Label22.Text = "ISP:中信网络"
            Case 118 : Label22.Text = "ISP:网联光通"
            Case 119 : Label22.Text = "ISP:CTM INTERNET SERVICES"
            Case 120 : Label22.Text = "ISP:华数网通"
            Case 121 : Label22.Text = "ISP:企商在线"
            Case 200 : Label22.Text = "ISP:阿联酋电信"
            Case 201 : Label22.Text = "ISP:NTT"
            Case 202 : Label22.Text = "ISP:Akamai Technologies"
            Case 203 : Label22.Text = "ISP:Hurricane Electric"
            Case 204 : Label22.Text = "ISP:奎斯特通讯"
            Case 205 : Label22.Text = "ISP:NTT America"
            Case 206 : Label22.Text = "ISP:Tiscali International Network B.V."
            Case 207 : Label22.Text = "ISP:NLAYER COMMUNICATIONS"
            Case 208 : Label22.Text = "ISP:Telus Communications"
            Case 209 : Label22.Text = "ISP:Server Central Network"
            Case 210 : Label22.Text = "ISP:Teleglobe"
            Case 211 : Label22.Text = "ISP:法国电信"
            Case 212 : Label22.Text = "ISP:Level 3 Communications"
            Case 213 : Label22.Text = "ISP:Global Crossings"
            Case 214 : Label22.Text = "ISP:Acanac"
            Case 215 : Label22.Text = "ISP:Telstra"
            Case 216 : Label22.Text = "ISP:Optus"
            Case 217 : Label22.Text = "ISP:Hanaro"
            Case 218 : Label22.Text = "ISP:Kt"
            Case 219 : Label22.Text = "ISP:Dacom"
            Case 220 : Label22.Text = "ISP:StarHub"
            Case 221 : Label22.Text = "ISP:Bell"
            Case 222 : Label22.Text = "ISP:British Telecom"
            Case 223 : Label22.Text = "ISP:SINGNET"
            Case 224 : Label22.Text = "ISP:KDDI"
            Case 225 : Label22.Text = "ISP:AOL"
            Case 226 : Label22.Text = "ISP:AT & T"
            Case 227 : Label22.Text = "ISP:Cingular"
            Case 228 : Label22.Text = "ISP:Sprint"
            Case 229 : Label22.Text = "ISP:Verizon"
            Case 230 : Label22.Text = "ISP:YAHOO JP"
            Case 231 : Label22.Text = "ISP:Softbank"
            Case 232 : Label22.Text = "ISP:Willcom"
            Case 233 : Label22.Text = "ISP:SK Telecom"
            Case 234 : Label22.Text = "ISP:LG Telecom"
            Case 235 : Label22.Text = "ISP:KTF"
            Case 236 : Label22.Text = "ISP:3 UK"
            Case 237 : Label22.Text = "ISP:Virgin Mobile"
            Case 238 : Label22.Text = "ISP:Vodafone"
            Case 239 : Label22.Text = "ISP:Orange"
            Case 240 : Label22.Text = "ISP:O2"
            Case 241 : Label22.Text = "ISP:T-Mobile"
            Case 242 : Label22.Text = "ISP:SFR"
            Case 243 : Label22.Text = "ISP:DeutscheTelekom"
            Case 244 : Label22.Text = "ISP:E-Plus"
            Case 245 : Label22.Text = "ISP:MobileCom"
            Case 246 : Label22.Text = "ISP:MobileOne"
            Case 247 : Label22.Text = "ISP:MAGNET"
            Case 248 : Label22.Text = "ISP:Rogers"
            Case 249 : Label22.Text = "ISP:TPG"
            Case 250 : Label22.Text = "ISP:FPT Telecom"
            Case 251 : Label22.Text = "ISP:Viettel"
            Case 252 : Label22.Text = "ISP:VNPT/VDC"
            Case 255 : Label22.Text = "ISP:General ISP"
            Case Else : Label22.Text = "ISP:未知"
        End Select
        Select Case NBSpeed
            Case 5 : Label23.Text = "速度:普通MODEM或ISDN"
            Case 6 : Label23.Text = "速度:ADSL或小区宽带"
            Case 7 : Label23.Text = "速度:企业专线"
            Case 8 : Label23.Text = "速度:IDC机房专线"
            Case Else : Label23.Text = "速度:未知"
        End Select
        If S11 = 1 Then
            Log = "[网络] " & Label20.Text & " " & Label21.Text & " " & Label22.Text & " " & Label23.Text
            WriteLog()
            S11 = 0
        End If
        Label19.ForeColor = Color.LimeGreen
        Label19.Text = "网络环境检测完毕"
    End Sub

    Private Sub GetNBData()
        On Error Resume Next
        If NBRunTime <> 0 Then
            Label24.Text = "每分钟可完成任务数:" & Math.Round(NBDT / (NBRunTime / 60), 1)
            Label25.Text = "每小时可完成任务数:" & Math.Round(NBDT / (NBRunTime / 3600), 1)
            Label26.Text = "每天可完成任务数:" & Math.Round(NBDT / (NBRunTime / 86400), 1)
        Else
            Label24.Text = "每分钟可完成任务数:0"
            Label25.Text = "每小时可完成任务数:0"
            Label26.Text = "每天可完成任务数:0"
        End If
        Label27.Text = "系统状态异常次数:" & ErrTime
    End Sub

    Private Sub GetNBConfig()
        On Error Resume Next
        NBAutoLogin = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\AutoLogin")
        NBAutoStart = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\AutoStart")
        If NBAutoLogin = 1 Then
            CheckBox1.Checked = True
        Else
            CheckBox1.Checked = False
        End If
        If NBAutoStart = 1 Then
            CheckBox2.Checked = False
        Else
            CheckBox2.Checked = True
        End If
        If Label1.Text = "服务已启动" Then
            Button1.Text = "停止任务"
        Else
            Button1.Text = "启动任务"
        End If
        CheckBox1.Enabled = True
        CheckBox2.Enabled = True
        Button1.Enabled = True
        S12 = 1
    End Sub

    Private Sub GetNBConsole()
        On Error Resume Next
        ComboBox1.SelectedItem = "1"
        ComboBox1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
        S13 = 1
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        On Error Resume Next
        Select Case Initial
            Case 0
                Initial = Initial + 1
            Case 1
                GetNBService()
                Initial = Initial + 1
            Case 2
                GetNBTask()
                Initial = Initial + 1
            Case 3
                GetNBStatus()
                Initial = Initial + 1
            Case 4
                GetNBNetwork()
                Initial = Initial + 1
            Case 5
                GetNBData()
                Initial = Initial + 1
            Case 6
                GetNBConfig()
                Initial = Initial + 1
            Case 7
                GetNBConsole()
                Timer1.Enabled = False
                Timer2.Enabled = True
        End Select
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        On Error Resume Next
        GetNBService()
        GetNBTask()
        GetNBStatus()
        GetNBNetwork()
        GetNBData()
        GetNBConfig()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        On Error Resume Next
        If Button1.Text = "启动任务" Then
            Timer2.Enabled = False
            ServiceController1.Start()
            NBStartTime = ""
            S1 = 0
            S2 = 0
            Log = "[客户端] 用户手动启动了监测任务"
            WriteLog()
            Button1.Text = "停止任务"
            Timer2.Enabled = True
        Else
            Timer2.Enabled = False
            ServiceController1.Stop()
            NBStartTime = ""
            S1 = 0
            S2 = 0
            Log = "[客户端] 用户手动停止了监测任务"
            WriteLog()
            Button1.Text = "启动任务"
            Timer2.Enabled = True
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        On Error Resume Next
        TextBox1.Text = ""
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        On Error Resume Next
        If Button3.Text = "暂停监控" Then
            Timer2.Enabled = False
            Button3.Text = "继续监控"
            Log = "[控制台] 监控被暂停"
            WriteLog()
        Else
            Timer2.Enabled = True
            Button3.Text = "暂停监控"
            Log = "[控制台] 监控继续运行"
            WriteLog()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        On Error Resume Next
        Shell("explorer.exe " & "NBMonitor_Logs\" & Format(TimeNow, "yyyy年M月d日") & ".txt", vbNormalFocus)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        On Error Resume Next
        Me.Close()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        On Error Resume Next
        If S12 = 1 Then
            If CheckBox1.Checked = True Then
                WshShell.RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\AutoLogin", 1, "REG_DWORD")
                Log = "[客户端] 已设置客户端为自动登录模式"
                WriteLog()
            Else
                WshShell.RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\AutoLogin", 0, "REG_DWORD")
                Log = "[客户端] 已设置客户端为不自动登录模式"
                WriteLog()
            End If
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        On Error Resume Next
        If S12 = 1 Then
            If CheckBox2.Checked = True Then
                WshShell.RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\AutoStart", 0, "REG_DWORD")
                Log = "[客户端] 已设置客户端为隐藏图标模式"
                WriteLog()
            Else
                WshShell.RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\AutoStart", 1, "REG_DWORD")
                Log = "[客户端] 已设置客户端为不隐藏图标模式"
                WriteLog()
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        On Error Resume Next
        If S13 = 1 Then
            Timer2.Interval = Val(ComboBox1.Text) * 1000
            Log = "[控制台] 状态刷新频率调整为" & ComboBox1.Text & "秒"
            WriteLog()
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        On Error Resume Next
        Shell("explorer.exe http://www.kagamiz.com/", vbNormalFocus)
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        On Error Resume Next
        Log = "[信息] 监控程序退出"
        WriteLog()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        WshShell = CreateObject("WScript.Shell")
        IO.Directory.CreateDirectory("NBMonitor_Logs")
        Log = "[信息] 欢迎使用 基调网络监测客户端状态监控工具 V1.4"
        WriteLog()
        Log = "[信息] 此版本编译于:2015年7月20日 23:00"
        WriteLog()
        Timer1.Enabled = True
    End Sub

End Class

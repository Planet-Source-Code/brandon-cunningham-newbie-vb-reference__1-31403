VERSION 5.00
Begin VB.UserControl DMDiskInfo 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "DMDisk.ctx":0000
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "DMDisk.ctx":0842
End
Attribute VB_Name = "DMDiskInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+-----------------------------------------------------------------------------------------------------------------
' Hi this is a simple little Active Control to help with getting disk information
' I know a lot of you want use all theres functions so this was many made for people thinking
' of writeing there own install programs and so i thought this whould come in handy
' and give some good information about disksizes , disktype and other stuff. if you find this control
' to be usfull them please let me know anyway have fun
'+------------------------------------------------------------------------------------------------------------------
' Ben Jones
' Mail: Dreamvb@yahoo.com
' Website1 : dreamvb.s5.com
' website2 : http://www.lastsale.com/vb
'+-----------------------------------------------------------------------------------------------------------------

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Sub About()
    ' Displays about box
    frmAbout.Show vbModal
    
End Sub
Private Sub UserControl_Resize()
    ' This just keeps the Control to 480 x 480 and stops it form beening resized
    UserControl.Size 480, 480
    
End Sub
Public Function DMGetVolumeName(lzDriveLetter As String) As String
' This function will return the volume name of the entered drive
Dim StrVolName As String
    If Len(lzDriveLetter) = 0 Then
        DMGetVolumeName = "ERORR"
        Exit Function
    ElseIf Len(Dir(lzDriveLetter)) <= 0 Then
        DMGetVolumeName = "ERROR"
        Exit Function
    Else
        StrVolName = String(255, Chr(0))
        GetVolumeInformation lzDriveLetter, StrVolName, 255, 0, 0, 0, 0, 255
        DMGetVolumeName = Left(StrVolName, InStr(1, StrVolName, Chr(0)) - 1)
    End If
    
End Function
Public Function DMGetDriveType(lzDriveLetter As String) As Integer
    ' This will return the drive type by responding with a number
    If Len(lzDriveLetter) <= 0 Then
        Exit Function
    Else
      DMGetDriveType = GetDriveType(lzDriveLetter)
    End If
End Function

Public Function DMGetSerialNumber(lzDriveLetter As String) As Long
' This will return the disk serial number of the entered drive
Dim SerNum As Long
    If Len(lzDriveLetter) <= 0 Then
        DMGetSerialNumber = 0
        Exit Function
    ElseIf Len(Dir(lzDriveLetter)) <= 0 Then
        DMGetSerialNumber = 0
        Exit Function
    Else
        GetVolumeInformation lzDriveLetter, 0, 255, SerNum, 0, 0, 0, 255
        If SerNum = 0 Then
            DMGetSerialNumber = 0
            Exit Function
        Else
            DMGetSerialNumber = SerNum
            ' You can format this if you wish useing the hex function
            ' just change the above line to DMGetSerialNumber = hex(SerNum)
            ' then you will need to change the return on the function to return a string value
        End If
    End If
    
End Function

Public Function DMGetFileSysType(lzDriveLetter As String) As String
' This returns the file system type based on the drive letter entered
Dim SysFileType As String
    If Len(lzDriveLetter) <= 0 Then
        DMGetFileSysType = "ERROR"
        Exit Function
    ElseIf Len(Dir(lzDriveLetter)) <= 0 Then
        DMGetFileSysType = "ERORR"
        Exit Function
    Else
        SysFileType = String(255, Chr(0))
        GetVolumeInformation lzDriveLetter, 0, 255, 0, 0, 0, SysFileType, 255
        DMGetFileSysType = Left(SysFileType, InStr(1, SysFileType, Chr(0)) - 1)
    End If
        SysFileType = ""
        
End Function

Public Function DMGetBytesPerSector(lzDriveLetter As String) As Long
' This will return the Bytes per cluster on a entered drive
Dim SectorsPerCluster As Long, BytesPerSector As Long, NFreeClusters As Long, _
TNumOfClusters As Long

    If Len(lzDriveLetter) <= 0 Then
        DMGetBytesPerSector = 0
        Exit Function
    ElseIf Len(Dir(lzDriveLetter)) <= 0 Then
        DMGetBytesPerSector = 0
        Exit Function
    Else
        GetDiskFreeSpace lzDriveLetter, SectorsPerCluster, BytesPerSector, NFreeClusters, TNumOfClusters
        DMGetBytesPerSector = BytesPerSector
    End If
    
End Function

Public Function DMGetSectorsPerCluster(lzDriveLetter As String) As Long
' This will return the Sectors per cluster on a entered drive
Dim SectorsPerCluster As Long, BytesPerSector As Long, NFreeClusters As Long, _
TNumOfClusters As Long
    If Len(lzDriveLetter) <= 0 Then
        DMGetSectorsPerCluster = 0
        Exit Function
    ElseIf Len(Dir(lzDriveLetter)) <= 0 Then
        DMGetSectorsPerCluster = 0
        Exit Function
    Else
        GetDiskFreeSpace lzDriveLetter, SectorsPerCluster, BytesPerSector, NFreeClusters, TNumOfClusters
        DMGetSectorsPerCluster = SectorsPerCluster
    End If

End Function

Public Function DMGetTotalNummerOfFreeClusters(lzDriveLetter As String) As Long
' This will return the Total number of free cluster on a entered drive
Dim SectorsPerCluster As Long, BytesPerSector As Long, NFreeClusters As Long, _
TNumOfClusters As Long
    If Len(lzDriveLetter) <= 0 Then
        DMGetTotalNummerOfFreeClusters = 0
        Exit Function
    ElseIf Len(Dir(lzDriveLetter)) <= 0 Then
        DMGetTotalNummerOfFreeClusters = 0
        Exit Function
    Else
        GetDiskFreeSpace lzDriveLetter, SectorsPerCluster, BytesPerSector, NFreeClusters, TNumOfClusters
        DMGetTotalNummerOfFreeClusters = TNumOfClusters
    End If

End Function

Public Function DMGetFreeClusters(lzDriveLetter As String) As Long
' This will return the number of free clusters on a entered drive
Dim SectorsPerCluster As Long, BytesPerSector As Long, NFreeClusters As Long, _
TNumOfClusters As Long
    If Len(lzDriveLetter) <= 0 Then
        DMGetFreeClusters = 0
        Exit Function
    ElseIf Len(Dir(lzDriveLetter)) <= 0 Then
        DMGetFreeClusters = 0
        Exit Function
    Else
        GetDiskFreeSpace lzDriveLetter, SectorsPerCluster, BytesPerSector, NFreeClusters, TNumOfClusters
        DMGetFreeClusters = NFreeClusters
    End If

End Function

Public Function DMGetFreeDiskSpace(lzDriveLetter As String) As Long
' This will return the total amount of free disk space on a entered drive
Dim SectorsPerCluster As Long, BytesPerSector As Long, NFreeClusters As Long, _
TNumOfClusters As Long
    If Len(lzDriveLetter) <= 0 Then
        DMGetFreeDiskSpace = 0
        Exit Function
    ElseIf Len(Dir(lzDriveLetter)) <= 0 Then
        DMGetFreeDiskSpace = 0
        Exit Function
    Else
        GetDiskFreeSpace lzDriveLetter, SectorsPerCluster, BytesPerSector, NFreeClusters, TNumOfClusters
        DMGetFreeDiskSpace = SectorsPerCluster * BytesPerSector * NFreeClusters
    End If

End Function

Public Function DMGetUsedDiskSpace(lzDriveLetter As String) As Long
' This will return the total amount of used free disk space on a entered drive
Dim SectorsPerCluster As Long, BytesPerSector As Long, NFreeClusters As Long

    If Len(lzDriveLetter) <= 0 Then
        DMGetUsedDiskSpace = 0
        Exit Function
    ElseIf Len(Dir(lzDriveLetter)) <= 0 Then
        DMGetUsedDiskSpace = 0
        Exit Function
    Else
        GetDiskFreeSpace lzDriveLetter, SectorsPerCluster, BytesPerSector, NFreeClusters, TNumOfClusters
        Capacity = SectorsPerCluster * BytesPerSector * TNumOfClusters
        FreeSpace = SectorsPerCluster * BytesPerSector * NFreeClusters
        DMGetUsedDiskSpace = Capacity - FreeSpace
    End If

End Function

Public Function DMGetDiskCapacity(lzDriveLetter As String) As Variant
' This will return the total amount of Disk Capacity used on a entered drive
Dim SectorsPerCluster As Long, BytesPerSector As Long, NFreeClusters As Long

    If Len(lzDriveLetter) <= 0 Then
        DMGetDiskCapacity = 0
        Exit Function
    ElseIf Len(Dir(lzDriveLetter)) <= 0 Then
        DMGetDiskCapacity = 0
        Exit Function
    Else
        GetDiskFreeSpace lzDriveLetter, SectorsPerCluster, BytesPerSector, NFreeClusters, TNumOfClusters
        Capacity = SectorsPerCluster * BytesPerSector * TNumOfClusters
        DMGetDiskCapacity = Capacity
    End If
    
End Function

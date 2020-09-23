<div align="center">

## cmdFormatDrive


</div>

### Description

Format Floppy Disk using API:Here is the code on How to Format Floppy Disk using API. Note -- This code can format your Hard Disk as well, so you should be careful!!!!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Duncan Diep](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/duncan-diep.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/duncan-diep-cmdformatdrive__1-68/archive/master.zip)

### API Declarations

```
Private Declare Function SHFormatDrive Lib "shell32" _
  (ByVal hwnd As Long, ByVal Drive As Long, ByVal fmtID As Long, _
  ByVal options As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias _
  "GetDriveTypeA" (ByVal nDrive As String) As Long
```


### Source Code

```
Add 2 command buttons named :
cmdFormat and cmdDiskCopy
Private Sub cmdFormatDrive_Click()
  Dim DriveLetter$, DriveNumber&, DriveType&
  Dim RetVal&, RetFromMsg%
  DriveLetter = UCase(Drive1.Drive)
  DriveNumber = (Asc(DriveLetter) - 65) ' Change letter to Number: A=0
  DriveType = GetDriveType(DriveLetter)
  If DriveType = 2 Then 'Floppies, etc
    RetVal = SHFormatDrive(Me.hwnd, DriveNumber, 0&, 0&)
  Else
    RetFromMsg = MsgBox("This drive is NOT a removeable" & vbCrLf & _
      "drive! Format this drive?", 276, "SHFormatDrive Example")
    Select Case RetFromMsg
      Case 6  'Yes
        ' UnComment to do it...
        'RetVal = SHFormatDrive(Me.hwnd, DriveNumber, 0&, 0&)
      Case 7  'No
        ' Do nothing
    End Select
  End If
End Sub
Private Sub cmdDiskCopy_Click()
' DiskCopyRunDll takes two parameters- From and To
  Dim DriveLetter$, DriveNumber&, DriveType&
  Dim RetVal&, RetFromMsg&
  DriveLetter = UCase(Drive1.Drive)
  DriveNumber = (Asc(DriveLetter) - 65)
  DriveType = GetDriveType(DriveLetter)
  If DriveType = 2 Then 'Floppies, etc
    RetVal = Shell("rundll32.exe diskcopy.dll,DiskCopyRunDll " _
      & DriveNumber & "," & DriveNumber, 1) 'Notice space after
  Else  ' Just in case             'DiskCopyRunDll
    RetFromMsg = MsgBox("Only floppies can" & vbCrLf & _
      "be diskcopied!", 64, "DiskCopy Example")
  End If
End Sub
Add 1 ListDrive name Drive1
Private Sub Drive1_Change()
  Dim DriveLetter$, DriveNumber&, DriveType&
  DriveLetter = UCase(Drive1.Drive)
  DriveNumber = (Asc(DriveLetter) - 65)
  DriveType = GetDriveType(DriveLetter)
  If DriveType 2 Then 'Floppies, etc
    cmdDiskCopy.Enabled = False
  Else
    cmdDiskCopy.Enabled = True
  End If
End Sub
```


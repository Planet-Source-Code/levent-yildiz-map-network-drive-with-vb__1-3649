<div align="center">

## MAP NETWORK DRIVE WITH VB


</div>

### Description

Well the simple way to map network drives with vb.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[LEVENT YILDIZ](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/levent-yildiz.md)
**Level**          |Unknown
**User Rating**    |4.9 (54 globes from 11 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/levent-yildiz-map-network-drive-with-vb__1-3649/archive/master.zip)

### API Declarations

```
Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpszName As String, ByVal bForce As Long) As Long
```


### Source Code

```
Private Sub cmdConnect_Click()
  Dim x As Long
  x = WNetAddConnection("\\CPU1\C\WINDOWS\DESKTOP", "", "R:")
  If x <> 0 Then
    MsgBox "connect failed"
  End If
End Sub
Private Sub cmdDisconnect_Click()
  Dim x As Long
  x = WNetCancelConnection("R:", 0)
  If x <> 0 Then
    MsgBox "Disconnect failed"
  End If
End Sub
```


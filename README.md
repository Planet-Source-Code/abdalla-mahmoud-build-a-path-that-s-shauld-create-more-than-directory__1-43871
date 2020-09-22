<div align="center">

## Build a Path That's Shauld Create More Than Directory


</div>

### Description

Hi All .. This Code Will Create a Path That's Shauld Make More Than One Directory .. For Example When You Call Function :

Call BuildPath ("C:\A1\A2\A3\A4")

It Will Create "C:\A1" and "C:\A1\A2" .etc

It Will Return False If The Drive Doesnt Exist Or For Any Other Error ..

Please Send Me Your Comments ..
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Abdalla Mahmoud](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/abdalla-mahmoud.md)
**Level**          |Beginner
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/abdalla-mahmoud-build-a-path-that-s-shauld-create-more-than-directory__1-43871/archive/master.zip)





### Source Code

```
Function BuildPath(ByVal Path As String) As Boolean
On Error Resume Next
Dim Fnd As Long
Dim Tmp As String
Dim FileSystemObj As Object
Set FileSystemObj = CreateObject("Scripting.FileSystemObject")
If FileSystemObj.DriveExists(FileSystemObj.GetDriveName(Path)) = False Then Exit Function
Path = Path & IIf(Right(Path, 1) = "\", vbNullString, "\")
Fnd = InStr(Path, "\")
Do While Fnd
Tmp = Tmp & Left(Path, Fnd)
Path = Mid(Path, Fnd + 1)
MkDir Tmp
If FileSystemObj.DriveExists(Tmp) = False And FileSystemObj.FolderExists(Tmp) = False Then Exit Function
Fnd = InStr(Path, "\")
Loop
BuildPath = True
End Function
Private Sub Command1_Click()
Call BuildPath("C:\A1\A2\A3\A4\A5\A6\A7")
End Sub
```


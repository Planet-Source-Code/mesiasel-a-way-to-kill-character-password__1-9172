<div align="center">

## A way to Kill Character password


</div>

### Description

Quita los caracteres "*" PASSWORD, funciona en varias aplicaciones w95/nt , algunas en las cuales el control cambia el "Class Name"

Hay programas que hacen practicamente lo mismo, este codigo sin embargo permite seguir escribiendo sin los caracteres PASSWORD almenos hasta que cierren la aplicaciòn
 
### More Info
 
El codigo se ejecuta al hacer un click en un boton, a travez de la posicion del mouse verifica el "Class Name" de las ventanas, cuando encuentra una con el tipo que buscamos, altera el caracter de PASSWORD a "", y luego sale del "loop"

El còdigo se puede reducir a la mitad, inclusive que siga verificando aunque encuentre la ventana deseada.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[mesiasel](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mesiasel.md)
**Level**          |Advanced
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mesiasel-a-way-to-kill-character-password__1-9172/archive/master.zip)

### API Declarations

Todas las declaraciones se encuentran en el còdigo


### Source Code

```
Private Type POINTAPI
 X As Long
 Y As Long
End Type
Private Declare Function GetClassNames Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal LpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetFocusAp Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Exper As Boolean
Private Sub Command1_Click()
Dim Point As POINTAPI, Cname As String, Resxxx As Long, LSta As Long
Dim Counter As Long, xxx As Long, Par As Long
Const Clase_Name As String = "ThunderTextBox"
Const Clase_Name2 As String = "Edit"
Exper = False
Do Until Exper = True
 Resxxx = GetCursorPos(Point)
 Resxxx = WindowFromPoint(Point.X, Point.Y)
 If Resxxx <> 0 Then
  Cname = String$(255, 0)
  xxx = GetClassNames(Resxxx, Cname, 254)
  If InStr(1, Cname, Clase_Name2, vbTextCompare) <> 0 Then
   Par = GetParent(Resxxx)
   xxx = SendMessage(Resxxx, &HCC, 0, 0)
   xxx = SetForegroundWindow(Par)
   xxx = UpdateWindow(Par)
   xxx = UpdateWindow(Resxxx)
   xxx = UpdateWindow(Resxxx)
   xxx = SetFocusAp(Resxxx)
   SetFocusAp xxx
   SetFocusAp Resxxx
  Exper = True
  End If
 End If
 DoEvents
Loop
End Sub
```


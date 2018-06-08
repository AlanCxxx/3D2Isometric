Option Explicit
Sub Main
  Dim StrId As Double
  Dim PStrId As Double
  Dim X As Double
  Dim Y As Double
  Dim Z As Double
  Dim XT As Double
  Dim YT As Double
  Dim ZT As Double
  Dim X0 As Double
  Dim Y0 As Double
  Dim Z0 As Double
  Dim X1 As Double
  Dim Y1 As Double
  Dim Z1 As Double
  Dim XI As Double
  Dim YI As Double
  Dim PX0 As Double
  Dim PY0 As Double
  Dim PZ0 As Double
  Dim PX1 As Double
  Dim PY1 As Double
  Dim PZ1 As Double
  Dim PXI As Double
  Dim PYI As Double
  Dim XMIN As Double
  Dim YMIN As Double
  Dim ZMIN As Double
  Dim XMAX As Double
  Dim YMAX As Double
  Dim ZMAX As Double
  Dim Clr As Long
  Dim AxisA(3) As Double
  Dim AxisB(3) As Double
  Dim AxisC(3) As Double
  Dim dir1() As String
  Dim FileList$(1000)
  Dim FileName$
  Dim FileDir$
  Dim FileExt$
  Dim FileLine$
  Dim CurrentLayer$
  Dim Temp As String
  Dim Rads As Double
  Dim Count As Long
  Dim i As Long
  Dim IsoOption As Long
  Dim Start As Long
  Dim Finish As Long

  ' Get drawing directory
  dcGetActiveWindow FileDir$
  i=Len(FileDir$)
  While (i>1) And (Mid(FileDir$,i,1)<>"\")
    i=i-1
  Wend
  If (i>1) Then 
    FileDir$=Mid(FileDir$,1,i)
  Else
    ' Unsaved file!
    ' Get a directory for user
    FileDir$="c:\..."
    FileDir$ = InputBox("Need a working directory path?","Drawing Not Saved!",FileDir$,100,100) 
    If ((FileDir$="c:\...") Or (FileDir$=""))  Then End
    If (Right(FileDir$,1)<>"\") Then FileDir$=FileDir$+"\"
  End If

  ' 0=Isometic View 1=Top View 2=Front View  3=Side View
  AxisA(0) = 54.735606: AxisB(0) = 0: AxisC(0) = 45      ' Isometric View
  AxisA(1) = 0:         AxisB(1) = 0: AxisC(1) = 0       ' Top View
  AxisA(2) = 90:        AxisB(2) = 0: AxisC(2) = 0       ' Front View
  AxisA(3) = 90:        AxisB(3) = 0: AxisC(3) =-90      ' Side View
  Rads = .01745329252

  ' Get File List
  FileExt$ = "CSV"
  temp = Dir(FileDir$+"*."+FileExt$ )
  count = 0
  While temp<>""
    count = count + 1
    temp = Dir
  Wend
  ReDim dir1(count)
  FileList$(0) = Dir(FileDir$+"*."+FileExt$ )
  For i = 1 To count
    FileList$(i) = Dir
  Next i

  ' Dialog definition
  Begin Dialog GetFileDIALOG 10,10,240,200, "Get CSV File"
    Text 50, 4, 40, 8, "Select CSV File"
    ComboBox 10,13,110,200,FileList$(),.Cmbo
    PushButton   130, 15, 37, 12, "Start", .StartFile
    CancelButton 130, 30, 37, 12 
    CheckBox 130,55,100,8,"Clear Drawing",.CheckBox2
    CheckBox 130,65,100,8,"Create Isometric Layers",.CheckBox3
    GroupBox 130,80,100,50, "View Setting", .Group1
      OptionGroup.OptionGroup1
      OptionButton 135, 90,  70, 8, "Isometric Only", .OptionButton1
      OptionButton 135, 100, 70 ,8, "Top Only",       .OptionButton2
      OptionButton 135, 110, 70, 8, "Full Isometric", .OptionButton3
  End Dialog

  ' Get dialog
  Dim Dlg1 As GetFileDIALOG
  Dim Button As Long 
  Do While (FileName$="")
    Dlg1.OptionGroup1=1
    Button = Dialog (Dlg1)
    If button = 0 Then End            ' Cancel - Exit
    FileName$ = Dlg1.Cmbo
  Loop

  ' Clear drawing
  If GetFileDIALOG.CheckBox2 Then
     dcCreatePoint 0,0
     dcSelectAll
     dcEraseSelObjs 
  End If

  ' Set layers
  dcGetCurrentLayer CurrentLayer$
  If GetFileDIALOG.CheckBox3 Then
    If Not(dcDoesLayerExist("ISOMETRIC")) Then
      dcAddLayer "ISOMETRIC"
    End If
    If Not(dcDoesLayerExist("TOP")) Then
      dcAddLayer "TOP"
    End If
    If Not(dcDoesLayerExist("FRONT")) Then
      dcAddLayer "FRONT"
    End If
    If Not(dcDoesLayerExist("SIDE")) Then
      dcAddLayer "SIDE"
    End If
  End If

  ' Set Option
   Start=Dlg1.OptionGroup1
   Finish=Dlg1.OptionGroup1
   If (Dlg1.OptionGroup1=2) Then
     Start=0
     Finish=3
   End If

  ' Get limits
  FileName$=FileDir$+FileName$
  Open FileName$ For Input As #1 
  Line Input #1, FileLine$
  XMIN=1E300
  YMIN=1E300
  ZMIN=1E300
  XMAX=-1E300
  YMAX=-1E300
  ZMAX=-1E300
  Do While Not  EOF(1) 
    Input #1, StrId,X,Y,Z
    If (StrId > 0) Then
      If (XMIN>X) Then XMIN=X
      If (YMIN>Y) Then YMIN=Y
      If (ZMIN>Z) Then ZMIN=Z
      If (XMAX<X) Then XMAX=X
      If (YMAX<Y) Then YMAX=Y
      If (ZMAX<Z) Then ZMAX=Z
    End If
  Loop
  ' Draw the data
  Seek #1,1
  Line Input #1, FileLine$
  X1=0
  Y1=0
  Z1=0
  StrId = 0
  Do While Not  EOF(1) 
    X0=X1
    Y0=Y1
    Z0=Z1
    PStrId=StrId
    Input #1, StrId,X1,Y1,Z1
    Clr = (StrId-1) Mod 16
    X1=X1-XMIN
    Y1=Y1-YMIN
    Z1=Z1-ZMIN
    If (StrId = PStrId) And (StrId > 0) Then
      For IsoOption=Start to Finish

        XT=Cos(AxisC(IsoOption)*Rads)*X0-Sin(AxisC(IsoOption)*Rads)*Y0
        YT=Sin(AxisC(IsoOption)*Rads)*X0+Cos(AxisC(IsoOption)*Rads)*Y0
        ZT=Cos(AxisA(IsoOption)*Rads)*Z0-Sin(AxisA(IsoOption)*Rads)*YT
        PY0=Sin(AxisA(IsoOption)*Rads)*Z0+Cos(AxisA(IsoOption)*Rads)*YT
        PZ0=Cos(AxisB(IsoOption)*Rads)*ZT-Sin(AxisB(IsoOption)*Rads)*XT
        PX0=Sin(AxisB(IsoOption)*Rads)*ZT+Cos(AxisB(IsoOption)*Rads)*XT

        XT=Cos(AxisC(IsoOption)*Rads)*X1-Sin(AxisC(IsoOption)*Rads)*Y1
        YT=Sin(AxisC(IsoOption)*Rads)*X1+Cos(AxisC(IsoOption)*Rads)*Y1
        ZT=Cos(AxisA(IsoOption)*Rads)*Z1-Sin(AxisA(IsoOption)*Rads)*YT
        PY1=Sin(AxisA(IsoOption)*Rads)*Z1+Cos(AxisA(IsoOption)*Rads)*YT
        PZ1=Cos(AxisB(IsoOption)*Rads)*ZT-Sin(AxisB(IsoOption)*Rads)*XT
        PX1=Sin(AxisB(IsoOption)*Rads)*ZT+Cos(AxisB(IsoOption)*Rads)*XT

        If (IsoOption=0) Then
          ' Isometric
          If (dcDoesLayerExist("ISOMETRIC")) Then dcSetCurrentLayer "ISOMETRIC"
          PX0=PX0/0.81649658+2*XMAX-XMIN
          PY0=PY0/0.816496581+YMIN
          PX1=PX1/0.81649658+2*XMAX-XMIN
          PY1=PY1/0.816496581+YMIN
        ElseIf (IsoOption=1) Then
          ' Top (no offsets)
          If (dcDoesLayerExist("TOP")) Then dcSetCurrentLayer "TOP"
          PX0=PX0+XMIN
          PY0=PY0+YMIN
          PX1=PX1+XMIN
          PY1=PY1+YMIN
        ElseIf (IsoOption=2) Then
          ' Front
          If (dcDoesLayerExist("FRONT")) Then dcSetCurrentLayer "FRONT"
          PX0=PX0+XMIN
          PY0=PY0+YMIN-1.5*(ZMAX-ZMIN)
          PX1=PX1+XMIN
          PY1=PY1+YMIN-1.5*(ZMAX-ZMIN)
        ElseIf (IsoOption=3) Then
          ' Side
          If (dcDoesLayerExist("SIDE")) Then dcSetCurrentLayer "SIDE"
          PX0=PX0+XMIN+1.5*(XMAX-XMIN)
          PY0=PY0+YMIN-1.5*(ZMAX-ZMIN)
          PX1=PX1+XMIN+1.5*(XMAX-XMIN)
          PY1=PY1+YMIN-1.5*(ZMAX-ZMIN)
        End If

         ' DeltaCad does not handle zero length lines very so use a point instead
        If ((Abs(PX1-PX0)<0.001) And (Abs(PY1-PY0)<0.001)) Then
          dcSetPointParms Clr, True
          dcCreatePoint PX0, PY0
        Else
          dcSetLineParms Clr, dcSOLID, dcNORMAL
          dcCreateLine PX0, PY0,PX1, PY1
        End If

      Next IsoOption

    End If

  Loop
  Close #1 ' Close file

  dcSetCurrentLayer CurrentLayer$ 
  dcSetPointParms dcBLACK, True
  dcSetLineParms dcBLACK, dcSOLID, dcNORMAL
  dcUpdateDisplay True
  dcViewAll
  MsgBox "Z offset is "+Format(ZMIN,"###0.000")

End Sub

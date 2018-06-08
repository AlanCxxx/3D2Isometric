Option Explicit
Sub Main
  ' LINE DATA
  Dim ot As Long
  Dim xa As Double 
  Dim ya As Double 
  Dim xb As Double 
  Dim yb As Double
  Dim color As Long 
  Dim lt As Long 
  Dim lw As Long
  Dim layer As String 
  Dim arrow As Long 
  Dim pflag As Boolean
  ' ISOMETRIC DATA
  Dim ns As Long
  Dim ysa() As Double
  Dim ysb() As Double
  Dim zsa() As Double
  Dim zsb() As Double
  Dim clrs() As Long
  Dim dones() As Boolean
  Dim nf As Long
  Dim xfa() As Double
  Dim xfb() As Double
  Dim zfa() As Double
  Dim zfb() As Double
  Dim clrf() As Long
  Dim donef() As Boolean
  Dim nt As Long
  Dim xta() As Double
  Dim xtb() As Double
  Dim yta() As Double
  Dim ytb() As Double
  Dim clrt() As Long
  Dim donet() As Boolean
  Dim n As Long
  Dim xla() As Double
  Dim xlb() As Double
  Dim yla() As Double
  Dim ylb() As Double
  Dim zla() As Double
  Dim zlb() As Double
  Dim clrl() As Long
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim jStart As Long 
  Dim kStart As Long
  Dim xmin As Double
  Dim ymin As Double
  Dim zmin As Double
  Dim ZOffset As Double
  Dim t As Double
  Dim Flag As Boolean
  ' File
  Dim FileDir$
  Dim FileName$
 
  ' Check if isometric layers exist
  If (dcDoesLayerExist("TOP") And dcDoesLayerExist("FACE") And dcDoesLayerExist("SIDE")) Then
    MsgBox "Isometric Layers missing?"
    End
  End If

  ' Get drawing directory
  dcGetActiveWindow FileDir$
  i=Len(FileDir$)
  While (i>1) And (Mid(FileDir$,i,1)<>"\")
    i=i-1
  Wend
  If (i>1) Then 
    FileName$=Mid(FileDir$,i+1,Len(FileDir$)-i-3)
    FileDir$=Mid(FileDir$,1,i)
  Else
    ' Unsaved file! Use DeltaCad program directory
    FileName$="test"
    dcGetDeltaCadProgramDirectory FileDir$
  End If

  ' Get file name
  FileName$ = InputBox("Export Isometric to CSV","CSV File name",FileName$,100,100) 
  If (FileName$="") Then End
  FileName$=FileDir$+FileName$+".csv"
 
  ' Get Offset Z
  ZOffset=InputBox("Input Z Offset","Z Offset",0.0,100,100) 

  ' Get array sizes
  nt=0
  nf=0
  ns=0
  ot=dcGetFirstObject 
  While (ot<>dcNone) 
    If (ot=dcLine) Then 
      dcGetLineData xa, ya, xb, yb, color, lt, lw, layer, arrow
      If (layer="SIDE") Then ns=ns+1
      If (layer="FRONT") Then nf=nf+1
      If (layer="TOP") Then nt=nt+1
    ElseIf (ot=dcPoint) Then
      dcGetPointData xa, ya, color, layer, pflag
      If (pflag=True) Then
        If (layer="SIDE") Then ns=ns+1
        If (layer="FRONT") Then nf=nf+1
        If (layer="TOP") Then nt=nt+1
      End If
    End If
    ot=dcGetNextObject
  Wend
  If (nt<>nf) Or (nt<>ns) Or (nf<>ns) Then
    MsgBox "Warning: Isometric layers do not match: Top "+Format(nt)+" Front "+Format(nf)+" Side "+Format(ns)   
    ' End
  End If
  If (nt>1000) Then
    MsgBox Format(nt)+" segments found - This will take some time!"   
  End If
  ReDim ysa(ns)
  ReDim ysb(ns) 
  ReDim zsa(ns) 
  ReDim zsb(ns) 
  ReDim clrs(ns) 
  ReDim dones(ns) 
  ReDim xfa(nf) 
  ReDim xfb(nf) 
  ReDim zfa(nf) 
  ReDim zfb(nf) 
  ReDim clrf(nf) 
  ReDim donef(nf) 
  ReDim xta(nt) 
  ReDim xtb(nt) 
  ReDim yta(nt) 
  ReDim ytb(nt) 
  ReDim clrt(nt)
  ReDim donet(nt)
  ReDim xla(nt) 
  ReDim xlb(nt) 
  ReDim yla(nt) 
  ReDim ylb(nt) 
  ReDim zla(nt) 
  ReDim zlb(nt) 
  ReDim clrl(nt)

  ' Load the data
  nt=0
  nf=0
  ns=0
  ot=dcGetFirstObject 
  While (ot<>dcNone) 
    If (ot=dcLine) Then 
      dcGetLineData xa, ya, xb, yb, color, lt, lw, layer, arrow
      If (layer="SIDE") Then
        ns=ns+1
        ysa(ns)=xa
        zsa(ns)=ya
        ysb(ns)=xb
        zsb(ns)=yb
        clrs(ns)=color+1
      ElseIf (layer="FRONT") Then
        nf=nf+1
        xfa(nf)=xa
        zfa(nf)=ya
        xfb(nf)=xb
        zfb(nf)=yb
        clrf(nf)=color+1
      ElseIf (layer="TOP") Then
        nt=nt+1
        xta(nt)=xa
        yta(nt)=ya
        xtb(nt)=xb
        ytb(nt)=yb
        clrt(nt)=color+1
      End If 
    ElseIf (ot=dcPoint) Then 
      dcGetPointData xa, ya, color, layer, pflag
      If (pflag=True) Then
        If (layer="SIDE") Then
          ns=ns+1
          ysa(ns)=xa
          zsa(ns)=ya
          ysb(ns)=xa
          zsb(ns)=ya
          clrs(ns)=color+1
        ElseIf (layer="FRONT") Then
          nf=nf+1
          xfa(nf)=xa
          zfa(nf)=ya
          xfb(nf)=xa
          zfb(nf)=ya
          clrf(nf)=color+1
        ElseIf (layer="TOP") Then
          nt=nt+1
          xta(nt)=xa
          yta(nt)=ya
          xtb(nt)=xa
          ytb(nt)=ya
          clrt(nt)=color+1
        End If 
      End If 
    End If 
    ot=dcGetNextObject
  Wend 

  ' Get/set range
  ymin=1e300
  zmin=1e300
  For i=1 to ns
    If ysa(i)<ymin Then ymin=ysa(i)
    If ysb(i)<ymin Then ymin=ysb(i)
    If zsa(i)<zmin Then zmin=zsa(i)
    If zsb(i)<zmin Then zmin=zsb(i)
  Next i
  For i=1 to ns
    ysa(i)=Int(1000*(ysa(i)-ymin)+0.5)/1000
    ysb(i)=Int(1000*(ysb(i)-ymin)+0.5)/1000
    zsa(i)=Int(1000*(zsa(i)-zmin)+0.5)/1000
    zsb(i)=Int(1000*(zsb(i)-zmin)+0.5)/1000
    If ((ysa(i)=ysb(i)) And (zsa(i)>zsb(i))) Then
      t=zsa(i):zsa(i)=zsb(i):zsb(i)=t
    End If
  Next i

  ' front
  xmin=1e300
  zmin=1e300
  For i=1 to nf
    If xfa(i)<xmin Then xmin=xfa(i)
    If xfb(i)<xmin Then xmin=xfb(i)
    If zfa(i)<zmin Then zmin=zfa(i)
    If zfb(i)<zmin Then zmin=zfb(i)
  Next i
  For i=1 to nf
    xfa(i)=Int(1000*(xfa(i)-xmin)+0.5)/1000
    xfb(i)=Int(1000*(xfb(i)-xmin)+0.5)/1000
    zfa(i)=Int(1000*(zfa(i)-zmin)+0.5)/1000
    zfb(i)=Int(1000*(zfb(i)-zmin)+0.5)/1000
    If ((xfa(i)=xfb(i)) And (zfa(i)>zfb(i))) Then
      t=zfa(i):zfa(i)=zfb(i):zfb(i)=t
    End If
  Next i

  ' top
  xmin=1e300
  ymin=1e300
  For i=1 to nt
    If xta(i)<xmin Then xmin=xta(i)
    If xtb(i)<xmin Then xmin=xtb(i)
    If yta(i)<ymin Then ymin=yta(i)
    If ytb(i)<ymin Then ymin=ytb(i)
  Next i
  For i=1 to nt
    xta(i)=Int(1000*(xta(i)-xmin)+0.5)/1000
    xtb(i)=Int(1000*(xtb(i)-xmin)+0.5)/1000
    yta(i)=Int(1000*(yta(i)-ymin)+0.5)/1000
    ytb(i)=Int(1000*(ytb(i)-ymin)+0.5)/1000
    If ((xta(i)=xtb(i)) And (yta(i)>ytb(i))) Then
      t=yta(i):yta(i)=ytb(i):ytb(i)=t
    End If
  Next i
  zmin=ZOffset

  ' Export check file
  ' Open FileDir$+"temp.csv" For output As #1
  ' For i=1 TO nt
  '   Print #1,"1, ",clrt(i),", ",xta(i),", ",yta(i),", ",    -1,", ",xtb(i),", ",ytb(i),", ",    -1
  '   Print #1,"2, ",clrf(i),", ",xfa(i),", ",    -1,", ",zfa(i),", ",xfb(i),", ",    -1,", ",zfb(i)
  '   Print #1,"3, ",clrs(i),", ",    -1,", ",ysa(i),", ",zsa(i),", ",    -1,", ",ysb(i),", ",zsb(i)
  ' Next i
  ' Close #1

 ' Sort Top
  k=nt
  While (k>1)
    k=Int(k/3+1)
    For i=k+1 to nt
      xta(0)=xta(i)
      yta(0)=yta(i)
      xtb(0)=xtb(i)
      ytb(0)=ytb(i)
      clrt(0)=clrt(i)
      j=i
      Flag=True
      While (Flag)
        Flag=False
        If (j>k) Then
          If (clrt(j-k)>clrt(0)) Then Flag=True
          If ((clrt(j-k)=clrt(0)) And (xta(j-k)>xta(0))) Then Flag=True
          If ((clrt(j-k)=clrt(0)) And (xta(j-k)=xta(0)) And (yta(j-k)>yta(0))) Then Flag=True
          If (Flag) Then
            xta(j)=xta(j-k)
            yta(j)=yta(j-k)
            xtb(j)=xtb(j-k)
            ytb(j)=ytb(j-k)
            clrt(j)=clrt(j-k)
            j=j-k
          End If
        End If
      Wend
      xta(j)=xta(0)
      yta(j)=yta(0)
      xtb(j)=xtb(0)
      ytb(j)=ytb(0)
      clrt(j)=clrt(0)
    Next i
  Wend
  
 ' Sort Front
  k=nf
  While (k>1)
    k=Int(k/3+1)
    For i=k+1 to nf
      xfa(0)=xfa(i)
      zfa(0)=zfa(i)
      xfb(0)=xfb(i)
      zfb(0)=zfb(i)
      clrf(0)=clrf(i)
      j=i
      Flag=True
      While (Flag)
        Flag=False
        If (j>k) Then
          If (clrf(j-k)>clrf(0)) Then Flag=True
          If ((clrf(j-k)=clrf(0)) And (xfa(j-k)>xfa(0))) Then Flag=True
          If ((clrf(j-k)=clrf(0)) And (xfa(j-k)=xfa(0)) And (zfa(j-k)>zfa(0))) Then Flag=True
          If (Flag) Then
            xfa(j)=xfa(j-k)
            zfa(j)=zfa(j-k)
            xfb(j)=xfb(j-k)
            zfb(j)=zfb(j-k)
            clrf(j)=clrf(j-k)
            j=j-k
          End If
        End If
      Wend
      xfa(j)=xfa(0)
      zfa(j)=zfa(0)
      xfb(j)=xfb(0)
      zfb(j)=zfb(0)
      clrf(j)=clrf(0)
    Next i
  Wend

 ' Sort Side
  k=ns
  While (k>1)
    k=Int(k/3+1)
    For i=k+1 to ns
      ysa(0)=ysa(i)
      zsa(0)=zsa(i)
      ysb(0)=ysb(i)
      zsb(0)=zsb(i)
      clrs(0)=clrs(i)
      j=i
      Flag=True
      While (Flag)
        Flag=False
        If (j>k) Then
          If (clrs(j-k)>clrs(0)) Then Flag=True
          If ((clrs(j-k)=clrs(0)) And (ysa(j-k)>ysa(0))) Then Flag=True
          If ((clrs(j-k)=clrs(0)) And (ysa(j-k)=ysa(0)) And (zsa(j-k)>zsa(0))) Then Flag=True
          If (Flag) Then
            ysa(j)=ysa(j-k)
            zsa(j)=zsa(j-k)
            ysb(j)=ysb(j-k)
            zsb(j)=zsb(j-k)
            clrs(j)=clrs(j-k)
            j=j-k
          End If
        End If
      Wend
      ysa(j)=ysa(0)
      zsa(j)=zsa(0)
      ysb(j)=ysb(0)
      zsb(j)=zsb(0)
      clrs(j)=clrs(0)
    Next i
  Wend

  ' Match and Dispatch
  For i=1 to nt
    donet(i)=False
    donef(i)=False
    dones(i)=False
  Next i
  n=0
  jStart=1
  kStart=1
  For i=1 to nt
    Do 
      If (jStart>=ns) Then Exit Do
      If (donef(jStart)=False) Then Exit Do
      jStart=jStart+1
    Loop
    For j=jStart to ns
      If (clrf(j)>clrt(i)) Then GoTo nextpnt
      If ((clrf(j)=clrt(i)) And (donef(j)=False)) Then
        If ((Abs(xta(i)-xfa(j))<=0.001) And (Abs(xtb(i)-xfb(j))<=0.001)) Then
          While (clrs(kStart)<clrf(j))
            kStart=kStart+1
          Wend 
          For k=kStart to ns
            If (clrs(k)>clrt(i)) Then Exit For
            If (clrs(k)=clrt(i)) Then
              If (yta(i)<=ytb(i)) Then
                If ((Abs(yta(i)-ysa(k))<=0.001) And (Abs(ytb(i)-ysb(k))<=0.001)) Then
                  If ((Abs(zfa(j)-zsa(k))<=0.001) And (Abs(zfb(j)-zsb(k))<=0.001)) Then
                    If (dones(k)=False) Then
                      n=n+1
                      xla(n)=xta(i)+xmin
                      yla(n)=yta(i)+ymin
                      zla(n)=zfa(j)+zmin
                      xlb(n)=xtb(i)+xmin
                      ylb(n)=ytb(i)+ymin
                      zlb(n)=zfb(j)+zmin
                      clrl(n)=clrt(i)
                      donet(i)=True
                      donef(j)=True
                      dones(k)=True
                      GoTo nextpnt
                    End If
                  End If
                End If
              End If
              If (yta(i)>=ytb(i)) Then
                If ((Abs(yta(i)-ysb(k))<=0.001) And (Abs(ytb(i)-ysa(k))<=0.001)) Then
                  If ((Abs(zfa(j)-zsb(k))<=0.001) And (Abs(zfb(j)-zsa(k))<=0.001)) Then
                    If (dones(k)=False) Then
                      n=n+1
                      xla(n)=xta(i)+xmin
                      yla(n)=yta(i)+ymin
                      zla(n)=zfa(j)+zmin
                      xlb(n)=xtb(i)+xmin
                      ylb(n)=ytb(i)+ymin
                      zlb(n)=zfb(j)+zmin
                      clrl(n)=clrt(i)
                      donet(i)=True
                      donef(j)=True
                      dones(k)=True
                      GoTo nextpnt
                    End If
                  End If
                End If
              End If
            End If
          Next k
        End If
      End If
    Next j
    nextpnt:
  Next i  


  ' Export the file
  Open FileName$ For output As #1
  Print #1, "StrId, ","X, ","Y, ","Z, ",FileName$;": ";n;" conversions out of ";nt;" segments"
  For i=1 TO n
    Print #1,Format(clrl(i),"0"),", ",Format(xla(i),"0.000"),", ",Format(yla(i),"0.000"),", ",Format(zla(i),"0.000")
    Print #1,Format(clrl(i),"0"),", ",Format(xlb(i),"0.000"),", ",Format(ylb(i),"0.000"),", ",Format(zlb(i),"0.000")
    Print #1,Format(0,"0"),", ",Format(0.0,"0.000"),", ",Format(0.0,"0.000"),", ",Format(0.0,"0.000")
  Next i
  Close #1

  MsgBox FileName$+": "+Format(n)+" conversions out of "+Format(nt)+" segments"
 
End Sub


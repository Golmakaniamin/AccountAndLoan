Attribute VB_Name = "PrintAll"
Dim fso As New FileSystemObject
Dim adoConnection1 As ADODB.Connection
Dim cmd1 As ADODB.Command
Dim adoRecordset1 As ADODB.Recordset
Dim a As String, e As Integer

'  cmd1.CommandText = "DELETE FROM allp"
'  cmd1.CommandType = adCmdText
'  cmd1.Properties.Refresh
'  Set adoRecordset1 = cmd1.Execute
  
'  cmd1.CommandText = "SELECT * FROM allp ORDER BY rad ASC"
'  cmd1.CommandType = adCmdText
'  cmd1.Properties.Refresh
'  Set adoRecordset1 = cmd1.Execute

Public Sub printforall()
  fso.CopyFile "Z:\PraticGroup\SavingBankSoftware(KarimAhlBeyt)\DATA BASE INFORMATION\Data\info2.mdb", "Z:\PraticGroup\SavingBankSoftware(KarimAhlBeyt)\DATA BASE INFORMATION\info2.mdb", True
  Form2.Adodc1.Refresh
  Form2.Adodc2.Refresh
  Form2.Adodc3.Refresh
  
  Form3.Adodc1.Recordset.Sort = "id"
  Form3.Adodc1.Recordset.MoveFirst
  Do
    'ãæÌæÏí ÚÇÏí
    Form2.Adodc1.Recordset.AddNew
    Form2.Adodc1.Recordset.Fields!id = Form3.Adodc1.Recordset.Fields!id
    Form2.Adodc1.Recordset.Fields!Money = Form3.Adodc1.Recordset.Fields!Money
    Form2.Adodc1.Recordset.Fields!edate = Form3.Adodc1.Recordset.Fields!edate
    q = Val(Amin.dateaminEktelafmoon(Form3.Adodc1.Recordset.Fields!edate, Form2.Label5.Caption))
    w = 0
    w = Val(Form2.Adodc1.Recordset.Fields!Money) - (q * 20000)
    If w >= 0 Then
      Form2.Adodc1.Recordset.Fields!kasr = "0"
      Form2.Adodc1.Recordset.Fields!afza = w
    Else
      Form2.Adodc1.Recordset.Fields!kasr = -1 * Val(w)
      Form2.Adodc1.Recordset.Fields!afza = "0"
    End If
    Form2.Adodc1.Recordset.Update

    'æÇã åÇí ÚÇÏí
    Form2.Adodc10.CommandType = adCmdUnknown
    Form2.Adodc10.RecordSource = "SELECT * From pvamadi WHERE (id1='" + Form3.Adodc1.Recordset.Fields!id + "')"
    Form2.Adodc10.Refresh
    If Form2.Adodc10.Recordset.RecordCount > 0 Then
      Form2.Adodc10.Recordset.MoveFirst
      Do
        Form2.Adodc2.Refresh
        Form2.Adodc2.Recordset.AddNew
        Form2.Adodc2.Recordset.Fields!id = Form3.Adodc1.Recordset.Fields!id
        Form2.Adodc2.Recordset.Fields!rad = "0"
        Form2.Adodc2.Recordset.Fields!idvam = Form2.Adodc10.Recordset.Fields!id
        Form2.Adodc2.Recordset.Fields!moneyvam = Form2.Adodc10.Recordset.Fields!moneyvam
        Form2.Adodc2.Recordset.Fields!allghest = Form2.Adodc10.Recordset.Fields!numberagsat
        Form2.Adodc2.Recordset.Fields!moneyghest1 = Form2.Adodc10.Recordset.Fields!moneyg1
        Form2.Adodc2.Recordset.Fields!moneyghest2 = Form2.Adodc10.Recordset.Fields!moneyg2
        Form2.Adodc2.Recordset.Fields!karmozd = Form2.Adodc10.Recordset.Fields!karmozd
        Form2.Adodc2.Recordset.Fields!numberallaghsat = Form2.Adodc10.Recordset.Fields!numberagsat

        Form2.Adodc11.CommandType = adCmdUnknown
        Form2.Adodc11.RecordSource = "SELECT * From GvamAdi WHERE (id='" + Form2.Adodc10.Recordset.Fields!id + "') ORDER BY rad ASC"
        Form2.Adodc11.Refresh
        q = 0
        w = 0

        If Form2.Adodc11.Recordset.RecordCount > 0 Then
          Form2.Adodc11.Recordset.MoveFirst
          Do
            q = Val(q) + Val(Form2.Adodc11.Recordset.Fields!Money)
            w = Val(w) + Val(Form2.Adodc11.Recordset.Fields!emteyaz)
            Form2.Adodc11.Recordset.MoveNext
          Loop Until Form2.Adodc11.Recordset.EOF = True
          Form2.Adodc2.Recordset.Fields!numberpardakhtaghsat = Form2.Adodc11.Recordset.RecordCount
          Form2.Adodc2.Recordset.Fields!numberpardakhtnaghsat = Form2.Adodc10.Recordset.Fields!numberagsat - Form2.Adodc11.Recordset.RecordCount
          Form2.Adodc2.Recordset.Fields!bestankari = Val(Form2.Adodc10.Recordset.Fields!moneyvam) - q
          Form2.Adodc2.Recordset.Fields!emteyaz = w
        Else
          Form2.Adodc2.Recordset.Fields!numberpardakhtaghsat = 0
          Form2.Adodc2.Recordset.Fields!numberpardakhtnaghsat = Form2.Adodc10.Recordset.Fields!numberagsat
          Form2.Adodc2.Recordset.Fields!bestankari = Form2.Adodc10.Recordset.Fields!moneyvam
          Form2.Adodc2.Recordset.Fields!emteyaz = 0
        End If

        'ÂÎÑíä ÞÓØ
        e = 0
        If Form2.Adodc10.Recordset.Fields!tasvie = "äÔÏå" Then
          If Form2.Adodc11.Recordset.RecordCount > 0 Then
            Form2.Adodc11.Recordset.MoveLast
            a = Form2.Adodc11.Recordset.Fields!saragsat
          Else
            a = Amin.dateaminEzafeMoon(Form2.Adodc10.Recordset.Fields!Date, 1)
          End If

          If (a <= Form2.Label5.Caption) Then
            e = Val(Amin.dateaminEktelafmoon(a, Form2.Label5.Caption))
          End If

          If e > Val(Form2.Adodc2.Recordset.Fields!numberpardakhtnaghsat) Then
            e = Val(Form2.Adodc2.Recordset.Fields!numberpardakhtnaghsat)
          End If
        End If
        Form2.Adodc2.Recordset.Fields!numbermoaghsat = Val(e)
        Form2.Adodc2.Recordset.Fields!moneymo = Form2.Adodc2.Recordset.Fields!moneyghest2 * Val(e)
        Form2.Adodc2.Recordset.Update
        Form2.Adodc10.Recordset.MoveNext
      Loop Until Form2.Adodc10.Recordset.EOF = True
    End If

'    æÇã åÇí ÇÖØÑÇÑí
    Form2.Adodc10.CommandType = adCmdUnknown
    Form2.Adodc10.RecordSource = "SELECT * From pvamaz WHERE (id1='" + Form3.Adodc1.Recordset.Fields!id + "')"
    Form2.Adodc10.Refresh
    If Form2.Adodc10.Recordset.RecordCount > 0 Then
      Form2.Adodc10.Recordset.MoveFirst
      Do
        Form2.Adodc2.Refresh
        Form2.Adodc2.Recordset.AddNew
        Form2.Adodc2.Recordset.Fields!id = Form3.Adodc1.Recordset.Fields!id
        Form2.Adodc2.Recordset.Fields!rad = "0"
        Form2.Adodc2.Recordset.Fields!idvam = Form2.Adodc10.Recordset.Fields!id
        Form2.Adodc2.Recordset.Fields!moneyvam = Form2.Adodc10.Recordset.Fields!moneyvam
        Form2.Adodc2.Recordset.Fields!allghest = Form2.Adodc10.Recordset.Fields!numberagsat
        Form2.Adodc2.Recordset.Fields!moneyghest1 = Form2.Adodc10.Recordset.Fields!moneyg1
        Form2.Adodc2.Recordset.Fields!moneyghest2 = Form2.Adodc10.Recordset.Fields!moneyg2
        Form2.Adodc2.Recordset.Fields!karmozd = Form2.Adodc10.Recordset.Fields!karmozd
        Form2.Adodc2.Recordset.Fields!numberallaghsat = Form2.Adodc10.Recordset.Fields!numberagsat

        Form2.Adodc11.CommandType = adCmdUnknown
        Form2.Adodc11.RecordSource = "SELECT * From GvamAz WHERE (id='" + Form2.Adodc10.Recordset.Fields!id + "') ORDER BY rad ASC"
        Form2.Adodc11.Refresh
        q = 0
        w = 0

        If Form2.Adodc11.Recordset.RecordCount > 0 Then
          Form2.Adodc11.Recordset.MoveFirst
          Do
            q = Val(q) + Val(Form2.Adodc11.Recordset.Fields!Money)
            w = Val(w) + Val(Form2.Adodc11.Recordset.Fields!emteyaz)
            Form2.Adodc11.Recordset.MoveNext
          Loop Until Form2.Adodc11.Recordset.EOF = True
          Form2.Adodc2.Recordset.Fields!numberpardakhtaghsat = Form2.Adodc11.Recordset.RecordCount
          Form2.Adodc2.Recordset.Fields!numberpardakhtnaghsat = Form2.Adodc10.Recordset.Fields!numberagsat - Form2.Adodc11.Recordset.RecordCount
          Form2.Adodc2.Recordset.Fields!bestankari = Val(Form2.Adodc10.Recordset.Fields!moneyvam) - q
          Form2.Adodc2.Recordset.Fields!emteyaz = w
        Else
          Form2.Adodc2.Recordset.Fields!numberpardakhtaghsat = 0
          Form2.Adodc2.Recordset.Fields!numberpardakhtnaghsat = Form2.Adodc10.Recordset.Fields!numberagsat
          Form2.Adodc2.Recordset.Fields!bestankari = Form2.Adodc10.Recordset.Fields!moneyvam
          Form2.Adodc2.Recordset.Fields!emteyaz = 0
        End If

'        ÂÎÑíä ÞÓØ
        e = 0
        If Form2.Adodc10.Recordset.Fields!tasvie = "äÔÏå" Then
          If Form2.Adodc11.Recordset.RecordCount > 0 Then
            Form2.Adodc11.Recordset.MoveLast
            a = Form2.Adodc11.Recordset.Fields!saragsat
          Else
            a = Amin.dateaminEzafeMoon(Form2.Adodc10.Recordset.Fields!Date, 1)
          End If

          If (a <= Form2.Label5.Caption) Then
            e = Val(Amin.dateaminEktelafmoon(a, Form2.Label5.Caption))
          End If

          If e > Val(Form2.Adodc2.Recordset.Fields!numberpardakhtnaghsat) Then
            e = Val(Form2.Adodc2.Recordset.Fields!numberpardakhtnaghsat)
          End If
        End If
        Form2.Adodc2.Recordset.Fields!numbermoaghsat = Val(e)
        Form2.Adodc2.Recordset.Fields!moneymo = Form2.Adodc2.Recordset.Fields!moneyghest2 * Val(e)
        Form2.Adodc2.Recordset.Update
        Form2.Adodc10.Recordset.MoveNext
      Loop Until Form2.Adodc10.Recordset.EOF = True
    End If
    
    z = 1
    'æÇã åÇí æíŽå
    Form2.Adodc12.CommandType = adCmdUnknown
    Form2.Adodc12.RecordSource = "SELECT * From Accountvig WHERE (idadi='" + Form3.Adodc1.Recordset.Fields!id + "')"
    Form2.Adodc12.Refresh
    If Form2.Adodc12.Recordset.RecordCount > 0 Then
      Form2.Adodc12.Recordset.MoveFirst
      Do
        Form2.Adodc10.CommandType = adCmdUnknown
        Form2.Adodc10.RecordSource = "SELECT * From pvamvig WHERE (id1='" + Form2.Adodc12.Recordset.Fields!id + "')"
        Form2.Adodc10.Refresh
        If Form2.Adodc10.Recordset.RecordCount > 0 Then
          Form2.Adodc10.Recordset.MoveFirst
          Do
            Form2.Adodc3.Refresh
            Form2.Adodc3.Recordset.AddNew
            Form2.Adodc3.Recordset.Fields!id = Form3.Adodc1.Recordset.Fields!id
            Form2.Adodc3.Recordset.Fields!rad = z
            Form2.Adodc3.Recordset.Fields!idaccount = Form2.Adodc12.Recordset.Fields!id
            Form2.Adodc3.Recordset.Fields!moneyaccount = Form2.Adodc12.Recordset.Fields!Money
            Form2.Adodc3.Recordset.Fields!idvam = Form2.Adodc10.Recordset.Fields!id
            Form2.Adodc3.Recordset.Fields!moneyvam = Form2.Adodc10.Recordset.Fields!moneyvam
            Form2.Adodc3.Recordset.Fields!allghest = Form2.Adodc10.Recordset.Fields!numberagsat
            Form2.Adodc3.Recordset.Fields!moneyghest1 = Form2.Adodc10.Recordset.Fields!moneyg1
            Form2.Adodc3.Recordset.Fields!moneyghest2 = Form2.Adodc10.Recordset.Fields!moneyg2
            Form2.Adodc3.Recordset.Fields!karmozd = Form2.Adodc10.Recordset.Fields!karmozd
            Form2.Adodc3.Recordset.Fields!numberallaghsat = Form2.Adodc10.Recordset.Fields!numberagsat

            Form2.Adodc11.CommandType = adCmdUnknown
            Form2.Adodc11.RecordSource = "SELECT * From GvamVig WHERE (id='" + Form2.Adodc10.Recordset.Fields!id + "') ORDER BY rad ASC"
            Form2.Adodc11.Refresh
            q = 0
            w = 0

            If Form2.Adodc11.Recordset.RecordCount > 0 Then
              Form2.Adodc11.Recordset.MoveFirst
              Do
                q = Val(q) + Val(Form2.Adodc11.Recordset.Fields!Money)
                w = Val(w) + Val(Form2.Adodc11.Recordset.Fields!emteyaz)
                Form2.Adodc11.Recordset.MoveNext
              Loop Until Form2.Adodc11.Recordset.EOF = True
              Form2.Adodc3.Recordset.Fields!numberpardakhtaghsat = Form2.Adodc11.Recordset.RecordCount
              Form2.Adodc3.Recordset.Fields!numberpardakhtnaghsat = Form2.Adodc10.Recordset.Fields!numberagsat - Form2.Adodc11.Recordset.RecordCount
              Form2.Adodc3.Recordset.Fields!bestankari = Val(Form2.Adodc10.Recordset.Fields!moneyvam) - q
              Form2.Adodc3.Recordset.Fields!emteyaz = w
            Else
              Form2.Adodc3.Recordset.Fields!numberpardakhtaghsat = 0
              Form2.Adodc3.Recordset.Fields!numberpardakhtnaghsat = Form2.Adodc10.Recordset.Fields!numberagsat
              Form2.Adodc3.Recordset.Fields!bestankari = Form2.Adodc10.Recordset.Fields!moneyvam
              Form2.Adodc3.Recordset.Fields!emteyaz = 0
            End If

            'ÂÎÑíä ÞÓØ
            e = 0
            If Form2.Adodc10.Recordset.Fields!tasvie = "äÔÏå" Then
              If Form2.Adodc11.Recordset.RecordCount > 0 Then
                Form2.Adodc11.Recordset.MoveLast
                a = Form2.Adodc11.Recordset.Fields!saragsat
              Else
                a = Amin.dateaminEzafeMoon(Form2.Adodc10.Recordset.Fields!Date, 1)
              End If

              If (a <= Form2.Label5.Caption) Then
                e = Val(Amin.dateaminEktelafmoon(a, Form2.Label5.Caption))
              End If

              If e > Val(Form2.Adodc3.Recordset.Fields!numberpardakhtnaghsat) Then
                e = Val(Form2.Adodc3.Recordset.Fields!numberpardakhtnaghsat)
              End If
            End If
            Form2.Adodc3.Recordset.Fields!numbermoaghsat = Val(e)
            Form2.Adodc3.Recordset.Fields!moneymo = Form2.Adodc3.Recordset.Fields!moneyghest2 * Val(e)
            Form2.Adodc3.Recordset.Update
            z = z + 1
            Form2.Adodc10.Recordset.MoveNext
          Loop Until Form2.Adodc10.Recordset.EOF = True
        Else
          Form2.Adodc3.Refresh
          Form2.Adodc3.Recordset.AddNew
          Form2.Adodc3.Recordset.Fields!id = Form3.Adodc1.Recordset.Fields!id
          Form2.Adodc3.Recordset.Fields!rad = z
          Form2.Adodc3.Recordset.Fields!idaccount = Form2.Adodc12.Recordset.Fields!id
          Form2.Adodc3.Recordset.Fields!moneyaccount = Form2.Adodc12.Recordset.Fields!Money
          Form2.Adodc3.Recordset.Fields!idvam = 0
          Form2.Adodc3.Recordset.Fields!moneyvam = 0
          Form2.Adodc3.Recordset.Fields!allghest = 0
          Form2.Adodc3.Recordset.Fields!moneyghest1 = 0
          Form2.Adodc3.Recordset.Fields!moneyghest2 = 0
          Form2.Adodc3.Recordset.Fields!karmozd = 0
          Form2.Adodc3.Recordset.Fields!numberallaghsat = 0
          Form2.Adodc3.Recordset.Fields!numberpardakhtaghsat = 0
          Form2.Adodc3.Recordset.Fields!numberpardakhtnaghsat = 0
          Form2.Adodc3.Recordset.Fields!bestankari = 0
          Form2.Adodc3.Recordset.Fields!emteyaz = 0
          Form2.Adodc3.Recordset.Fields!numbermoaghsat = 0
          Form2.Adodc3.Recordset.Fields!moneymo = 0
          Form2.Adodc3.Recordset.Update
          z = z + 1
        End If

        Form2.Adodc12.Recordset.MoveNext
      Loop Until Form2.Adodc12.Recordset.EOF = True
    End If
    Form3.Adodc1.Recordset.MoveNext
  Loop Until Form3.Adodc1.Recordset.EOF = True
End Sub

  

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTimbradoLayoutXML 
   Caption         =   "Ejemplo de Timbrado Layout y XML"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   7830
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   1800
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdTimbrarLayout 
      Caption         =   "Timbrar Layout"
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   3120
      Width           =   7095
   End
   Begin VB.CommandButton cmdExaminarLayout 
      Caption         =   "Examinar"
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtfile 
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Text            =   "--- Selecciona Layout ---"
      Top             =   2400
      Width           =   5655
   End
   Begin VB.CommandButton cmdTimbrarXML 
      Caption         =   "Timbrar XML"
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   7095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Timbrar Xml de Retenciones"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdExaminarXML 
      Caption         =   "Examinar"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtxml 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Text            =   "--- Selecciona XML ---"
      Top             =   600
      Width           =   5655
   End
End
Attribute VB_Name = "frmTimbradoLayoutXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Referencias:
' 1.- Windows Script Host Object Models
' 2.- CFD (Dll creada por facturación moderna)
' 3.- WSConecction (Dll creada por facturación moderna)
'
' Instalar OpenSSL Requerido
' 1.- Download openssl for windows
' http://gnuwin32.sourceforge.net/packages/openssl.htm
' 2.- Configurar variable de entorno para openssl

Dim fso As New FileSystemObject

Private Sub cmdExaminarLayout_Click()
    CommonDialog2.Filter = "Files TXT (*.txt)|*.txt|Files INI (*.ini)|*.ini"
    CommonDialog2.DefaultExt = "txt"
    CommonDialog2.DialogTitle = "Selecciona archivo"
    CommonDialog2.ShowOpen
    fname = CommonDialog2.FileName
    If fname = "" Then
        fname = "---- Selecciona layout ----"
    End If
    txtfile.Text = fname
End Sub

Private Sub cmdExaminarXML_Click()
    CommonDialog1.Filter = "Files XML (*.xml)|*.xml"
    CommonDialog1.DefaultExt = "xml"
    CommonDialog1.DialogTitle = "Selecciona archivo"
    CommonDialog1.ShowOpen
    fname = CommonDialog1.FileName
    If fname = "" Then
        fname = "---- Selecciona xml ----"
    End If
    txtxml.Text = fname
End Sub

Private Sub cmdTimbrarLayout_Click()
    Dim Path As String
    Dim obj_op As New opciones
    Dim obj_TC As New WSConnec
    Dim numero_certificado As String
    Dim archivo_cer As String
    Dim result As Resultados
    Dim archivo_pem As String
    Dim str_layout As String
    Dim str_filename As String
    Dim str_linea As String
    Dim str_file As String
    Dim outputPath As String
    
    Screen.MousePointer = 0
    Screen.MousePointer = vbHourglass

    Path = App.Path
    outputPath = Path + "\..\Resultados\"
    str_filename = Dir(txtfile.Text)
    If str_filename <> "" Then
        str_layout = txtfile.Text
    Else
        MsgBox ("No se encuentra el archivo layout")
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

    With obj_op
        .int_debug = 1
        .str_emisorRFC = "ESI920427886"
        .str_url = "https://t1demo.facturacionmoderna.com/timbrado/soap"
        .str_UserID = "UsuarioPruebasWS"
        .str_UserPass = "b9ec2afa3361a59af4b4d102d3f704eabdf097d4"
        .bol_generarCBB = False
        .bol_generarPDF = True
        .bol_generarTXT = False
        .str_log = Path + "\..\logs\FacturacionModerna-log.txt"
    End With

    Set result = obj_TC.timbrar(str_layout, obj_op)

    If result.Status = True Then
        MsgBox (result.message)
        
        'Almacenamiento del CFDI en formato xml
        Open outputPath + result.UUID + ".xml" For Binary Access Write As 1
        Put #1, , obj_TC.DecodeBase64(result.xmlB64)
        Close
        
        If obj_op.bol_generarPDF = True Then
            Open outputPath + result.UUID + ".pdf" For Binary Access Write As 2
            Put #2, , obj_TC.DecodeBase64(result.pdfB64)
            Close
        End If
        
        If obj_op.bol_generarCBB = True Then
            Open outputPath + result.UUID + ".png" For Binary Access Write As 3
            Put #3, , obj_TC.DecodeBase64(result.cbbB64)
            Close
        End If
        
        If obj_op.bol_generarTXT = True Then
            Open outputPath + result.UUID + ".txt" For Binary Access Write As 4
            Put #4, , obj_TC.DecodeBase64(result.txtB64)
            Close
        End If
        
    Else
        MsgBox (result.message)
    End If

    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdTimbrarXML_Click()
    Path = App.Path
    Dim keyfile As String
    keyfile = Path + "\..\utilerias\certificados\20001000000200000278.key"
    Dim certfile As String
    certfile = Path + "\..\utilerias\certificados\20001000000200000278.cer"
    Dim outPath As String
    outPath = Path + "\..\Resultados"
    Dim password As String
    password = "12345678a"
    Dim xmlfile As String
    xmlfile = txtxml.Text
    Dim xsltPath As String
    Dim comprobante As comprobante
    Set comprobante = New comprobante
    Dim cadenaO As String
    Dim cert As String
    Dim certNo As String
    Dim obj_op As New opciones
    Dim obj_TC As New WSConnec
    Dim result As Resultados
    Dim sello As String
    
    If Check1.Value = 1 Then
        xsltPath = Path + "\..\utilerias\xslt_retenciones\retenciones.xslt"
    Else
        xsltPath = Path + "\..\utilerias\xslt3_2\cadenaoriginal_3_2.xslt"
    End If
    
    Screen.MousePointer = 0
    Screen.MousePointer = vbHourglass

    If Not fso.FileExists(xmlfile) Then
        MsgBox ("No se encuentra el archivo " + xmlfile)
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    If Not fso.FileExists(keyfile) Then
        MsgBox ("No se encuentra el archivo " + keyfile)
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    If Not fso.FileExists(certfile) Then
        MsgBox ("No se encuentra el archivo " + certfile)
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    ' Obtener informacion del certificado
    If comprobante.getInfoCertificate(certfile) Then
        cert = comprobante.certificate
        certNo = comprobante.certificateNumber
    Else
        MsgBox (comprobante.message)
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

    ' Agregar numero y contenido del certificado al xml
    xmlfile = comprobante.addCertificateToXml(xmlfile, cert, certNo)
    If xmlfile = "" Then
        MsgBox (comprobante.message)
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    
    ' Crear cadena original
    cadenaO = comprobante.createOriginalChain(xmlfile, xsltPath)
    If cadenaO = "" Then
        MsgBox (comprobante.message)
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

    sello = comprobante.createDigitalStamp(keyfile, cadenaO, password)
    If sello = "" Then
        MsgBox (comprobante.message)
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

    xmlfile = comprobante.addDigitalStampToXml(xmlfile, sello)
    
    If xmlfile = "" Then
        MsgBox (comprobante.message)
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

    With obj_op
        .int_debug = 1
        .str_emisorRFC = "ESI920427886"
        .str_url = "https://t1demo.facturacionmoderna.com/timbrado/soap"
        .str_UserID = "UsuarioPruebasWS"
        .str_UserPass = "b9ec2afa3361a59af4b4d102d3f704eabdf097d4"
        .bol_generarCBB = False
        .bol_generarPDF = True
        .bol_generarTXT = False
        .str_log = Path + "\..\logs\FacturacionModerna-log.txt"
    End With

    Set result = obj_TC.timbrar(xmlfile, obj_op)
    outputPath = App.Path + "\..\Resultados\"
    
    If result.Status = True Then
        'Almacenamiento del CFDI en formato xml
        Open outputPath + result.UUID + ".xml" For Binary Access Write As 1
        Put #1, , obj_TC.DecodeBase64(result.xmlB64)
        Close
        
        If obj_op.bol_generarPDF = True Then
            Open outputPath + result.UUID + ".pdf" For Binary Access Write As 2
            Put #2, , obj_TC.DecodeBase64(result.pdfB64)
            Close
        End If
        
        If obj_op.bol_generarCBB = True Then
            Open outputPath + result.UUID + ".png" For Binary Access Write As 3
            Put #3, , obj_TC.DecodeBase64(result.cbbB64)
            Close
        End If
        
        If obj_op.bol_generarTXT = True Then
            Open outputPath + result.UUID + ".txt" For Binary Access Write As 4
            Put #4, , obj_TC.DecodeBase64(result.txtB64)
            Close
        End If
        MsgBox (result.message)
    Else
        MsgBox (result.message)
    End If
    
    Screen.MousePointer = vbNormal
End Sub

VERSION 5.00
Begin VB.Form frmMarca 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "mml_FRASE0888"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEspacio 
      Caption         =   "mml_FRASE0891"
      Height          =   495
      Left            =   2205
      TabIndex        =   8
      Top             =   2745
      Width           =   1665
   End
   Begin VB.CommandButton cmdMarca 
      Caption         =   "mml_FRASE0892"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   180
      TabIndex        =   7
      Top             =   2745
      Width           =   1665
   End
   Begin VB.CommandButton cmdFallo 
      Caption         =   "mml_FRASE0889"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   6
      Top             =   3330
      Width           =   3705
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   4200
      Begin VB.TextBox tbMarca 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2430
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2160
         Width           =   825
      End
      Begin VB.TextBox tbPuntos 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1590
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2160
         Width           =   795
      End
      Begin VB.TextBox tbBlanco 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   780
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2160
         Width           =   765
      End
      Begin VB.PictureBox picMarca 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         DrawStyle       =   2  'Dot
         DrawWidth       =   2
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   765
         ScaleHeight     =   78
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   161
         TabIndex        =   1
         Top             =   870
         Width           =   2475
      End
      Begin VB.Label lblMarca 
         Alignment       =   2  'Center
         Caption         =   "mml_FRASE0890"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   255
         TabIndex        =   2
         Top             =   270
         Width           =   3660
      End
   End
End
Attribute VB_Name = "frmMarca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iRetorno As Integer
Dim bEditarAlFinal As Boolean

Private Sub cmdEspacio_Click()
    iRetorno = 0
    Me.Hide
End Sub

Private Sub cmdFallo_Click()
    If tbPuntos.BackColor = vbRojoClaro Then
        tbPuntos.BackColor = vbAmarilloClaro
        bEditarAlFinal = False
    Else
        tbPuntos.BackColor = vbRojoClaro
        bEditarAlFinal = True
    End If
End Sub

Private Sub cmdMarca_Click()
    iRetorno = 1
    Me.Hide
End Sub

Public Function Visualizar(X As Integer, Y As Integer, iValorMarca As Integer, iPuntos As Integer, bFalloHoja As Boolean) As Integer
Dim cx As Integer, cy As Integer, ix As Integer, iy As Integer
Dim dEsX As Double, dEsY As Double
Dim lColor As Long
Dim iPos As Integer

    frmRecOptico.pbFicha.Visible = True
    frmRecOptico.mrcImg.Visible = False
    If bFalloHoja Then
        tbPuntos.BackColor = vbRojoClaro
        bEditarAlFinal = True
    Else
        tbPuntos.BackColor = vbAmarilloClaro
        bEditarAlFinal = False
    End If
    
    ProcesarEventos
    
    picMarca.Cls
    If iValorMarca = 1 Then
        cmdMarca.Default = True
    Else
        cmdEspacio.Default = True
    End If
    
    tbMarca.Text = C_MARCA
    tbBlanco.Text = C_BLANCO
    tbPuntos.Text = iPuntos

    lblMarca.Caption = mml_FRASE0372 & X & " , " & Y & mml_FRASE0893 & Chr$(13) & Chr$(10) & mml_FRASE0894

    dEsX = (picMarca.ScaleWidth / picMarca.Width)
    dEsY = (picMarca.ScaleHeight / picMarca.Height)
    
    ix = 10
    For cx = aMarcas(X, 0).iXi - C_MARGEN_CONTROL_ERROR_MARCA To aMarcas(X, 0).iXf + C_MARGEN_CONTROL_ERROR_MARCA
        ix = ix + 1
        iy = 10
        For cy = aMarcas(C_MAX_MARCAS_X, Y + 1).iYi - C_MARGEN_CONTROL_ERROR_MARCA To aMarcas(C_MAX_MARCAS_X, Y + 1).iYf + C_MARGEN_CONTROL_ERROR_MARCA
            iy = iy + 1
            
            lColor = ValorColor(ColorPunto(cx, cy))
            If lColor < C_UMBRAL Then
                picMarca.PSet (ix * dEsX * 20, iy * dEsY * 20), lColor
            End If
        Next
    Next
    
    DesplazarVisualizacionHoja Y
    Me.Show vbModal
    frmRecOptico.pbFicha.Top = iPos
    DesplazarVisualizacionHoja 0
    Visualizar = iRetorno

    If C_REC_OPTICO_RAPIDO Then
        frmRecOptico.pbFicha.Visible = False
        frmRecOptico.mrcImg.Visible = True
    End If
    ProcesarEventos
    bFalloHoja = bEditarAlFinal
End Function

Function ColorPunto(ix As Integer, iy As Integer) As Long
Dim i As Integer, dPteY As Double, dPteX As Double
Dim dEsX As Double, dEsY As Double

    dPteY = (aMarcas(C_MAX_MARCAS_X - 1, 0).iYi - aMarcas(0, 0).iYi) / (aMarcas(C_MAX_MARCAS_X - 1, 0).iXi - aMarcas(0, 0).iXi)
    dPteX = (aMarcas(C_MAX_MARCAS_X, C_MAX_MARCAS_Y - 1).iXi - aMarcas(C_MAX_MARCAS_X, 1).iXi) / (aMarcas(C_MAX_MARCAS_X, C_MAX_MARCAS_Y - 1).iYi - aMarcas(C_MAX_MARCAS_X, 1).iYi)
    
    dEsX = (frmRecOptico.pbFicha.ScaleWidth / frmRecOptico.pbFicha.Width)
    dEsY = (frmRecOptico.pbFicha.ScaleHeight / frmRecOptico.pbFicha.Height)
    
    ColorPunto = frmRecOptico.pbFicha.Point((ix + (iy - aMarcas(C_MAX_MARCAS_X - 1, 0).iYi) * dPteX) * dEsX, (iy + (ix - aMarcas(C_MAX_MARCAS_X - 1, 0).iXi) * dPteY) * dEsY)
    
End Function


Private Sub Form_Load()
    TraducirCadenas Me

End Sub

VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDeskripsi 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   1080
      TabIndex        =   6
      Top             =   1080
      Width           =   4455
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2625
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4260
      TabIndex        =   2
      Top             =   3075
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   4
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   5
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warning: ..."
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   3
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'dideklarasikan.

'Pilihan Reg Key Security ...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = _
      KEY_QUERY_VALUE + KEY_SET_VALUE + _
      KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
      KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
     'Tipe Reg Key ROOT ...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1     ' Unicode nul terminated string
Const REG_DWORD = 4  ' 32-bit number

Const gREGKEYSYSINFOLOC = _
      "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = _
      "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

'Konstanta untuk mengambil data pemilik & organisasi OS 'Windows di PC
Const gRegKeyLokasi1 = _
      "SOFTWARE\Microsoft\Windows\CurrentVersion"
Const gRegKeyOwner = "RegisteredOwner"
Const gRegKeyLokasi2 = _
      "SOFTWARE\Microsoft\Windows\CurrentVersion"
Const gRegKeyOrganization = "RegisteredOrganization"

Private Declare Function RegOpenKeyEx Lib _
"advapi32" Alias "RegOpenKeyExA" _
(ByVal hKey As Long, ByVal lpSubKey As String, _
ByVal ulOptions As Long, ByVal samDesired As Long, _
ByRef phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib _
"advapi32" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, _
ByVal lpReserved As Long, ByRef lpType As Long, _
ByVal lpData As String, ByRef lpcbData As Long) _
As Long

Private Declare Function RegCloseKey Lib _
"advapi32" (ByVal hKey As Long) As Long

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Versi " & App.Major & _
    "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    txtDeskripsi.Text = App.FileDescription
    StartOwner  'Ambil dan tampilkan nilai pemilik OS
                'Windows
End Sub

Public Sub StartOwner()
On Error GoTo SysInfoErr
  Dim rc As Long
  Dim Pemilik As String
  Dim Organisasi As String
  Dim lReturn As Long
  'Tampung nama pemilik OS Windows
  lReturn = GetKeyValue(HKEY_LOCAL_MACHINE, _
            gRegKeyLokasi1, gRegKeyOwner, Pemilik)
  lblOwner.Caption = Pemilik
  'Tampung nama organisasi OS Windows
  lReturn = GetKeyValue(HKEY_LOCAL_MACHINE, _
            gRegKeyLokasi2, gRegKeyOrganization, _
            Organisasi)
  lblOrganization.Caption = Organisasi
  Exit Sub
SysInfoErr:
  MsgBox "Tidak ada informasi pemilik Windows", _
         vbInformation, "NIHIL"
End Sub

Public Sub StartSysInfo()
On Error GoTo SysInfoErr
  Dim rc As Long
  Dim SysInfoPath As String
  'Ambil System Info Program Path\Name dari
  'Registry...
  If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, _
                 gREGVALSYSINFO, SysInfoPath) Then
  'Ambil hanya path System Info Program dari Registry.
  ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, _
         gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, _
         SysInfoPath) Then
      'Validasi keberadaan versi file 32 Bit
      If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
          SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
      'Error - File tidak dapat ditemukan...
      Else
          GoTo SysInfoErr
      End If
  'Error - Registry Entry tidak dapat ditemukan...
  Else
      GoTo SysInfoErr
  End If
  Call Shell(SysInfoPath, vbNormalFocus)
  Exit Sub
SysInfoErr:
  MsgBox "System Information Is Unavailable At This ", Time, vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, _
KeyName As String, SubKeyRef As String, _
ByRef KeyVal As String) As Boolean

    Dim i As Long      'Counter untuk looping
    Dim rc As Long     'Code pengembalian
    Dim hKey As Long   'Penanganan membuka Registry Key
    Dim hDepth As Long
    Dim KeyValType As Long  'Tipe Data Registry Key
    Dim tmpVal As String    'Penyimpanan sementara
                            'nilai Registry Key
    Dim KeyValSize As Long  'Ukuran variabel Registry
                            'Key

    'Buka RegKey di bawah KeyRoot
    '{HKEY_LOCAL_MACHINE...}
 
    'Buka Registry Key
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, _
         KEY_ALL_ACCESS, hKey)
    'Penanganan Error...
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    'Alokasi Variable Space
    tmpVal = String$(1024, 0)
    'Penanda Variable Size
    KeyValSize = 1024
    'Ambil Nilai Registry Key ...
    'Ambil/Buat nilai Key
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
         KeyValType, tmpVal, KeyValSize)
    
'Penanganan Errors
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    'Win95 Adds Null Terminated String...
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
        'Null ditemukan, Extract dari String
        tmpVal = Left(tmpVal, KeyValSize - 1)
    Else  'WinNT tidak bernilai Null Terminate String.
        'Null tidak ditemukan, Extract String saja
        tmpVal = Left(tmpVal, KeyValSize)
    End If

    'Memeriksa nilai tipe Key untuk konversi ...
    Select Case KeyValType  ' Cari tipe data...
    Case REG_SZ         'Tipe data string Registry Key
        KeyVal = tmpVal     'Copy nilai String
    Case REG_DWORD  'Tipe data Double Word Registry Key
        'Konversikan setiap bit
       For i = Len(tmpVal) To 1 Step -1
           'Bangun nilai Char. Dengan Char.
          KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
        Next
        'Konversi Double Word ke String
        KeyVal = Format$("&h" + KeyVal)
    End Select
    GetKeyValue = True      'Pengembalian sukses
    rc = RegCloseKey(hKey)  'Tutup Registry Key
    Exit Function           'Keluar dari fungsi
GetKeyError:   'Bersihkan memori jika terjadi error...
    KeyVal = ""      'Set Return Val ke string kosong
    GetKeyValue = False     'Pengembalian gagal
    rc = RegCloseKey(hKey)  'Tutup Registry Key
End Function


Private Sub lblDescription_Click()

End Sub

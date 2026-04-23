VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmEnctryp 
   Caption         =   "Enctryption"
   ClientHeight    =   4185
   ClientLeft      =   7980
   ClientTop       =   2700
   ClientWidth     =   7215
   Icon            =   "frmEnctryp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   7215
   Begin PT_XP_Button.PT_XP cmdDeEncrypt 
      Height          =   405
      Left            =   3915
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "DeEncrypt"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmEnctryp.frx":058A
      PictureAlignment=   5
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdEncrypt 
      Height          =   405
      Left            =   2355
      TabIndex        =   1
      Top             =   3570
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Encrypt"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmEnctryp.frx":08DC
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmEnctryp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

'Function GenerateHashMD5(FilePath As String) As String
'' /// Generates an MD5 hash of a file
''/// <param name="FilePath">File to hash</param>
''/// <returns>Uppercase HEX hash value</returns>
'End Function
'Function GenerateHashSHA384(FilePath As String) As String
'
''    /// Generates an SHA-384 hash of a file
''    /// <param name="FilePath">File to hash</param>
''    /// <returns>Uppercase HEX hash value</returns>
'End Function
 
'Function GenerateAES256Key(FilePathOut As String) As Integer
''    /// Generates one Symmetric Key File
''    /// <param name="FilePathOut">File Path for key, automatically adds extension ".symkey" to FilePathOut</param>
''    /// <returns>1 for success, 0 for failure</returns>
'End Function
'Function GenerateRSA512Keys(FilePathOut As String) As Integer
''    /// Generates one Public Key File and one Private Key File
''    /// <param name="FilePathOut">File Path for keys, automatically adds extensions ".pubkey" and ".prikey" respectively</param>
''    /// <returns>1 for success, 0 for failure</returns>
''    public int GenerateRSA512Keys(string FilePathOut)
'End Function
'Function EncryptRSA512(FilePathIn As String, KeyFilePath As String, olderPathOut As String) As Integer
'
''    /// Encrypts file(s) using RSA 512
''    /// <param name="FilePathIn">File(s) to encrypt, wildcards accepted</param>
''    /// <param name="KeyFilePath">File path of the keyfile, appends ".pubkey" extension if missing</param>
''    /// <param name="FolderPathOut">Destination folder for encrypted files. File named with original_name + ".encr"</param>
''    /// <returns>1 for success, 0 for failure</returns>
'End Function
'Function EncryptAES256(FilePathIn As String, KeyFilePath As String, FolderPathOut As String) As Integer
''    /// Encrypts file(s) using AES 256
''    /// <param name="FilePathIn">File(s) to encrypt, wildcards accepted</param>
''    /// <param name="KeyFilePath">File path of the keyfile, appends ".symkey" extension if missing</param>
''    /// <param name="FolderPathOut">Destination folder for encrypted files. File named with original_name + ".encr"</param>
''    /// <returns>1 for success, 0 for failure</returns>
'End Function
'Function DecryptRSA512(FilePathIn As String, KeyFilePath As String, FolderPathOut As String) As Integer
''    /// Decrypts file(s) using RSA 512
''    /// <param name="FilePathIn">File(s) to decrypt, wildcards accepted</param>
''    /// <param name="KeyFilePath">File path of the keyfile, appends ".prikey" if missing</param>
''    /// <param name="FolderPathOut">Destination folder for decrypted files. File named with original_name - ".encr"</param>
''    /// <returns>1 for success, 0 for failure</returns>
'End Function
'Function DecryptAES256(FilePathIn As String, KeyFilePath As String, FolderPathOut As String) As Integer
'    '/// Decrypts file(s) using AES 256
'    '/// <param name="FilePathIn">File(s) to decrypt, wildcards accepted</param>
'    '/// <param name="KeyFilePath">File path of the keyfile, appends ".symkey" if missing</param>
'    '/// <param name="FolderPathOut">Destination folder for decrypted files. File named with original_name - ".encr"</param>
'    '/// <returns>1 for success, 0 for failure</returns>
'
'End Function

Private Sub cmdDeEncrypt_Click()

End Sub
Private Sub cmdEncrypt_Click()
Dim dhitEnc As New DHITCOMHelper.Encryption
Dim strKeyFilePath As String
Dim lngResult As Long
Dim FSO As FileSystemObject
Dim FileName As file
Dim TextStream As TextStream
Dim strText As String


'
'    Set FSO = CreateObject("Scripting.FileSystemObject")
    With dhitEnc
        If Not FSO.FolderExists("C:\UnitTest") Then FSO.CreateFolder ("C:\UnitTest")
        strKeyFilePath = "C:\UnitTest\RSA512\KeyFile\RSA512Key"
        lngResult = .GenerateRSA512Keys(KeyFilePath)
    End With
 
''/* Generate the key file */
'string KeyFilePath = @"C:\UnitTest\RSA512\KeyFile\RSA512Key";
'result = dhitEnc.GenerateRSA512Keys(KeyFilePath);

''/* Encrypt the original file */
'string encryptInFilePath = @"C:\UnitTest\RSA512\Original\*.HL7";
'string encryptedOutFolder = @"C:\UnitTest\RSA512\Encrypted";
'result = dhitEnc.EncryptRSA512(encryptInFilePath, KeyFilePath, encryptedOutFolder);
'
''/* Decrypt the encrypted file */
'string decryptInFilePath = @"C:\UnitTest\RSA512\Encrypted\*.encr";
'string decryptedOutFolder = @"C:\UnitTest\RSA512\Decrypted";
'result = dhitEnc.DecryptRSA512(decryptInFilePath, KeyFilePath, decryptedOutFolder);
'
''/* compare the original file to the decrypted file */
'string checkOriginal = @"C:\UnitTest\RSA512\Original\Integrity.HL7";
'string checkDecrypted = @"C:\UnitTest\RSA512\Decrypted\Integrity.HL7";
'string checkOriginalFile = "";
'string checkDecryptedFile = "";
'using (StreamReader srXML = new StreamReader(checkOriginal)) {
'    checkOriginalFile = srXML.ReadToEnd();
'}
'using (StreamReader srXML = new StreamReader(checkDecrypted)) {
'    checkDecryptedFile = srXML.ReadToEnd();
'}
'Assert.AreEqual(checkOriginalFile, checkDecryptedFile);

 
    
    
    If FSO.FileExists("C:/TestIt.txt") Then
        Set FileName = FSO.GetFile("C:/TestIt.txt")
    Else
        'Do something else
    End If
    
    Set TextStream = FileName.OpenAsTextStream(ForReading, TristateUseDefault)
        
    strText = TextStream.ReadLine
    
    Debug.Print Mid(strText, 1, 1)
        
    Do While Not TextStream.AtEndOfStream
        strText = TextStream.ReadLine
    Loop
    
    Debug.Print Mid(strText, 1, 1)
    
    TextStream.Close
    Set TextStream = Nothing
    Set FSO = Nothing

    Exit Sub


End Sub

'Encryption class has the following exposed methods:
'
'NOTE: All int methods return 1 for success, 0 for failure
'string GenerateHashMD5(string FilePath);
'
'string GenerateHashSHA384(string FilePath);
'
'    int GenerateAES256Key(string FilePathOut);
'
'int GenerateRSA512Keys(string FilePathOut);
'
'int EncryptRSA512(string FilePathIn, string KeyFilePath, string FilePathOut);
'
'int EncryptAES256(string FilePathIn, string KeyFilePath, string FilePathOut);
'
'int DecryptRSA512(string FilePathIn, string KeyFilePath, string FilePathOut);
'
'int DecryptAES256(string FilePathIn, string KeyFilePath, string FilePathOut);
'
'Encryption methods are best explained with their XMLDoc and example usage.
'Note: for most of these, the input file is either a file or wildcard and the file-out is simply a folder path.
'
'/// Generates an MD5 hash of a file
'/// <param name="FilePath">File to hash</param>
'/// <returns>Uppercase HEX hash value</returns>
'public string GenerateHashMD5(string FilePath)
'
'/// Generates an SHA-384 hash of a file
'/// <param name="FilePath">File to hash</param>
'/// <returns>Uppercase HEX hash value</returns>
'public string GenerateHashSHA384(string FilePath)
'
'/// Generates one Symmetric Key File
'/// <param name="FilePathOut">File Path for key, automatically adds extension ".symkey" to FilePathOut</param>
'/// <returns>1 for success, 0 for failure</returns>
'public int GenerateAES256Key(string FilePathOut)
'
'/// Generates one Public Key File and one Private Key File
'/// <param name="FilePathOut">File Path for keys, automatically adds extensions ".pubkey" and ".prikey" respectively</param>
'/// <returns>1 for success, 0 for failure</returns>
'public int GenerateRSA512Keys(string FilePathOut)
'
'/// Encrypts file(s) using RSA 512
'/// <param name="FilePathIn">File(s) to encrypt, wildcards accepted</param>
'/// <param name="KeyFilePath">File path of the keyfile, appends ".pubkey" extension if missing</param>
'/// <param name="FolderPathOut">Destination folder for encrypted files. File named with original_name + ".encr"</param>
'/// <returns>1 for success, 0 for failure</returns>
'public int EncryptRSA512(string FilePathIn, string KeyFilePath, string FolderPathOut)
'
'/// Encrypts file(s) using AES 256
'/// <param name="FilePathIn">File(s) to encrypt, wildcards accepted</param>
'/// <param name="KeyFilePath">File path of the keyfile, appends ".symkey" extension if missing</param>
'/// <param name="FolderPathOut">Destination folder for encrypted files. File named with original_name + ".encr"</param>
'/// <returns>1 for success, 0 for failure</returns>
'public int EncryptAES256(string FilePathIn, string KeyFilePath, string FolderPathOut)
'
'/// Decrypts file(s) using RSA 512
'/// <param name="FilePathIn">File(s) to decrypt, wildcards accepted</param>
'/// <param name="KeyFilePath">File path of the keyfile, appends ".prikey" if missing</param>
'/// <param name="FolderPathOut">Destination folder for decrypted files. File named with original_name - ".encr"</param>
'/// <returns>1 for success, 0 for failure</returns>
'public int DecryptRSA512(string FilePathIn, string KeyFilePath, string FolderPathOut)
'
'/// Decrypts file(s) using AES 256
'/// <param name="FilePathIn">File(s) to decrypt, wildcards accepted</param>
'/// <param name="KeyFilePath">File path of the keyfile, appends ".symkey" if missing</param>
'/// <param name="FolderPathOut">Destination folder for decrypted files. File named with original_name - ".encr"</param>
'/// <returns>1 for success, 0 for failure</returns>
'public int DecryptAES256(string FilePathIn, string KeyFilePath, string FolderPathOut)
'Print
'Encryption example usage:
'
'Encryption dhitEnc = new Encryption();
'int result;
'
'/* Generate the key file */
'string KeyFilePath = @"C:\UnitTest\RSA512\KeyFile\RSA512Key";
'result = dhitEnc.GenerateRSA512Keys(KeyFilePath);
'
'/* Encrypt the original file */
'string encryptInFilePath = @"C:\UnitTest\RSA512\Original\*.HL7";
'string encryptedOutFolder = @"C:\UnitTest\RSA512\Encrypted";
'result = dhitEnc.EncryptRSA512(encryptInFilePath, KeyFilePath, encryptedOutFolder);
'
'/* Decrypt the encrypted file */
'string decryptInFilePath = @"C:\UnitTest\RSA512\Encrypted\*.encr";
'string decryptedOutFolder = @"C:\UnitTest\RSA512\Decrypted";
'result = dhitEnc.DecryptRSA512(decryptInFilePath, KeyFilePath, decryptedOutFolder);
'
'/* compare the original file to the decrypted file */
'string checkOriginal = @"C:\UnitTest\RSA512\Original\Integrity.HL7";
'string checkDecrypted = @"C:\UnitTest\RSA512\Decrypted\Integrity.HL7";
'string checkOriginalFile = "";
'string checkDecryptedFile = "";
'using (StreamReader srXML = new StreamReader(checkOriginal)) {
'    checkOriginalFile = srXML.ReadToEnd();
'}
'using (StreamReader srXML = new StreamReader(checkDecrypted)) {
'    checkDecryptedFile = srXML.ReadToEnd();
'}
'Assert.AreEqual(checkOriginalFile, checkDecryptedFile);

Private Sub Form_Load()

End Sub



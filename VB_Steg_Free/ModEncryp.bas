Attribute VB_Name = "ModEncryp"
Option Explicit
'******************************************************'
'------------------------------------------------------'
' Project: EzCryptoAPI v1.0.7
'
' Date: July-28-2001
'
' Programmer: Antonio Ramirez Cobos
'
' Module: Module1
'
' Description: Includes all the API stuff necessary for the
'              control to work.
'                               THIS IS IMPORTANT
'              Remember to register rsaenh.dll on your system's registry
'              using Regsvr32.dll.
'
'              From the Author:
'              'cause I consider myself in a continuous learning
'              path with no end on programming, please, if you
'              can improve this program
'              contact me at: *TONYDSPANIARD@HOTMAIL.COM*
'
'              I would be pleased to hear from your opinions,
'              suggestions, and/or recommendations. Also, if you
'              know something I don't know and wish to share it
'              with me, here you'll have your techy pal from Spain
'              that will do exactly the same towards you. If I can
'              help you in any way, just ask.
'
'              INTELLECTUAL COPYRIGHT STUFF [Is up to you anyway]
'              --------------------------------------------------
'              This code is copyright 2001 Antonio Ramirez Cobos
'              This code may be reused and modified for non-commercial
'              purposes only as long as credit is given to the author
'              in the programmes about box and it's documentation.
'              If you use this code, please email me at:
'              TonyDSpaniard@hotmail.com and let me know what you think
'              and what you are doing with it.
'
'              PS: Don't forget to vote for me buddy programmer!
'------------------------------------------------------'
'******************************************************'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Functions to handle Provider
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'CryptAcquireContext(
'HCRYPTPROV *phProv,    [out] The address to which the function copies a handle to the CSP.
'LPCTSTR pszContainer,  [in]  The key container name. This is a zero-terminated string that identifies the key container to the CSP
'                               vbNullChar used in this example to get the default key container.
'LPCTSTR pszProvider,   [in] The provider name. This is a zero-terminated string that specifies the CSP to be used
'DWORD dwProvType,      [in] The type of provider to acquire
'DWORD dwFlags          [in] The flag values
') As Long

Public Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
    phProv As Long, pszContainer As String, pszProvider As String, _
    ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
    
'CryptReleaseContext(
'HCRYPTPROV hProv,  [in] A handle to the application’s CSP
'DWORD dwFlags      [in] The flag values. This parameter is reserved for future use and should always be zero
');
Public Declare Function CryptReleaseContext Lib "advapi32.dll" _
    (ByVal hProv As Long, _
    ByVal dwFlags As Long) As Long
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Functions to handle and operate with Hash objects
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'CryptCreateHash(
'    HCRYPTPROV hProv,      [in] A handle to the CSP to use. An application obtains this handle using the CryptAcquireContext function.
'    ALG_ID Algid,          [in] An algorithm identifier of the hash algorithm to use
'    HCRYPTKEY hKey,        [in] For nonkeyed algorithms, this parameter should be set to zero
'    DWORD dwFlags,         [in] The flag values. This parameter is reserved for future use and should always be zero
'    HCRYPTHASH *phHash);   [out] The address to which the function copies a handle to the new hash object.
Public Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, _
    ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
'CryptDestroyHash(
'    HCRYPTHASH hHash [in] A handle to the hash object to be destroyed
');
Public Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
'CryptHashData(
'    HCRYPTHASH hHash,      [in] A handle to the hash object. An application obtains this handle using the CryptCreateHash function.
'    CONST BYTE *pbData,    [in] The address of the data to be hashed
'    DWORD dwDataLen,       [in] The number of bytes of data to be hashed.
'    DWORD dwFlags);        [in] The flag values (The Microsoft RSA Base Provider ignores this parameter)
Public Declare Function CryptHashData Lib "advapi32.dll" ( _
ByVal hHash As Long, ByVal pbdata As String, _
ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptBinHashData Lib "advapi32.dll" Alias "CryptHashData" ( _
ByVal hHash As Long, pbdata As Any, _
ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
'CryptGetHashParam(
'   HCRYPTHASH hHash,   [in] A handle to the hash object on which to query parameters
'   DWORD dwParam,      [in] The parameter number
'   BYTE *pbData,       [out] The parameter data buffer
'   DWORD *pdwDataLen,  [in/out] The address of the parameter data length.
'   DWORD dwFlags       [in] The flag values. This parameter is reserved for future use and should always be zero.
');
Public Declare Function CryptGetHashParam Lib "advapi32.dll" _
(ByVal hHash As Long, ByVal dwParam As Long, pbdata As Any, _
 pdwDataLen As Long, ByVal dwFlags As Long) As Long
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Key Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'CryptGenKey(
'    HCRYPTPROV hProv,      [in] A handle to the application’s CSP
'    ALG_ID Algid,          [in] The identifier for the algorithm for which the key is to be generated
'    DWORD dwFlags,         [in] The flags specifying the type of key generated. 0 or CRYPT_EXPORTABLE, CRYPT_CREATE_SALT
'    HCRYPTKEY *phKey);     [out] The address that the function copies the handle of the newly generated key to
Public Declare Function CryptGenKey Lib "advapi32.dll" ( _
    ByVal hProv As Long, ByVal Algid As Long, ByVal dwFlags As Long, phKey As Long) As Long
'CryptDeriveKey(
'    HCRYPTPROV hProv,      [in] A handle to the application’s CSP
'    ALG_ID Algid,          [in] The identifier for the algorithm for which the key is to be generated
'    HCRYPTHASH hBaseData,  [in] A handle to a hash object that has been fed exactly the base data
'    DWORD dwFlags, (0)     [in] The flags specifying the type of key generated
'    HCRYPTKEY *phKey);     [in/out] The address to which the function copies the handle of the newly generated key
Public Declare Function CryptDeriveKey Lib "advapi32.dll" ( _
    ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, _
    phKey As Long) As Long
'CryptExportKey(
'    HCRYPTKEY hKey,    [in] A handle to the key to be exported
'    HCRYPTKEY hExpKey, [in] A handle to a cryptographic key belonging to the destination user
'    DWORD dwBlobType,  [in] Type of key blob (i.e., SIMPLEBLOB)
'    DWORD dwFlags,
'    BYTE *pbData,
'    DWORD *pdwDataLen);
Public Declare Function CryptExportKey Lib "advapi32.dll" (ByVal hKey As Long, _
    ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, _
        pbdata As Any, pdwDataLen As Long) As Long
'CryptImportKey(
'HCRYPTPROV hProv,  [in] A handle to the application’s CSP
'BYTE *pbData,      [in] The buffer containing the key blob
'DWORD dwDataLen,   [in] The length, in bytes, of the key blob.
'HCRYPTKEY hImpKey, [in] The meaning of this parameter differs, depending on the CSP type and the type of key blob being imported (0 with SIMPLEBOB)
'DWORD dwFlags,     [in] The flag values
'HCRYPTKEY *phKey   [out] The address to which the function copies a handle to the key that was imported
');
Public Declare Function CryptImportKey Lib "advapi32.dll" ( _
ByVal hProv As Long, pbdata As Any, _
ByVal dwDataLen As Long, ByVal hImpKey As Long, ByVal dwFlags As Long, _
phKey As Long) As Long
'CryptDestroyKey(
'    HCRYPTKEY hKey);       [in] A handle to the key to be destroyed
Public Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Encrypting/Decrypting Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'CryptEncrypt(
'    HCRYPTKEY hKey,        [in] A handle to the key to use for the encryption
'    HCRYPTHASH hHash,      [in] A handle to a hash object
'    BOOL Final,            [in] The Boolean value that specifies whether this is the last section in a series being encrypted
'    DWORD dwFlags,         [in] The flag values. This parameter is reserved for future use and should always be zero
'    BYTE *pbData,          [in/out] The buffer holding the data to be encrypted
'    DWORD *pdwDataLen,     [in/out] The address of the data length
'    DWORD dwBufLen);       [in] The number of bytes in the pbData buffer
Public Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, _
    ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, pbdata As Any, _
    pdwDataLen As Long, ByVal dwBufLen As Long) As Long
'Variation for string Data blocks
Public Declare Function CryptStringEncrypt Lib "advapi32.dll" Alias "CryptEncrypt" (ByVal hKey As Long, _
    ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbdata As String, _
    pdwDataLen As Long, ByVal dwBufLen As Long) As Long

'CryptDecrypt(
'    HCRYPTKEY hKey,        [in] A handle to the key to use for the decryption
'    HCRYPTHASH hHash,      [in] A handle to a hash object
'    BOOL Final,            [in] The Boolean value that specifies whether this is the last section in a series being decrypted. This will be TRUE if this is the last or only block. If it is not, then it will be FALSE
'    DWORD dwFlags,         [in] The flag values. This parameter is reserved for future use and should always be zero
'    BYTE *pbData,          [in/out] The buffer holding the data to be decrypted. Once that decryption has been performed, the plaintext is placed back in this same buffer.
'    DWORD *pdwDataLen);    [in/out] The address of the data length
Public Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, _
    ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, pbdata As Any, _
    pdwDataLen As Long) As Long
'Variation for string Data blocks
Public Declare Function CryptStringDecrypt Lib "advapi32.dll" Alias "CryptDecrypt" (ByVal hKey As Long, _
    ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbdata As String, _
    pdwDataLen As Long) As Long
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Error Functions [Not used this time]
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Public Declare Function GetLastError Lib "kernel32" () As Long
'Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Get Temp Folder Path/File Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Draw Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Draw Constants
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
'Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_ADJUST = &H2000   ' Calculate the space left over.
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' File Function Constants
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Const MOVEFILE_REPLACE_EXISTING = &H1
Public Const MOVEFILE_COPY_ALLOWED = &H2
Public Const FILE_BEGIN = 0
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const CREATE_ALWAYS = 2
Public Const OPEN_ALWAYS = 4
Public Const TRUNCATE_EXISTING = 5
Public Const CREATE_NEW = 1
Public Const OPEN_EXISTING = 3
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_POINTER_FAIL = &HFFFFFFFF
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' My Error Constants
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Const ERROR_ILLEGAL_PROPERTY& = 1001&
Public Const ERROR_NO_HASH_CREATE& = 1002&
Public Const ERROR_NO_KEY_CONTAINER& = 1003&
Public Const ERROR_NO_HASH_CREATED& = 1004&
Public Const ERROR_NO_DIGEST& = 1005&
Public Const ERROR_NO_HASH_DATA& = 1006&
Public Const ERROR_FILE_NOT_FOUND& = 1007&
Public Const ERROR_NO_HASH_DESTROY& = 1008&
Public Const ERROR_NO_HASH_PASSW& = 1010&
Public Const ERROR_NO_KEY_DERIVED& = 1011&
Public Const ERROR_NO_DECRYPT& = 1012&
Public Const ERROR_NO_ENCRYPT& = 1009
Public Const ERROR_TMPPTH_NOT_FOUND& = 1013&
Public Const ERROR_ALGO_NOT_SUPP& = 1014&
Public Const ERROR_NO_TMP_FILE& = 1015&
Public Const ERROR_NO_FILE_OPEN& = 1016&
Public Const ERROR_NO_READ& = 1017&
Public Const ERROR_NO_WRITE& = 1018&
Public Const ERROR_NO_TMP_OPEN& = 1019&
Public Const ERROR_NOTHING_DIGESTED& = 1020&
Public Const ERROR_IS_DIR& = 1021&
Public Const NO_DATASET& = 0&
Public Const ERROR_FILE_SIZE& = 1022&
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Error Function Constants [Not used this time]
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Public Const FORMAT_MESSAGE_FROM_SYSTEM& = &H1000
'Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER& = &H100
'Public Const LANG_NEUTRAL& = &H0
'Public Const SUBLANG_DEFAULT& = &H1 '  user default
'Public Const ERROR_MSG_FAIL& = &H0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Name of the provider shipped with Windows by default
' Both support below declared hashing algorithms.
' MS_DEF_PROV is bundled with the operating system.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Const MS_DEF_PROV = "Microsoft Base Cryptographic Provider v1.0"
Public Const MS_ENHANCED_PROV = "Microsoft Enhanced Cryptographic Provider v1.0"
Public Const PROV_RSA_FULL = 1
Public Const CRYPT_NEWKEYSET = &H8

'CryptGetHashParam parameter number values
'Public Const HP_ALGID = 1
Public Const HP_HASHVAL = 2
'Public Const HP_HASHSIZE = 4
'
'Block size to read/write/encrypt/decrypt
Public Const HP_FILE_RW_BLOCKSIZE_1k = 1000& '[1kb at a time]
Public Const HP_FILE_RW_BLOCKSIZE_2k = 2000& '[2kb at a time]
Public Const HP_FILE_RW_BLOCKSIZE_4k = 4000& '[4kb at a time]
Public Const HP_FILE_RW_BLOCKSIZE_8k = 8000& '[8kb at a time]
Public Const HP_FILE_RW_BLOCKSIZE_16k = 16000& '[16kb at a time]
Public Const HP_FILE_RW_BLOCKSIZE_30k = 30000& '[30kb at a time]
Public Const HP_FILE_RW_BLOCKSIZE_40k = &H9C40  '[40kb at a time]
Public Const HP_FILE_RW_BLOCKSIZE_50k = &HC350  '[50kb at a time]
Public Const HP_FILE_RW_BLOCKSIZE_60k& = &HEA60 '[60kb]
Public Const HP_FILE_RW_BLOCKSIZE_80k& = &H13880 '[80kb]
Public Const HP_FILE_RW_BLOCKSIZE_100k& = &H186A0 '[100kb]
' Exported key blob definitions
'Public Const SIMPLEBLOB = 1
'Public Const PUBLICKEYBLOB = 6
'Public Const PRIVATEKEYBLOB = 7
'Public Const PLAINTEXTKEYBLOB = 8
'Algorithm classes
'Public Const ALG_CLASS_SIGNATURE = 8192
Public Const ALG_CLASS_DATA_ENCRYPT = 24576
Public Const ALG_CLASS_HASH = 32768
'Algorithm types
Public Const ALG_TYPE_ANY = 0
Public Const ALG_TYPE_BLOCK = 1536
Public Const ALG_TYPE_STREAM = 2048
'Block cipher sub ids
Public Const ALG_SID_DES = 1
Public Const ALG_SID_3DES = 3
Public Const ALG_SID_3DES_112 = 9
Public Const ALG_SID_RC2 = 2
'Stream cipher sub-ids
Public Const ALG_SID_RC4 = 1
'Hash sub ids
Public Const ALG_SID_MD2 = 1
Public Const ALG_SID_MD4 = 2
Public Const ALG_SID_MD5 = 3
Public Const ALG_SID_SHA = 4
Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
' Algorithm identifier definitions
' Hashing algorithms
Public Const CALG_MD2 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2)
Public Const CALG_MD4 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4)
Public Const CALG_MD5 = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Public Const CALG_SHA = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA)
' Encryption/Decryption algorithms
' Block ciphers
Public Const CALG_DES = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_DES)
Public Const CALG_3DES = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_3DES)
Public Const CALG_3DES_112 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_3DES_112)
Public Const CALG_RC2 = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK) Or ALG_SID_RC2)
' Stream ciphers
Public Const CALG_RC4 = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)
' Temp File
Public Const TEMP_FILE = "ECA"
Public Const TEMP_SIZE = 255&
' Default size
Public Const DEF_WIDTH = 930&
Public Const DEF_HEIGHT = 1080&
Public Const DEF_MAX_FILE_SIZE& = &H2710 '10MB [not used -can be any size]



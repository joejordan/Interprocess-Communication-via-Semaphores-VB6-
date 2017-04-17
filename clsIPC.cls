VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Interprocess Communication via Semaphores
' © Joe Jordan 2010-2011
'
' - v1.2.1 - (4/25/2011)
'         + Added optional DecrementSize to the Decrement function to reduce the
'             semaphore value a custom amount
'         - Cleaned up a few functions, removed some unnecessary variables and declares
'
' - v1.2 - (12/31/2010)
'         - Fixed QueryHandleCount function
'         - Added (currently unused) code to convert LARGE_INTEGER to VB Date, in case
'             one day Windows decides to start returning CreationTime for semaphores
'         - Removed VirtualAlloc declares and constants; using all Heap functions now
'
' - v1.1 - (12/28/2010)
'         + Added IsSemaphore function to quickly test for the existence of a semaphore
'         + Added HandleSecurity feature which prevents the closing of the semaphore
'             handle via CloseHandle
'         + Added caching of ValidateDLL result to improve speed
'         + Started QueryHandleCount function; if anyone can fix it please let me know
'         - Fixed logic error in getting the semaphore global state if we opened an
'             existing semaphore on initialize call
'
' - v1.0 - (12/3/2010) Initial Release
'
' Contact with questions, comments, improvements & suggestions: groups@ignitesoft.com
'
' Ah, semaphores, the staple of any good operating system. I searched for a VB6 example
' implementing the semaphore functions and came up mostly empty-handed. This class attempts
' to fill that missing gap in the world of VB6 examples.
'
' In developing this class, I initially thought it would be simple to create a global semaphore
' that all users would have access to. After all, the documentation states that:
' "The semaphore name can have a "Global\" or "Local\" prefix to explicitly create the
' object in the global or session name space."
' Little did I know that I would have to delve into the depths  security functions
' in order to actually provide *true* Global semaphore functionality. After many failed
' attempts, approaches and cryptic error messages (The revision level is unknown. WTF?),
' I believe that the infamous ACL dragon has, for our intentions at least, been slayed.
'
' I took the advice of one Anne Gunn and implemented some additional security for our global
' semaphore, so rogue applications can't steal our lunch money completely.
'
' The majority of the trial and error took place in finding the proper way to call and declare
' the security APIs. I thought I was 98% done, so I tested on XP to see if it worked there,
' as I figured if I could get it to work on Windows 7-64 bit, surely it would work in the
' UAC-less environment of XP. Well, it worked fine in the IDE, but spit out an invalid
' memory access error when compiled. I had gotten a similar error while testing in Windows 7
' and tracked it down to using the actual struct when calling CreateSemaphore rather than
' the pointer. So I had to go back through each call and test to see which one needed the
' actual struct instead of the pointer. Turns out it was SetSecurityDescriptorDacl that
' needed to accept an actual SECURITY_DESCRIPTOR rather than a pointer to one. After the 2nd
' such discovery, I went back and used the actual structs whenever possible as a precaution.
'
' The class was lightly tested on Windows 2000, XP, Vista and 7. If you come across any issues
' or have any improvements or suggestions please let me know.
'
'
' Credits:
' http://undocumented.ntinternals.net/for information on the undocumented NtQuerySemaphore
' API function and the (also undocumented) SEMAPHORE_QUERY_STATE permission constant.
'
' Anne Gunn for her excellent, thorough and well written article and accompanying code
' on creating a not-quite-null dacl, and explaining the benefits of doing so.
' http://www.codeguru.com/cpp/w-p/win32/tutorials/article.php/c4545
'
' Matts_User_Name of the SysInternals forums for the QueryName function.
' http://forum.sysinternals.com/handle-name-help-ntqueryObject_topic14435_page2.html
'
' IrfanAhmad on the MSDN forums for his thread on how to share a semaphore:
' http://social.msdn.microsoft.com/Forums/en/windowssdk/thread/335db156-b1f7-45e2-b3d1-f0e79e386744
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Semaphore related constants
Private Const MAX_PATH                  As Long = 260
Private Const INVALID_HANDLE_VALUE      As Long = -1
Private Const WAIT_OBJECT_0             As Long = 0
Private Const m_SemGlobalHeader         As String = "Global\"
Private Const m_SemLocalHeader          As String = "Local\"
Private Const SEMAQUERYINFOCLASS        As Long = 0 ' only one query type currently exists for semaphores
Private Const HANDLE_FLAG_PROTECT_FROM_CLOSE As Long = &H2

' Semaphore rights-related constants
Private Const WRITE_OWNER               As Long = &H80000
Private Const WRITE_DAC                 As Long = &H40000
Private Const SEMAPHORE_MODIFY_STATE    As Long = &H2
Private Const SEMAPHORE_QUERY_STATE     As Long = &H3
Private Const STANDARD_RIGHTS_REQUIRED  As Long = &HF0000
Private Const SYNCHRONIZE               As Long = &H100000
Private Const SEMAPHORE_ALL_ACCESS      As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or SEMAPHORE_QUERY_STATE)

' Heap memory allocation related constants
Private Const HEAP_ZERO_MEMORY          As Long = &H8

' Security related constants
Private Const SECURITY_DESCRIPTOR_REVISION = (1)
Private Const SECURITY_WORLD_SID_AUTHORITY = 1
Private Const ACL_REVISION = 2

' Error Reporting
Private Const FORMAT_MESSAGE_FROM_HMODULE   As Long = &H800
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000

' Various types used across the land...
Private Type ACE_HEADER
        AceType As Byte
        AceFlags As Byte
        AceSize As Long
End Type

Private Type ACCESS_ALLOWED_ACE
       Header As ACE_HEADER
       mask As Long
       SidStart As Long
End Type

Private Type ACCESS_DENIED_ACE
       Header As ACE_HEADER
       mask As Long
       SidStart As Long
End Type

Private Type ACL
       AclRevision As Byte
       Sbz1 As Byte
       AclSize As Integer
       AceCount As Integer
       Sbz2 As Integer
End Type

Private Type SECURITY_DESCRIPTOR
       Revision As Byte
       Sbz1 As Byte
       Control As Integer
       Owner As Long
       Group As Long
       sacl As ACL
       dacl As ACL
End Type

Private Type SID_IDENTIFIER_AUTHORITY
        Value(5) As Byte
End Type

Private Type SEMAPHORE_BASIC_INFORMATION
        CurrentCount As Long
        MaximumCount As Long
End Type

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Type UNICODE_STRING
        Length          As Integer
        MaximumLength   As Integer
        Buffer          As Long
End Type

Private Type OBJECT_NAME_INFORMATION  'OIC = 1
        ObjectName As UNICODE_STRING
End Type

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Type LARGE_INTEGER
        LowPart As Long
        HighPart As Long
End Type

Private Type OBJECT_BASIC_INFORMATION
        Attributes As Long
        GrantedAccess As Long
        HandleCount As Long
        ReferenceCount As Long
        PagedPoolUsage As Long
        NonPagedPoolUsage As Long
        Reserved(2) As Long
        NameInformationLength As Long
        TypeInformationLength As Long
        SecurityDescriptorLength As Long
        CreationTime As Currency   ' aka LARGE_INTEGER - Does not return anything as far as I can tell, too bad...
End Type

Private Enum OBJECT_INFORMATION_CLASS
        ObjectBasicInformation = 0
        ObjectNameInformation
        ObjectTypeInformation
        ObjectAllInformation
        ObjectDataInformation
End Enum

' Semaphore functions
Private Declare Function CreateSemaphore Lib "kernel32.dll" Alias "CreateSemaphoreA" (lpSemaphoreAttributes As SECURITY_ATTRIBUTES, ByVal lInitialCount As Long, ByVal lMaximumCount As Long, ByVal lpName As String) As Long
Private Declare Function OpenSemaphore Lib "kernel32.dll" Alias "OpenSemaphoreA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseSemaphore Lib "kernel32.dll" (ByVal hSemaphore As Long, ByVal lReleaseCount As Long, ByRef lpPreviousCount As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function SetHandleInformation Lib "kernel32.dll" (ByVal hObject As Long, ByVal dwMask As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetHandleInformation Lib "kernel32.dll" (ByVal hObject As Long, ByRef lpdwFlags As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

' Query semaphore information functions
Private Declare Function NtQueryObject Lib "ntdll.dll" (ByVal ObjectHandle As Long, ByVal OBJECT_INFORMATION_CLASS As Long, ByRef ObjectInformation As Any, ByVal ObjectInformationLength As Long, ByRef ReturnLength As Long) As Long
Private Declare Function NtQuerySemaphore Lib "ntdll.dll" (ByVal SemaphoreHandle As Long, ByVal SEMAPHORE_INFORMATION_CLASS As Long, ByRef SemaphoreInformation As Any, ByVal SemaphoreInformationLength As Long, ByRef ReturnLength As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" (ByVal lpFileTime As Currency, ByRef lpSystemTime As SYSTEMTIME) As Long

' Access control list, access control entries, security descriptors and more!
Private Declare Function InitializeAcl Lib "advapi32.dll" (ByVal pAcl As Long, ByVal nAclLength As Long, ByVal dwAclRevision As Long) As Long
Private Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" (ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal dwRevision As Long) As Long
Private Declare Function AddAccessDeniedAce Lib "advapi32.dll" (ByVal pAcl As Long, ByVal dwAceRevision As Long, ByVal AccessMask As Long, ByVal Psid As Long) As Long
Private Declare Function AddAccessAllowedAce Lib "advapi32.dll" (ByVal pAcl As Long, ByVal dwAceRevision As Long, ByVal AccessMask As Long, ByVal Psid As Long) As Long
Private Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" (ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bDaclPresent As Long, ByVal pDacl As Long, ByVal bDaclDefaulted As Long) As Long
Private Declare Function GetLengthSid Lib "advapi32.dll" (ByVal Psid As Long) As Long
Private Declare Function IsValidSid Lib "advapi32.dll" (ByVal Psid As Long) As Long
Private Declare Sub FreeSid Lib "advapi32.dll" (ByVal Psid As Long)
Private Declare Function AllocateAndInitializeSid Lib "advapi32.dll" ( _
        pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte, _
        ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, _
        ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, _
        ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, _
        ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, _
Psid As Long) As Long

' Memory allocation functions
Private Declare Sub CopyMemory Lib "kernel32.dll " Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32.dll" () As Long

' Test DLL functions
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandleA Lib "kernel32.dll" (ByVal lpModuleName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long

' Error testing
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

' Private member variables
Private m_SemaphoreHandle As Long       ' Stores our local semaphore handle
Private m_SemaphoreName As String       ' Stores our local semaphore name
Private m_SemaphoreFullName As String   ' Stores the full semaphore name (includes global or local header)
Private m_SemaphoreGlobal As Boolean    ' Stores whether we created the semaphore as local or global
Private m_SemaphoreInitialCount As Long ' The initial count of our semaphore
Private m_SemaphoreMaxCount As Long     ' The maximum count of our semaphore
Private m_bQSValidated As Long          ' Stores whether we found the undocumented NtQuerySemaphore function
Private m_bQOValidated As Long          ' Stores whether we found the undocumented NtQueryObject function
Private m_pDACL As Long                 ' Local pointer stored when we create our custom security descriptor


Public Property Get SemaphoreHandle() As Long
    SemaphoreHandle = m_SemaphoreHandle
End Property

Public Property Get SemaphoreName() As Long
    SemaphoreName = m_SemaphoreName
End Property

Public Property Get SemMaxCount()
    SemMaxCount = m_SemaphoreMaxCount
End Property

Public Property Get SemInitialCount() As Long
    SemInitialCount = m_SemaphoreInitialCount
End Property

Public Property Get IsSemGlobal()
    IsSemGlobal = m_SemaphoreGlobal
End Property

Public Property Get SemaphoreSecurityEnabled() As Boolean
    SemaphoreSecurityEnabled = IsHandleSecured(m_SemaphoreHandle)
End Property

Public Property Let SemaphoreSecurityEnabled(bEnabled As Boolean)
    Call SetHandleSecurity(m_SemaphoreHandle, bEnabled)
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This function initializes the semaphore settings and either creates or opens an existing semaphore.
' The optional EnableHandleSecurity will prevent CloseHandle from closing the semaphore.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Initialize(SemName As String, SemInitialCount As Long, SemMaxCount As Long, SemGlobal As Boolean, Optional EnableHandleSecurity As Boolean = False) As Boolean
    Dim hSemaphoreHandle As Long, pSD As Long, lRet As Long, bSuccess As Boolean
    Dim SecurityAttributes As SECURITY_ATTRIBUTES
    Dim SecurityDescriptor As SECURITY_DESCRIPTOR

    ' Exit if SemName is invalid (ValidateSemaphoreName also first attempts to correct input)
    If ValidateSemaphoreName(SemName) = False Then Exit Function

    ' Store member variables
    m_SemaphoreName = SemName
    m_SemaphoreGlobal = SemGlobal
    
    ' Assign global or local header to the passed sem name
    If SemGlobal = True Then
        m_SemaphoreFullName = m_SemGlobalHeader & SemName
    Else
        m_SemaphoreFullName = m_SemLocalHeader & SemName
    End If
    
    ' Initial semaphore count must be 0 or greater
    If SemInitialCount > -1 Then
        m_SemaphoreInitialCount = SemInitialCount
    Else
        m_SemaphoreInitialCount = 0
    End If
    
    ' Max semaphore count must be 1 or greater
    If SemMaxCount > 0 Then
        m_SemaphoreMaxCount = SemMaxCount
    Else
        m_SemaphoreMaxCount = 1
    End If
    
    ' Free any existing handle
    If m_SemaphoreHandle <> 0 Then
        Call CloseSemaphore
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' First check to see if we can open the semaphore. If so, it already exists and we get its existing values.
    ' SEMAPHORE_MODIFY_STATE allows us to increment the semaphore value. SYNCHRONIZE allows us to decrement the value,
    ' and SEMAPHORE_QUERY_STATE allows us access to query the semaphore's current values.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    hSemaphoreHandle = OpenSemaphore(SEMAPHORE_MODIFY_STATE Or SEMAPHORE_QUERY_STATE Or SYNCHRONIZE, False, m_SemaphoreFullName)
    
    If hSemaphoreHandle = 0 Or hSemaphoreHandle = INVALID_HANDLE_VALUE Then
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Opening an existing semaphore failed. We must now attempt to create one ourselves.
        ' The custom security descriptor is only necessary when we want our semaphore to be
        ' available across userspace boundaries (global), so we only initialize the security descriptor
        ' when creating a global semaphore.
        ' In a local semaphore, we can pass it null security attributes and it will create a local
        ' semaphore with default security.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If m_SemaphoreGlobal = True Then
            ' Create a security descriptor that will allow us to use the semaphore across user boundaries;
            SecurityDescriptor = GetSecurityDescriptor
            ' Prepare SecurityAttributes for passage to CreateSemaphore
            SecurityAttributes.lpSecurityDescriptor = VarPtr(SecurityDescriptor)
            SecurityAttributes.bInheritHandle = False
            SecurityAttributes.nLength = Len(SecurityAttributes)
            
            ' Attempt to create our global semaphpore
            If m_pDACL <> 0 Then
                hSemaphoreHandle = CreateSemaphore(SecurityAttributes, m_SemaphoreInitialCount, m_SemaphoreMaxCount, m_SemaphoreFullName)
            End If
            
        Else
            ' For local semaphores, we'll use the default security descriptor, we just need to
            ' initialize the SECURITY_ATTRIBUTES struct with its length.
            SecurityAttributes.nLength = Len(SecurityAttributes)
            
            ' Create a new local semaphore
            hSemaphoreHandle = CreateSemaphore(SecurityAttributes, m_SemaphoreInitialCount, m_SemaphoreMaxCount, m_SemaphoreFullName)
        End If
    
        If hSemaphoreHandle = INVALID_HANDLE_VALUE Or hSemaphoreHandle = 0 Then
            ' Create semaphore failed, check the debug window for what happened
            Call RaiseError
            Initialize = False
        Else
            ' CreateSemaphore success, secure handle if requested
            If EnableHandleSecurity = True Then
                Call SetHandleSecurity(hSemaphoreHandle, True)
            End If
            ' Store valid semaphore handle
            m_SemaphoreHandle = hSemaphoreHandle
            Initialize = True
        End If
    Else
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' We were able to open an existing semaphore, grab its values and set them to our member variables
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        m_SemaphoreHandle = hSemaphoreHandle
        m_SemaphoreGlobal = Not CBool(InStr(QueryName, "\Sessions\"))
        m_SemaphoreMaxCount = QueryMaxValue
        If EnableHandleSecurity = True Then
            ' Set handle security on existing handle
            Call SetHandleSecurity(hSemaphoreHandle, True)
        End If
        Initialize = True
    End If

    ' Finally, release our allocated heap memory
    If m_pDACL <> 0 Then
        Call HeapFree(GetProcessHeap, 0, m_pDACL)
        m_pDACL = 0
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This function creates a custom SECURITY_DESCRIPTOR we can use to define security
' settings for our semaphore.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetSecurityDescriptor() As SECURITY_DESCRIPTOR
    Dim lRet As Long
    Dim pEveryoneSID As Long
    Dim lACLSize As Long
    Dim SecurityDescriptor As SECURITY_DESCRIPTOR
    Dim udtSidIdentifierAuthority As SID_IDENTIFIER_AUTHORITY
    
    ' Initialize our security descriptor
    lRet = InitializeSecurityDescriptor(SecurityDescriptor, SECURITY_DESCRIPTOR_REVISION)
    
    If lRet <> 0 Then
        ' Init structure with SECURITY_WORLD_SID_AUTHORITY (aka "Everyone")
        udtSidIdentifierAuthority.Value(5) = SECURITY_WORLD_SID_AUTHORITY
        ' Get Everyone SID
        lRet = AllocateAndInitializeSid(udtSidIdentifierAuthority, 1, 0, 0, 0, 0, 0, 0, 0, 0, pEveryoneSID)

        If lRet <> 0 Then
            ' Check for valid SID
            lRet = IsValidSid(pEveryoneSID)
            If lRet <> 0 Then
                'Calculate the ACL size
                lACLSize = GetACLSize(pEveryoneSID)
                ' Allocate memory for the ACL
                m_pDACL = HeapAlloc(GetProcessHeap, HEAP_ZERO_MEMORY, lACLSize)
                If m_pDACL <> 0 Then
                    ' Initialize the DACL
                    lRet = InitializeAcl(m_pDACL, lACLSize, ACL_REVISION)
                    If lRet <> 0 Then
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        ' It is essential that the order of the AddAces stay how they are here,
                        ' else our security entries will not work as intended. We must first deny, then allow.
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        ' Adds an ace that denies WRITE_OWNER and WRITE_DAC to 'Everyone'
                        ' This denies everyone but the object owner access to the object's security.
                        lRet = AddAccessDeniedAce(m_pDACL, ACL_REVISION, WRITE_OWNER Or WRITE_DAC, pEveryoneSID)
                        If lRet <> 0 Then
                            lRet = SetSecurityDescriptorDacl(SecurityDescriptor, True, m_pDACL, False)
                            If lRet = 0 Then
                                GoTo ReportErr
                            End If
                        Else
                            GoTo ReportErr
                        End If
                        
                        ' Add the Everyone Sid to have all permissions, with the exception of the ones denied above
                        lRet = AddAccessAllowedAce(m_pDACL, ACL_REVISION, SEMAPHORE_ALL_ACCESS, pEveryoneSID)
                        If lRet <> 0 Then
                            lRet = SetSecurityDescriptorDacl(SecurityDescriptor, True, m_pDACL, False)
                            If lRet <> 0 Then
                                ' Both ACE's were added successfully
                                GetSecurityDescriptor = SecurityDescriptor
                            Else
                                GoTo ReportErr
                            End If
                        Else
                            GoTo ReportErr
                        End If
                        
                    Else
                        GoTo ReportErr
                    End If
                Else
                    GoTo ReportErr
                End If
            Else
                GoTo ReportErr
            End If
        Else
            GoTo ReportErr
        End If
    Else
        GoTo ReportErr
    End If
    
    If pEveryoneSID <> 0 Then
        Call FreeSid(pEveryoneSID)
    End If
    
    Exit Function
ReportErr:
    Call RaiseError
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Returns the size of our desired ACL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetACLSize(Psid As Long) As Long
    Dim tmpAclLength As Long
    Dim tmpACL As ACL
    Dim tmpAccessAllowedAce As ACCESS_ALLOWED_ACE
    Dim tmpAccessDeniedAce As ACCESS_DENIED_ACE
    
    ' For each ACE we add to the ACL, we need to allocate the size of the Sid.
    ' In our case, we're adding two ACE's, so we multiply SidLength by 2
    tmpAclLength = GetLengthSid(Psid) * 2
    
    ' Calculate the ACE sizes to be added to the ACL (an extra 4 bytes can be removed from each struct length)
    tmpAclLength = tmpAclLength + Len(tmpAccessAllowedAce) - 4
    tmpAclLength = tmpAclLength + Len(tmpAccessDeniedAce) - 4
    
    ' Also calculate the size of the ACL itself
    tmpAclLength = tmpAclLength + Len(tmpACL) - 4
    GetACLSize = tmpAclLength
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checks to see if the passed semaphore name exists on the system
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsSemaphore(SemName As String, SemGlobal As Boolean) As Boolean
    Dim hSemaphoreHandle As Long
    Dim tmpSemName As String
    
    ' Assign global or local header to the passed sem name
    If SemGlobal = True Then
        tmpSemName = m_SemGlobalHeader & SemName
    Else
        tmpSemName = m_SemLocalHeader & SemName
    End If
    
    ' Attempt to open semaphore with minimal rights, just to see if it exists
    hSemaphoreHandle = OpenSemaphore(SYNCHRONIZE, False, tmpSemName)
    'Call RaiseError
    If hSemaphoreHandle > 0 Then
        IsSemaphore = True
        Call CloseHandle(hSemaphoreHandle)
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''
' Increments the value of the local semaphore
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Increment(Optional IncrementSize As Long = 1, Optional ByRef RetPreviousVal As Long) As Boolean
    Dim lRet As Long
    Dim lPrevVal As Long
    
    If m_SemaphoreHandle <> 0 Then
        lRet = ReleaseSemaphore(m_SemaphoreHandle, IncrementSize, lPrevVal)
        If lRet <> 0 Then
            RetPreviousVal = lPrevVal
            Increment = True
        Else
            Increment = False
        End If
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''
' Decrements the value of the local semaphore
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Decrement(Optional DecrementSize As Long = 1) As Boolean
    Dim lRet As Long, i As Long
    
    If m_SemaphoreHandle <> 0 Then
    
        For i = 1 To DecrementSize
            lRet = WaitForSingleObject(m_SemaphoreHandle, 0)
        Next
        
        If lRet = WAIT_OBJECT_0 Then
            Decrement = True
        Else
            Decrement = False
        End If
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Returns the current value of a semaphore
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function QueryCurrentValue() As Long
    Dim SemInfo As SEMAPHORE_BASIC_INFORMATION
    Dim lRet As Long, lRetQuery As Long, lRetLen As Long
    
    If m_SemaphoreHandle <> 0 Then
        ' Tests for the existence of the undocumented NtQuerySemaphore function
        If m_bQSValidated = False Then
            m_bQSValidated = ValidateDLL("ntdll.dll", "NtQuerySemaphore")
        End If
        If m_bQSValidated = True Then
            ' Query semaphore values
            lRet = NtQuerySemaphore(m_SemaphoreHandle, SEMAQUERYINFOCLASS, SemInfo, Len(SemInfo), lRetLen)
            ' Store semaphore values
            lRetQuery = SemInfo.CurrentCount
            ' Return semaphore values
            QueryCurrentValue = lRetQuery
        Else
            ' NtQuerySemaphore function not present, unable to query value
            QueryCurrentValue = -1
        End If
    Else
        ' No open semaphore
        QueryCurrentValue = -1
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Returns max value of semaphore
' QueryValue can also check external semaphore handle
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function QueryMaxValue() As Long
    Dim SemInfo As SEMAPHORE_BASIC_INFORMATION
    Dim lRet As Long, lRetMax As Long, lRetLen As Long
    
    If m_SemaphoreHandle <> 0 Then
        ' Tests for the existence of the undocumented NtQuerySemaphore function
        If m_bQSValidated = False Then
            m_bQSValidated = ValidateDLL("ntdll.dll", "NtQuerySemaphore")
        End If
        If m_bQSValidated = True Then
            ' Query semaphore values
            lRet = NtQuerySemaphore(m_SemaphoreHandle, SEMAQUERYINFOCLASS, SemInfo, Len(SemInfo), lRetLen)
            ' Store semaphore values
            lRetMax = SemInfo.MaximumCount
            ' Return semaphore values
            QueryMaxValue = lRetMax
        End If
    Else
        ' No open semaphore
        QueryMaxValue = -1
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Returns current system-wide open handle count of the semaphore
'
' Credits: Based off of code from Matts_User_Name of the SysInternals forums
' http://forum.sysinternals.com/handle-name-help-ntqueryObject_topic14435_page2.html
'
' A special thanks to Matt for helping me debug this function as well. Apparently,
' NtQueryObject will return nonsense in ReturnLength when the allocated memory size (Len(OBI))
' is LARGER than what is needed. My OBJECT_BASIC_INFORMATION was one Reserved() larger than
' was necessary. I have since switched to the undocumented version of OBJECT_BASIC_INFORMATION,
' to expose additional semaphore properties if anyone is interested in those.
'
' NOTE: It appears that the OBJECT_BASIC_INFORMATION CreationTime property is not filled out
'       when querying a semaphore handle. Please correct me if I'm wrong though; I was unable
'       to get it to return anything besides 0.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function QueryHandleCount() As Long
    Dim pHeap As Long, pMem As Long, lInitSize As Long, lReqSize As Long, lRet As Long
    Dim OBI As OBJECT_BASIC_INFORMATION
    
    If m_SemaphoreHandle <> 0 Then
        ' Tests for the existence of the undocumented NtQueryObject function
        If m_bQOValidated = False Then
            m_bQOValidated = ValidateDLL("ntdll.dll", "NtQueryObject")
        End If
        
        If m_bQOValidated = True Then
        
            ' Measure size of struct
            lInitSize = Len(OBI)
            ' Allocate memory for OBI
            pMem = HeapAlloc(GetProcessHeap, HEAP_ZERO_MEMORY, lInitSize)
            
            ' Get required size; if we were right with lInitSize, we can get the information immediately,
            ' else a second call using lReqSize is necessary.
            lRet = NtQueryObject(m_SemaphoreHandle, ObjectBasicInformation, ByVal pMem, lInitSize, lReqSize)
            
            If lRet = 0 Then
                ' Success calling NtQueryObject on first try; Copy data buffer to UDT
                Call CopyMemory(OBI, ByVal pMem, lReqSize)
                ' Return open handle count from OBI object
                QueryHandleCount = OBI.HandleCount
                ' Free requested memory
                lRet = HeapFree(GetProcessHeap, 0, pMem)
                
            Else ' First NtQueryObject call failed; can use NTRaiseError to find out the specific cause
            
                ' Free requested memory
                lRet = HeapFree(GetProcessHeap, 0, pMem)
                ' Abort if NtQueryObject fails to inform us of the incoming structure size
                If lReqSize = 0 Then
                    QueryHandleCount = -1
                    Exit Function
                End If

                ' 2nd attempt; allocate requested amount of memory
                pMem = HeapAlloc(GetProcessHeap, HEAP_ZERO_MEMORY, lReqSize)
                
                ' 2nd API Call - Query Data with previously requested data size
                lRet = NtQueryObject(m_SemaphoreHandle, ObjectBasicInformation, ByVal pMem, lReqSize, ByVal 0)
                If lRet = 0 Then
                    ' Success! Copy data buffer to UDT
                    Call CopyMemory(OBI, ByVal pMem, lReqSize)
                    ' Return open handle count from OBI object
                    QueryHandleCount = OBI.HandleCount
                Else
                    QueryHandleCount = -1
                End If
                
                ' Free requested memory
                lRet = HeapFree(GetProcessHeap, 0, pMem)
            End If
        Else
            ' NtQueryObject function does not exist!
            QueryHandleCount = -1
        End If
            
    Else
        ' No semaphore open in our class object
        QueryHandleCount = -1
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Returns the name of a semaphore from its handle; if hSemaphoreHandle is 0,
' it will return the name of the local semaphore
'
' Credits: Based off of code from Matts_User_Name of the SysInternals forums
' http://forum.sysinternals.com/handle-name-help-ntqueryObject_topic14435_page2.html
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function QueryName() As String
    Dim pHeap As Long, lInitSize As Long, lReqSize As Long, lRet As Long
    Dim ONI As OBJECT_NAME_INFORMATION
    
    If m_SemaphoreHandle <> 0 Then
        ' Tests for the existence of the undocumented NtQueryObject function
        If m_bQOValidated = False Then
            m_bQOValidated = ValidateDLL("ntdll.dll", "NtQueryObject")
        End If
        
        If m_bQOValidated = True Then
        
            ' Measure size of struct
            lInitSize = Len(ONI)
            ' Allocate memory for OBI
            pHeap = HeapAlloc(GetProcessHeap, HEAP_ZERO_MEMORY, lInitSize)
            ' Get required size
            lRet = NtQueryObject(m_SemaphoreHandle, ObjectNameInformation, ByVal pHeap, lInitSize, lReqSize)
            ' Free requested memory
            lRet = HeapFree(GetProcessHeap, 0, pHeap)
    
            ' Abort if NtQueryObject fails to inform us of the incoming structure size
            If lReqSize = 0 Then
                QueryName = ""
                Exit Function
            End If
            
            ' Allocate requested amount of memory
            pHeap = HeapAlloc(GetProcessHeap, HEAP_ZERO_MEMORY, lReqSize)
            ' 2nd API Call - Query Data
            lRet = NtQueryObject(m_SemaphoreHandle, ObjectNameInformation, ByVal pHeap, lReqSize, ByVal 0)
            ' Copy data buffer to UDT
            Call CopyMemory(ONI, ByVal pHeap, Len(ONI))
            ' Return data from OBI
            QueryName = US2Str(ONI.ObjectName)
            ' Free memory
            lRet = HeapFree(GetProcessHeap, 0, pHeap)
        Else
            ' NtQueryObject function does not exist!
            QueryName = ""
        End If
    Else
        QueryName = ""
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''
' Converts a UNICODE_STRING to a VB String
'''''''''''''''''''''''''''''''''''''''''''''''''
Private Function US2Str(US As UNICODE_STRING) As String
    If US.Length <> 0 And US.Buffer <> 0 And US.Length <> 0 Then
        Dim sBuffer As String: sBuffer = Space$(US.Length \ 2)
        Call CopyMemory(ByVal StrPtr(sBuffer), ByVal US.Buffer, US.Length)
        US2Str = sBuffer
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''
' Sets the semaphore handle security settings
'''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetHandleSecurity(hSemaphoreHandle As Long, HandlePreventClose As Boolean) As Boolean
    Dim lRet As Long
    
    If HandlePreventClose = True Then
        ' Turn on handle security
        lRet = SetHandleInformation(hSemaphoreHandle, HANDLE_FLAG_PROTECT_FROM_CLOSE, HANDLE_FLAG_PROTECT_FROM_CLOSE)
    Else
        ' Shut off handle security
        lRet = SetHandleInformation(hSemaphoreHandle, HANDLE_FLAG_PROTECT_FROM_CLOSE, 0)
    End If
    
    If lRet <> 0 Then
        SetHandleSecurity = True
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets the current semaphore handle security settings
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsHandleSecured(hSemaphoreHandle As Long) As Boolean
    Dim lRet As Long, lFlags As Long

    lRet = GetHandleInformation(hSemaphoreHandle, lFlags)
    If lRet <> 0 Then
        If lFlags = HANDLE_FLAG_PROTECT_FROM_CLOSE Then
            IsHandleSecured = True
        Else
            IsHandleSecured = False
        End If
    End If

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Closes the semaphore handle; disables semaphore security if ForceClose = True.
' Note: This does not affect the semaphore count, however if this was the last handle open
' for the semaphore, the system will close it.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CloseSemaphore(Optional ForceClose As Boolean = False) As Boolean
    Dim bRet As Boolean, lRet As Long
    
    If ForceClose = True Then
        If IsHandleSecured(m_SemaphoreHandle) = True Then
            ' Turn off handle security
            bRet = SetHandleSecurity(m_SemaphoreHandle, False)
        End If
    End If
    
    lRet = CloseHandle(m_SemaphoreHandle)
    If lRet <> 0 Then
        CloseSemaphore = True
        m_SemaphoreHandle = 0
    Else
        CloseSemaphore = False
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' A private helper function to validate our semaphore input name.
' Takes the sem name byref and corrects it if possible
' Probably not exhaustive, but good enough for most solutions.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateSemaphoreName(ByRef SemName As String) As Boolean

    ' Semaphore names cannot contain a backslash
    If InStr(SemName, "\") <> 0 Then
        SemName = Replace(SemName, "\", "")
    End If
    
    ' Semaphore names cannot be longer than MAX_PATH
    If Len(SemName) > MAX_PATH Then
        SemName = Left$(SemName, MAX_PATH)
    End If

    ' Named semaphores cannot be empty
    If Trim$(SemName) = "" Then
        ValidateSemaphoreName = False
    Else
        ValidateSemaphoreName = True
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Converts the LARGE_INTEGER struct into a VB Date
' This will be useful if Windows ever decides to return a CreationTime when
' querying the semaphore; as of now, it is unused.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LargeIntegerToDate(CreationTime As Currency) As Date
    Dim stSystemTime As SYSTEMTIME
    
    ' Convert the FILETIME into a SYSTEMTIME.
    Call FileTimeToSystemTime(CreationTime, stSystemTime)

    ' Convert the returned SYSTEMTIME into a VB Date
    With stSystemTime
        LargeIntegerToDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''
' Checks a DLL for a specific function
'
' NOTE: This class only checks ntdll, which is already open in every NT
'       application, so no LoadLibrary is necessary
'''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateDLL(DLLName As String, DLLProc As String) As Boolean
Dim lProc As Long

    ' Retrieve the address of the function call
    lProc = GetProcAddress(GetModuleHandleA(DLLName), DLLProc)
    ' return result
    ValidateDLL = (Not lProc = 0)
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Prints out error messages; very helpful in debugging this solution
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RaiseError() As Boolean
    Dim sErrMsg As String, lRet As Long
    ' Allocate string for message
    sErrMsg = Space$(1024)
    lRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, Err.LastDllError, 0&, sErrMsg, Len(sErrMsg), ByVal 0&)
    If lRet <> 0 Then
        Debug.Print sErrMsg
        'MsgBox sErrMsg
        RaiseError = True
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Prints out NTSTATUS error messages returned by NtQueryObject
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NTRaiseError(NTStatusMessage As Long) As Boolean
    Dim lpMessageBuffer As Long
    Dim sErrMsg As String
    Dim hNTDLL As Long
    Dim lRet As Long
    ' Allocate string for message
    sErrMsg = Space$(1024)
    
    hNTDLL = LoadLibrary("NTDLL.DLL")
    
    lRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_FROM_HMODULE, _
                        ByVal hNTDLL, NTStatusMessage, MakeLangID, sErrMsg, Len(sErrMsg), ByVal 0&)
    
    If lRet <> 0 Then
        Debug.Print sErrMsg
    End If

   Call FreeLibrary(hNTDLL)

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Generates a US-English Language ID for NTRaiseError
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MakeLangID() As Long
    Dim SubLang As Integer, Lang As Integer
    
    Lang = 9
    SubLang = 1

    MakeLangID = (SubLang * (2 ^ 10)) Or Lang
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Close the semaphore handle upon object termination
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    Call CloseSemaphore
End Sub


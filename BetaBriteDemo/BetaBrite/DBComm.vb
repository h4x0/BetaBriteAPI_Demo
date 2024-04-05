Option Explicit On 
Option Strict On

Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading

'TODO: The number of data bits must be 5 to 8 bits.
'TODO: The use of 5 data bits with 2 stop bits is an invalid combination, as is 6, 7, or 8 data bits with 1.5 stop bits.

''' <summary>
''' DBComm serial library (C) 2003 Cory Smith http://www.addressof.com/
''' Note that .NET 2.0 will have a native Serial component
''' </summary>
''' <remarks>
'''   http://workspaces.gotdotnet.com/dbcomm
''' </remarks>
Public Class RS232
  Implements IDisposable

#Region "Win32API"

  Private Const CE_BREAK As Integer = &H10    ' The hardware detected a break condition. 
  Private Const CE_FRAME As Integer = &H8     ' The hardware detected a framing error. 
  Private Const CE_IOE As Integer = &H400     ' An I/O error occurred during communications with the device. 
  Private Const CE_MODE As Integer = &H8000   ' The requested mode is not supported, or the hFile parameter is invalid. If this value is specified, it is the only valid error. 
  Private Const CE_OVERRUN As Integer = &H2   ' A character-buffer overrun has occurred. The next character is lost. 
  Private Const CE_RXOVER As Integer = &H1    ' An input buffer overflow has occurred. There is either no room in the input buffer, or a character was received after the end-of-file (EOF) character. 
  Private Const CE_RXPARITY As Integer = &H4  ' The hardware detected a parity error. 
  Private Const CE_TXFULL As Integer = &H100  ' The application tried to transmit a character, but the output buffer was full. 

  Private Const PURGE_RXABORT As Integer = &H2
  Private Const PURGE_RXCLEAR As Integer = &H8
  Private Const PURGE_TXABORT As Integer = &H1
  Private Const PURGE_TXCLEAR As Integer = &H4

  Private Const GENERIC_READ As Integer = &H80000000
  Private Const GENERIC_WRITE As Integer = &H40000000

  Private Const OPEN_EXISTING As Integer = 3

  Private Const INVALID_HANDLE_VALUE As Integer = -1
  Private Const IO_BUFFER_SIZE As Integer = 1024
  Private Const FILE_FLAG_OVERLAPPED As Integer = &H40000000
  Private Const ERROR_IO_PENDING As Integer = 997
  Private Const WAIT_OBJECT_0 As Integer = 0
  Private Const ERROR_IO_INCOMPLETE As Integer = 996
  Private Const WAIT_TIMEOUT As Integer = &H102&
  Private Const INFINITE As Integer = &HFFFFFFFF

  'Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Integer, ByVal lpSource As Integer, ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, ByVal lpBuffer As StringBuilder, ByVal nSize As Integer, ByVal Arguments As Integer) As Integer

  <Flags()> Private Enum DCBFlags As Integer
    ' Specifies if binary mode is enabled. The Win32 API does not support 
    ' nonbinary mode transfers, so this member must be TRUE. Using FALSE will 
    ' not work. 
    Binary = 1
    ' Specifies if parity checking is enabled. If this member is TRUE, 
    ' parity checking is performed and errors are reported. 
    Parity = 2
    ' Specifies if the CTS (clear-to-send) signal is monitored for output 
    ' flow control. If this member is TRUE and CTS is turned off, output 
    ' is suspended until CTS is sent again. 
    OutXCTSFlow = 4
    ' Specifies if the DSR (data-set-ready) signal is monitored for output 
    ' flow control. If this member is TRUE and DSR is turned off, output is 
    ' suspended until DSR is sent again. 
    OutXDSRFlow = 8
    ' Specifies the DTR (data-terminal-ready) flow control. This member can 
    ' be one of the following values:
    ' DTR_CONTROL_DISABLE - Disables the DTR line when the device is opened and leaves it disabled.
    ' DTR_CONTROL_ENABLE - Enables the DTR line when the device is opened and leaves it on. 
    ' DTR_CONTROL_HANDSHAKE - Enables DTR handshaking. If handshaking is enabled, it is an error for the application to adjust the line by using the EscapeCommFunction function.
    DTRControl1 = 16
    DTRControl2 = 32
    ' Specifies if the communications driver is sensitive to the state of 
    ' the DSR signal. If this member is TRUE, the driver ignores any bytes 
    ' received, unless the DSR modem input line is high. 
    DSRSensitivity = 64
    ' Specifies if transmission stops when the input buffer is full and the 
    ' driver has transmitted the XoffChar character. If this member is TRUE, 
    ' transmission continues after the input buffer has come within XoffLim 
    ' bytes of being full and the driver has transmitted the XoffChar character 
    ' to stop receiving bytes. If this member is FALSE, transmission does not 
    ' continue until the input buffer is within XonLim bytes of being empty and 
    ' the driver has transmitted the XonChar character to resume reception. 
    TXContinueOnXOff = 128
    ' Specifies if XON/XOFF flow control is used during transmission. If this 
    ' member is TRUE, transmission stops when the XoffChar character is received 
    ' and starts again when the XonChar character is received. 
    OutX = 256
    ' Specifies if XON/XOFF flow control is used during reception. If this member 
    ' is TRUE, the XoffChar character is sent when the input buffer comes within 
    ' XoffLim bytes of being full, and the XonChar character is sent when the 
    ' input buffer comes within XonLim bytes of being empty. 
    InX = 512
    ' Specifies if bytes received with parity errors are replaced with the 
    ' character specified by the ErrorChar member. If this member is TRUE and 
    ' the ParityCheck member is TRUE, replacement occurs. 
    ErrorChar = 1024
    ' Specifies if null bytes are discarded. If this member is TRUE, null bytes 
    ' are discarded when received. 
    Null = 2048
    ' Specifies the RTS (request-to-send) flow control. If this value is zero, 
    ' the default is RTS_CONTROL_HANDSHAKE. This member can be one of the 
    ' following values:
    ' RTS_CONTROL_DISABLE - Disables the RTS line when the device is opened and leaves it disabled.
    ' RTS_CONTROL_ENABLE - Enables the RTS line when the device is opened and leaves it on.
    ' RTS_CONTROL_HANDSHAKE - Enables RTS handshaking. The driver raises the RTS line when the “type-ahead” (input) buffer is less than one-half full and lowers the RTS line when the buffer is more than three-quarters full. If handshaking is enabled, it is an error for the application to adjust the line by using the EscapeCommFunction function.
    ' RTS_CONTROL_TOGGLE - Specifies that the RTS line will be high if bytes are available for transmission. After all buffered bytes have been sent, the RTS line will be low.
    RTSControl1 = 4096
    RTSControl2 = 8192
    ' Specifies if read and write operations are terminated if an error occurs. 
    ' If this member is TRUE, the driver terminates all read and write operations 
    ' with an error status if an error occurs. The driver will not accept any 
    ' further communications operations until the application has acknowledged 
    ' the error by calling the ClearCommError function. 
    AbortOnError = 16384
    ' Reserved; do not use
    'Dummy2 = 17 bits
  End Enum

  ' This is the DCB structure used by the calls to the Windows API.
  <StructLayout(LayoutKind.Sequential, Pack:=1)> Private Structure DCB
    ' Specifies the DCB structure length, in bytes. 
    Public DCBlength As Integer
    ' Specifies the baud rate at which the communication device operates. 
    ' It is an actual baud rate value, or one of the following baud rate indexes: 
    ' Note: See BaudRates Enum for complete list.
    Public BaudRate As Integer

    Public Flags As DCBFlags 'Integer

    ' Not used; set to zero. 
    Public Reserved As Short

    ' Specifies the minimum number of bytes accepted in the input buffer before 
    ' the XON character is sent. 
    Public XonLimit As Short
    ' Specifies the maximum number of bytes accepted in the input buffer before 
    ' the XOFF character is sent. The maximum number of bytes accepted is 
    ' calculated by subtracting this value from the size, in bytes, of the 
    ' input buffer. 
    Public XoffLimit As Short
    ' Specifies the number of bits in the bytes transmitted and received. 
    Public ByteSize As Byte
    ' Specifies the parity scheme to be used. It is one of the following values: 
    ' EVENPARITY
    ' MARKPARITY
    ' NOPARITY
    ' ODDPARITY
    ' SPACEPARITY
    Public Parity As Byte
    ' Specifies the number of stop bits to be used. It is one of the following values: 
    ' ONESTOPBIT - 1
    ' ONE5STOPBITS - 1.5
    ' TWOSTOPBITS - 2
    Public StopBits As Byte
    ' Specifies the value of the XON character for both transmission and reception. 
    Public XonChar As Byte
    ' Specifies the value of the XOFF character for both transmission and reception. 
    Public XoffChar As Byte
    ' Specifies the value of the character used to replace bytes received with a parity error. 
    Public ErrorChar As Byte
    ' Specifies the value of the character used to signal the end of data. 
    Public EofChar As Byte
    ' Specifies the value of the character used to signal an event. 
    Public EvtChar As Byte
    ' Reserved; do not use. 
    Public Reserved1 As Short

  End Structure

  <Flags()> Private Enum COMSTATFlags As Integer
    ' If this member is TRUE, transmission is waiting for the CTS (clear-to-send) 
    ' signal to be sent. 
    CTSHold = 1
    ' If this member is TRUE, transmission is waiting for the DSR (data-set-ready)
    ' signal to be sent. 
    DSRHold = 2
    ' If this member is TRUE, transmission is waiting for the RLSD (receive-line-
    ' signal-detect) signal to be sent. 
    RLSDHold = 4
    ' If this member is TRUE, transmission is waiting because the XOFF character
    ' was received. 
    XOffHold = 8
    ' If this member is TRUE, transmission is waiting because the XOFF character
    ' was transmitted. (Transmission halts when the XOFF character is transmitted
    ' to a system that takes the next character as XON, regardless of the actual
    ' character.) 
    XOffSent = 16
    ' If this member is TRUE, the end-of-file (EOF) character has been received. 
    EOF = 32
    ' If this member is TRUE, there is a character queued for transmission that
    ' has come to the communications device by way of the TransmitCommChar function.
    ' The communications device transmits such a character ahead of other
    ' characters in the device's output buffer. 
    Txim = 64
    ' Reserved; do not use. 
    'Reserved = 25 bits
  End Enum

  <StructLayout(LayoutKind.Sequential, Pack:=1)> Private Structure COMSTAT
    ' See COMSTATFlags for flag documenation
    Public Flags As COMSTATFlags
    ' Number of bytes received by the serial provider but not yet read by a ReadFile operation. 
    Public InQue As Integer
    ' Number of bytes of user data remaining to be transmitted for all write operations. This value will be zero for a nonoverlapped write. 
    Public OutQue As Integer
  End Structure

  <StructLayout(LayoutKind.Sequential, Pack:=1)> Private Structure COMMTIMEOUTS

    ' Maximum time allowed to elapse between the arrival of two characters on 
    ' the communications line, in milliseconds. During a ReadFile operation, 
    ' the time period begins when the first character is received. If the 
    ' interval between the arrival of any two characters exceeds this amount, 
    ' the ReadFile operation is completed and any buffered data is returned. A 
    ' value of zero indicates that interval time-outs are not used. 
    ' A value of MAXDWORD, combined with zero values for both the 
    ' ReadTotalTimeoutConstant and ReadTotalTimeoutMultiplier members, specifies 
    ' that the read operation is to return immediately with the characters that 
    ' have already been received, even if no characters have been received.
    Public ReadIntervalTimeout As Integer

    ' Multiplier used to calculate the total time-out period for read operations, 
    ' in milliseconds. For each read operation, this value is multiplied by the 
    ' requested number of bytes to be read. 
    Public ReadTotalTimeoutMultiplier As Integer

    ' Constant used to calculate the total time-out period for read operations, 
    ' in milliseconds. For each read operation, this value is added to the 
    ' product of the ReadTotalTimeoutMultiplier member and the requested number 
    ' of bytes. 
    ' A value of zero for both the ReadTotalTimeoutMultiplier and 
    ' ReadTotalTimeoutConstant members indicates that total time-outs are not 
    ' used for read operations.
    Public ReadTotalTimeoutConstant As Integer

    ' Multiplier used to calculate the total time-out period for write operations, 
    ' in milliseconds. For each write operation, this value is multiplied by the 
    ' number of bytes to be written. 
    Public WriteTotalTimeoutMultiplier As Integer

    ' Constant used to calculate the total time-out period for write operations, 
    ' in milliseconds. For each write operation, this value is added to the 
    ' product of the WriteTotalTimeoutMultiplier member and the number of bytes 
    ' to be written. 
    ' A value of zero for both the WriteTotalTimeoutMultiplier and 
    ' WriteTotalTimeoutConstant members indicates that total time-outs are not 
    ' used for write operations.
    Public WriteTotalTimeoutConstant As Integer

  End Structure

  <StructLayout(LayoutKind.Sequential, Pack:=1)> Private Structure COMMCONFIG

    ' Size of the structure, in bytes. 
    Public Size As Integer

    ' Version number of the structure. This parameter can be 1. The version of 
    ' the provider-specific structure should be included in the wcProviderData 
    ' member. 
    Public Version As Int16

    ' Reserved; do not use. 
    Public Reserved As Int16

    ' Device-control block ( DCB) structure for RS-232 serial devices. A DCB 
    ' structure is always present regardless of the port driver subtype specified 
    ' in the device's COMMPROP structure. 
    Public DCB As DCB

    ' Type of communications provider, and thus the format of the provider-
    ' specific data. For a list of communications provider types, see the 
    ' description of the COMMPROP structure. 
    Public ProviderSubType As Integer

    ' Offset of the provider-specific data relative to the beginning of the 
    ' structure, in bytes. This member is zero if there is no provider-specific 
    ' data. 
    Public ProviderOffset As Integer

    ' Size of the provider-specific data, in bytes. 
    Public ProviderSize As Integer

    ' Optional provider-specific data. This member can be of any size or can
    ' be omitted. Because the COMMCONFIG structure may be expanded in the future, 
    ' applications should use the dwProviderOffset member to determine the 
    ' location of this member. 
    Public ProviderData As Byte

  End Structure

  <StructLayout(LayoutKind.Sequential, Pack:=1)> Public Structure OVERLAPPED

    ' Reserved for operating system use. This member, which specifies a 
    ' system-dependent status, is valid when the GetOverlappedResult function 
    ' returns without setting the extended error information to ERROR_IO_PENDING. 
    Public Internal As Integer

    ' Reserved for operating system use. This member, which specifies the length 
    ' of the data transferred, is valid when the GetOverlappedResult function 
    ' returns TRUE. 
    Public InternalHigh As Integer

    ' File position at which to start the transfer. The file position is a byte 
    ' offset from the start of the file. The calling process sets this member 
    ' before calling the ReadFile or WriteFile function. This member is ignored 
    ' when reading from or writing to named pipes and communications devices and 
    ' should be zero. 
    Public Offset As Integer

    ' High-order word of the byte offset at which to start the transfer. This 
    ' member is ignored when reading from or writing to named pipes and 
    ' communications devices and should be zero. 
    Public OffsetHigh As Integer

    ' Handle to an event set to the signaled state when the operation has been 
    ' completed. The calling process must set this member either to zero or a 
    ' valid event handle before calling any overlapped functions. To create an 
    ' event object, use the CreateEvent function. 
    ' Functions such as WriteFile set the event to the nonsignaled state before 
    ' they begin an I/O operation.
    Public [Event] As Integer

  End Structure

  <DllImport("kernel32.dll")> Private Shared Function FormatMessage(ByVal dwFlags As Integer, ByVal lpSource As Integer, ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, ByVal lpBuffer As StringBuilder, ByVal nSize As Integer, ByVal Arguments As Integer) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function BuildCommDCB(ByVal lpDef As String, ByRef lpDCB As DCB) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function ClearCommError(ByVal hFile As Integer, ByVal lpErrors As Integer, ByVal l As Integer) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function CloseHandle(ByVal hObject As Integer) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function CreateEvent(ByVal lpEventAttributes As Integer, ByVal bManualReset As Integer, ByVal bInitialState As Integer, <MarshalAs(UnmanagedType.LPStr)> ByVal lpName As String) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function CreateFile(<MarshalAs(UnmanagedType.LPStr)> ByVal lpFileName As String, ByVal dwDesiredAccess As Integer, ByVal dwShareMode As Integer, ByVal lpSecurityAttributes As Integer, ByVal dwCreationDisposition As Integer, ByVal dwFlagsAndAttributes As Integer, ByVal hTemplateFile As Integer) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function EscapeCommFunction(ByVal hFile As Integer, ByVal ifunc As Long) As Boolean
  End Function
  <DllImport("kernel32.dll")> Private Shared Function ClearCommError(ByVal hFile As Integer, ByRef lpErrors As Integer, ByRef lpStat As COMSTAT) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function FormatMessage(ByVal dwFlags As Integer, ByVal lpSource As Integer, ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, <MarshalAs(UnmanagedType.LPStr)> ByVal lpBuffer As String, ByVal nSize As Integer, ByVal Arguments As Integer) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function GetCommModemStatus(ByVal hFile As Integer, ByRef lpModemStatus As Integer) As Boolean
  End Function
  <DllImport("kernel32.dll")> Private Shared Function GetCommState(ByVal hCommDev As Integer, ByRef lpDCB As DCB) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function GetCommTimeouts(ByVal hFile As Integer, ByRef lpCommTimeouts As COMMTIMEOUTS) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function GetLastError() As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function GetOverlappedResult(ByVal hFile As Integer, ByRef lpOverlapped As Overlapped, ByRef lpNumberOfBytesTransferred As Integer, ByVal bWait As Integer) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function PurgeComm(ByVal hFile As Integer, ByVal dwFlags As Integer) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function ReadFile(ByVal hFile As Integer, ByVal Buffer As Byte(), ByVal nNumberOfBytesToRead As Integer, ByRef lpNumberOfBytesRead As Integer, ByRef lpOverlapped As Overlapped) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function SetCommTimeouts(ByVal hFile As Integer, ByRef lpCommTimeouts As COMMTIMEOUTS) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function SetCommState(ByVal hCommDev As Integer, ByRef lpDCB As DCB) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function SetupComm(ByVal hFile As Integer, ByVal dwInQueue As Integer, ByVal dwOutQueue As Integer) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function SetCommMask(ByVal hFile As Integer, ByVal lpEvtMask As Integer) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function WaitCommEvent(ByVal hFile As Integer, ByRef Mask As EventMasks, ByRef lpOverlap As Overlapped) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function WaitForSingleObject(ByVal hHandle As Integer, ByVal dwMilliseconds As Integer) As Integer
  End Function
  <DllImport("kernel32.dll")> Private Shared Function WriteFile(ByVal hFile As Integer, ByVal Buffer As Byte(), ByVal nNumberOfBytesToWrite As Integer, ByRef lpNumberOfBytesWritten As Integer, ByRef lpOverlapped As Overlapped) As Integer
  End Function

#End Region

#Region "Public Enums"

  Public Enum BaudRates
    Baud110 = 110
    Baud300 = 300
    Baud600 = 600
    Baud1200 = 1200
    Baud2400 = 2400
    Baud4800 = 4800
    Baud9600 = 9600
    Baud14400 = 14400
    Baud19200 = 19200
    Baud38400 = 38400
    Baud56000 = 56000
    Baud57600 = 56700
    Baud115200 = 115200
    Baud128000 = 128000
    Baud256000 = 256000
  End Enum

  Public Enum Parities
    None
    Odd
    Even
    Mark
    Space
  End Enum

  Public Enum DataSizes
    Size4 = 4
    Size5 = 5
    Size6 = 6
    Size7 = 7
    Size8 = 8
  End Enum

  Public Enum StopBits
    Bit1 = 1
    Bit2 = 2
    Bit1_5 '= 1.5
  End Enum

  Public Enum Handshakes
    None
    XOnXOff
    RTS
    RTSXOnXOff
  End Enum

  Public Enum InputModes
    Text = 0
    Binary = 1
  End Enum

#End Region

#Region "Private Enums"

  ' This enumeration contains values used to purge the various buffers.
  Private Enum PurgeBuffers
    RXAbort = &H2
    RXClear = &H8
    TxAbort = &H1
    TxClear = &H4
  End Enum

  ' This enumeration provides values for the lines sent to the Comm Port
  Private Enum Lines
    SetRts = 3
    ClearRts
    SetDtr
    ClearDtr
    ResetDev      '	 Reset device if possible
    SetBreak      '	 Set the device break line.
    ClearBreak    '	 Clear the device break line.
  End Enum

  ' This enumeration provides values for the Modem Status, since
  ' we'll be communicating primarily with a modem.
  ' Note that the Flags() attribute is set to allow for a bitwise
  ' combination of values.
  <Flags()> Private Enum ModemStatusBits
    ClearToSendOn = &H10
    DataSetReadyOn = &H20
    RingIndicatorOn = &H40
    CarrierDetect = &H80
  End Enum

  ' This enumeration provides values for the Working mode
  Private Enum Modes
    NonOverlapped
    Overlapped
  End Enum

  ' This enumeration provides values for the Comm Masks used.
  ' Note that the Flags() attribute is set to allow for a bitwise
  '   combination of values.
  <Flags()> Private Enum EventMasks
    RxChar = &H1
    RXFlag = &H2
    TxBufferEmpty = &H4
    ClearToSend = &H8
    DataSetReady = &H10
    ReceiveLine = &H20
    Break = &H40
    StatusError = &H80
    Ring = &H100
  End Enum

#End Region

#Region "Public Events"

  ' The OnComm event is generated whenever the value of the CommEvent property
  ' changes, indicating that either a communication event or an error occured.
  ' The CommEvent property containes the numeric code of the actual error or 
  ' event that generated the OnComm event.  Note that setting the RThreshold or
  ' SThreshold propety to 0 disables trapping for the EvReceive and EvSend events,
  ' respectively.
  Public Event OnComm(ByVal sender As Object, ByVal e As CommEventArgs)

  'Public Event DataReceived(ByVal Source As RS232, ByVal DataBuffer() As Byte)
  'Public Event TxCompleted(ByVal Source As RS232)
  'Public Event CommEvent(ByVal Source As RS232, ByVal Mask As EventMasks)

#End Region

#Region "Member variables"

  Dim _timer As Timers.Timer

  Dim _status As COMSTAT

  ' Settings.
  Private _break As Boolean = False
  Private _cdHolding As Boolean = False
  'Private _commEvent As CommEventsErrors
  Private _commID As Integer = -1 ' This value holds the CreateFile handle.
  Private _commPort As Integer = 1
  Private _ctsHolding As Boolean = False
  Private _dsrHolding As Boolean = False
  Private _dtrEnable As Boolean = False
  Private _eofEnable As Boolean = False
  Private _handshaking As Handshakes = Handshakes.None
  Private _inBufferCount As Integer = 0
  Private _inBufferSize As Integer = 1024
  Private _inputMode As InputModes = InputModes.Text
  Private _nullDiscard As Boolean = False
  Private _outBufferCount As Integer = 0
  Private _outBufferSize As Integer = 512
  Private _parityReplace As String = ""
  Private _portOpen As Boolean = False
  Private _receiveThreshold As Integer = 0
  Private _receiveThresholdReached As Boolean = False
  Private _rtsEnable As Boolean = False
  Private _sendThreshold As Integer

  ' Comm Settings
  Private _baudRate As BaudRates = BaudRates.Baud9600
  Private _parity As Parities = Parities.None
  Private _dataSize As DataSizes = DataSizes.Size8
  Private _stopBit As StopBits = StopBits.Bit1
  Private _inputLength As Integer = 512 ' 0 = read whole buffer

  ' right now, only supporting the non-overlapped mode.
  Private _mode As Modes = Modes.NonOverlapped ' Class working mode	

  Private _timeout As Integer = 70   ' Timeout in ms

  Private _receiveBuffer As Byte()

#End Region

#Region "Constructor/Dispose/Finalize"

  Public Sub New()
    ' implement any defaults here.
    _timer = New Timers.Timer(50)
    _timer.Enabled = False
    AddHandler _timer.Elapsed, AddressOf _timer_Elapsed
  End Sub

  Public Sub Dispose() Implements System.IDisposable.Dispose
    If _portOpen Then
      Close()
    End If
    GC.SuppressFinalize(Me)
  End Sub

  Protected Overrides Sub Finalize()
    If _portOpen Then
      Close()
    End If
    MyBase.Finalize()
  End Sub

#End Region

#Region "Helper functions"

  ' This function returns an integer specifying the number of bytes 
  '   read from the Comm Port. It accepts a parameter specifying the number
  '   of desired bytes to read.
  Private Function Read(ByVal bytes As Integer) As Integer

    Dim readChars, rc As Integer

    ' If Bytes2Read not specified uses Buffersize
    If bytes = 0 Then bytes = _inputLength
    If _commID = -1 Then
      Throw New ApplicationException( _
          "Please initialize and open port before using this method.")
    Else
      ' Get bytes from port
      Try
        ' Purge buffers
        ReDim _receiveBuffer(bytes - 1)
        rc = ReadFile(_commID, _receiveBuffer, bytes, readChars, Nothing)
        If rc = 0 Then
          ' Read Error
          Throw New ApplicationException("ReadFile error " & rc.ToString)
        Else
          ' Handles timeout or returns input chars
          If readChars < bytes Then
            Throw New IOTimeoutException("Timeout error")
          Else
            Return (readChars)
          End If
        End If
        'End If
      Catch Ex As Exception
        ' Others generic erroes
        Throw New ApplicationException("Read Error: " & Ex.Message, Ex)
      End Try
    End If
  End Function

  ' This subroutine writes the passed array of bytes to the 
  ' Comm Port to be written.
  Private Overloads Sub Write(ByVal buffer As Byte())

    Dim iBytesWritten, iRc As Integer

    If _commID = -1 Then
      Throw New ApplicationException( _
          "Please initialize and open port before using this method")
    Else
      ' Transmit data to COM Port
      Try
        ' Clears IO buffers
        PurgeComm(_commID, PURGE_RXCLEAR Or PURGE_TXCLEAR)
        iRc = WriteFile(_commID, buffer, buffer.Length, _
            iBytesWritten, Nothing)
        If iRc = 0 Then
          Throw New ApplicationException( _
              "Write Error - Bytes Written " & _
              iBytesWritten.ToString & " of " & _
              buffer.Length.ToString)
        End If
      Catch Ex As Exception
        Throw
      End Try
    End If
  End Sub

  Private Sub GetState()

    If _portOpen Then

      Dim errors As Integer
      Dim result As Integer = ClearCommError(_commID, errors, _status)

      If result = 0 Then
        'TODO: Set CommEvent property.
        ' error - use getlasterror to get error code.
        'Return 0
      Else
        If errors <> 0 Then
          ' The hardware detected a break condition.
          If (errors And CE_BREAK) = CE_BREAK Then
            RaiseEvent OnComm(Me, New CommEventArgs(CommEventArgs.CommEvents.Break))
          End If
          ' The hardware detected a framing error.
          If (errors And CE_FRAME) = CE_FRAME Then
            RaiseEvent OnComm(Me, New CommEventArgs(CommEventArgs.CommEvents.Frame))
          End If
          ' An I/O error occurred during communications with the device.
          If (errors And CE_IOE) = CE_IOE Then
            RaiseEvent OnComm(Me, New CommEventArgs(CommEventArgs.CommEvents.IOE))
          End If
          ' The requested mode is not supported, or the hFile parameter is invalid. If this value is specified, it is the only valid error.
          If (errors And CE_MODE) = CE_MODE Then
            RaiseEvent OnComm(Me, New CommEventArgs(CommEventArgs.CommEvents.Mode))
          End If
          ' A character-buffer overrun has occurred. The next character is lost.
          If (errors And CE_OVERRUN) = CE_OVERRUN Then
            RaiseEvent OnComm(Me, New CommEventArgs(CommEventArgs.CommEvents.Overrun))
          End If
          ' An input buffer overflow has occurred. There is either no room in the input buffer, or a character was received after the end-of-file (EOF) character.
          If (errors And CE_RXOVER) = CE_RXOVER Then
            RaiseEvent OnComm(Me, New CommEventArgs(CommEventArgs.CommEvents.RxOver))
          End If
          ' The hardware detected a parity error.
          If (errors And CE_RXPARITY) = CE_RXPARITY Then
            RaiseEvent OnComm(Me, New CommEventArgs(CommEventArgs.CommEvents.RxParity))
          End If
          ' The application tried to transmit a character, but the output buffer was full.
          If (errors And CE_TXFULL) = CE_TXFULL Then
            RaiseEvent OnComm(Me, New CommEventArgs(CommEventArgs.CommEvents.TxFull))
          End If
        End If
        ' success
        _ctsHolding = ((_status.Flags And COMSTATFlags.CTSHold) = COMSTATFlags.CTSHold)
        _dsrHolding = ((_status.Flags And COMSTATFlags.DSRHold) = COMSTATFlags.DSRHold)
        '_eof = ((_status.Flags And COMSTATFlags.EOF) = COMSTATFlags.EOF)
        _inBufferCount = _status.InQue
        If _receiveThreshold > 0 Then
          If _receiveThresholdReached Then
            ' check to see if we've dropped below the threshold
            If _inBufferCount < _receiveThreshold Then
              _receiveThresholdReached = False
            End If
          Else
            If _inBufferCount >= _receiveThreshold Then
              _receiveThresholdReached = True
              RaiseEvent OnComm(Me, New CommEventArgs(CommEventArgs.CommEvents.Receive, _inBufferCount))
            End If
          End If
        End If
        _outBufferCount = _status.OutQue
      End If

    Else
      _ctsHolding = False
      _dsrHolding = False
      _inBufferCount = 0
      _outBufferCount = 0
    End If

  End Sub

  ' This subroutine opens and initializes the Comm Port
  Private Overloads Sub Open()

    Dim dcb As DCB
    Dim result As Integer
    Dim errorCode As Integer

    ' Get Dcb block,Update with current data

    ' Set working mode
    Dim mode As Integer = CInt(IIf(_mode = Modes.Overlapped, FILE_FLAG_OVERLAPPED, 0))

    ' Initializes Com Port
    Try

      ' Creates a COM Port stream handle 
      _commID = CreateFile("COM" & _commPort.ToString, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, mode, 0)

      If _commID <> -1 Then

        ' Clear all comunication errors
        result = ClearCommError(_commID, errorCode, 0&)

        ' Clears I/O buffers
        result = PurgeComm(_commID, PurgeBuffers.RXClear Or PurgeBuffers.TxClear)

        ' Gets COM Settings
        result = GetCommState(_commID, dcb)

        ' Updates COM Settings
        Dim parity As String = "NOEMS"
        parity = parity.Substring(_parity, 1)

        ' Set DCB State
        Dim dcbState As String = String.Format("baud={0} parity={1} data={2} stop={3}", _
            CInt(_baudRate).ToString, parity, CInt(_dataSize), CInt(_stopBit))

        result = BuildCommDCB(dcbState, dcb)
        result = SetCommState(_commID, dcb)

        If result = 0 Then
          Dim message As String = Error2Text(GetLastError())
          Throw New CIOChannelException("Unable to set COM state0 " & message)
        End If

        ' Setup Buffers (Rx,Tx)
        result = SetupComm(_commID, _inBufferSize, _outBufferSize)

        ' Set Timeouts
        SetTimeout()

        _portOpen = True

      Else
        ' Raise Initialization problems
        Throw New CIOChannelException("Unable to open COM" & _commPort.ToString)
      End If

    Catch Ex As Exception
      ' Generic error
      Throw New CIOChannelException(Ex.Message, Ex)

    End Try


  End Sub

  ' This subroutine sets the Comm Port timeouts.
  Private Sub SetTimeout()

    Dim ctm As COMMTIMEOUTS

    ' Set ComTimeout
    If _commID = -1 Then
      Exit Sub
    Else
      ' Changes setup on the fly
      With ctm
        .ReadIntervalTimeout = 0
        .ReadTotalTimeoutMultiplier = 0
        .ReadTotalTimeoutConstant = _timeout
        .WriteTotalTimeoutMultiplier = 10
        .WriteTotalTimeoutConstant = 100
      End With
      SetCommTimeouts(_commID, ctm)
    End If

  End Sub

  ' This subroutine closes the Comm Port.
  Private Sub Close()
    If _portOpen AndAlso _commID <> -1 Then
      CloseHandle(_commID)
      _commID = -1
      _portOpen = False
    End If
  End Sub

  ' This function translates an API error code to text.
  Private Function Error2Text(ByVal lCode As Integer) As String

    Dim rc As New StringBuilder(256)
    Dim ret As Integer

    ret = FormatMessage(&H1000, 0, lCode, 0, rc, 256, 0)
    If ret > 0 Then
      Return rc.ToString
    Else
      Return "Error not found."
    End If

  End Function

#End Region

#Region "Public Properties and Methods"

  ' Sets or clears the break signal state.
  ' When set to True, the Break property sends a break signal.
  ' The break signal suspends character transmission and places
  ' the transmission line in a break state until you set the
  ' Break property to False.
  ' Typically, you set the break state for a short interval, and 
  ' *only* if the device with which you are communicating requires
  ' that a break signal be set.
  Public Property Break() As Boolean
    Get
      ' Just returns the current break state.
      Return _break
    End Get
    Set(ByVal Value As Boolean)
      _break = Value
      'TODO: suspend or resume transmission based on break state.
    End Set
  End Property

  ' Determines whether the carrier is present by querying the state of the
  ' Carrier Detect (CD) line.  Carrier Detect is a signal sent from a modem
  ' to the attached computer to indicate that the modem is online.
  Public ReadOnly Property CDHolding() As Boolean
    Get
      ' True = High
      ' False = Low
      'TODO: figure out how to get CD high/low state.
      Return _cdHolding
    End Get
  End Property

  '' Returns the most recent communications event or error.
  '' Not used... replaced by the e parameter passed with the event.
  'Public ReadOnly Property CommEvent() As CommEventsErrors
  '  Get
  '    'TODO: be sure to set any comm event errors/etc. so that when
  '    ' reading this value, it will work.
  '    Return _commEvent
  '  End Get
  'End Property

  ' Returns the handle that identifies the communications device.
  ' This is the same value that's returned by the Windows API 
  ' CreateFile function.  Use this value when calling any
  ' communications routines in the Windows API.
  Public ReadOnly Property CommID() As Integer
    Get
      Return _commID
    End Get
  End Property

  ' Sets or returns the communications port number.
  ' You can set value to any number between 1 and 16; defaults to 1.
  ' However, the component control generates error *Device Unavailable*
  ' if the port does not exist when you attempt to open it with the 
  ' PortOpen property.
  Public Property CommPort() As Integer
    Get
      Return _commPort
    End Get
    Set(ByVal Value As Integer)
      If Value > 0 AndAlso Value < 17 Then
        _commPort = Value
      Else
        Throw New ApplicationException("CommPort value must be between 1 and 16.")
      End If
    End Set
  End Property

  ' Determines whether you can send data by querying the state of the
  ' Clear To Send (CTS) line.  Typically, the Clear To Send signal is
  ' sent from a modem to the attached computer to indicate that transmission
  ' can proceed.
  ' When the CTS is low (CTSHolding = False) and times out, the component
  ' control sets the CommEvent property to EventCTSTO (CTS Timeout) and invokes
  ' the OnComm event.
  ' The CTS line is used in RTS/CTS (Request ToSend/Clear To Send) hardware
  ' handshaking.  The CTSHolding property gives you a way to manually poll the
  ' Clear To Send line if you need to determine it's state.
  Public ReadOnly Property CTSHolding() As Boolean
    Get
      Return _ctsHolding
    End Get
  End Property

  ' Determines the state of the Data Set Ready (DSR) line.  Typically, the
  ' DSR signal is sent by a modem to it's attached computer to indicate that
  ' it is ready to operate.
  ' When the DSR line is high (DSRHolding = True) and has timed out, the
  ' component control sets the CommEvent property to EventDSRTO (DSR Timeout)
  ' and invokes the OnComm event.
  ' This property is useful when writing Data Set Ready/Data Terminal Ready
  ' handshaking routine for a Data Terminal Equipment (DTE) machine.
  Public ReadOnly Property DSRHolding() As Boolean
    Get
      Return _dsrHolding
    End Get
  End Property

  ' Determines whether to enable the Data Terminal Ready (DTR) line during
  ' communications.  Typically, the DTR signal is sent by a computer to it's 
  ' modem to indicate that the computer is ready to accept incoming transmission.
  Public Property DTREnable() As Boolean
    Get
      Return _dtrEnable
    End Get
    Set(ByVal Value As Boolean)
      ' When DTREnable is set to True, The DTR line is set to high (on) when
      ' the port is opened, and low (off) when the port is closed.  When DTREnable
      ' is set to False, the DTR always remains low.
      ' Note: In most cases, setting the DTR line to low hangs up the telephone.
      If _portOpen Then
        If Value Then
          EscapeCommFunction(_commID, Lines.SetDtr)
        Else
          EscapeCommFunction(_commID, Lines.ClearDtr)
        End If
        _dtrEnable = Value
      Else
        Throw New ApplicationException("You must call PortOpen before using this method.")
      End If
    End Set
  End Property

  ' The EOFEnable property determines if the control looks for the End Of File
  ' (EOF) characters during input.  If an EOF character is found, the input will
  ' stop and the OnComm event will fire with the CommEvent property set to EvEOF.
  Public Property EOFEnable() As Boolean
    Get
      Return _eofEnable
    End Get
    Set(ByVal Value As Boolean)
      _eofEnable = Value
    End Set
  End Property

  ' Sets or returns the hardware handshaking protocol.
  ' Handshaking refers to the internal communications protocol by which data
  ' is transferred from the hardware port to the receive buffer.  When a character
  ' of data arrives at the serial port, the communications device has to move 
  ' it into the receive buffer so that your program can read it.  If there is no 
  ' receive buffer and your program is expected to read every character directly 
  ' from the hardware, you will probably lose data because the characters can 
  ' arrive very quickly.
  ' A handshaking protocol insures data is not lost due to a buffer overrun, where
  ' data arives at the port too quickly for the communications device to move the
  ' data into the receive buffer.
  Public Property Handshaking() As Handshakes
    Get
      Return _handshaking
    End Get
    Set(ByVal Value As Handshakes)
      _handshaking = Value
    End Set
  End Property

  ' Returns the number of characters waiting in the receive buffer.
  ' InBufferCount refers to the number of characters that have been received 
  ' by the modem and are waiting in the receive buffer for you to take them 
  ' out.  You can clear the receive buffer by setting the InBufferCount 
  ' property to 0.
  Public Property InBufferCount() As Integer
    Get
      GetState()
      Return _inBufferCount
    End Get
    Set(ByVal Value As Integer)
      If Value = 0 Then
        'Clear receive buffer.
        If _portOpen Then
          PurgeComm(_commID, PURGE_RXCLEAR)
          GetState()
        End If
      Else
        Throw New ApplicationException("You can only set InBufferCount to 0 (reset).")
      End If
    End Set
  End Property

  ' Sets and returns the size of the receive buffer in bytes.
  ' InBufferSize referes to the total size of the receive buffer.  The default
  ' size is 1024 bytes.  
  Public Property InBufferSize() As Integer
    Get
      Return _inBufferSize
    End Get
    Set(ByVal Value As Integer)
      _inBufferSize = Value
    End Set
  End Property

  ' Returns and removes a stream of data from the receive buffer.
  ' The InputLength property determines the number of characters that are read
  ' by the Input property.  Setting InputLength to 0 causes the Input property
  ' to read the entire contents of the receive buffer.
  ' The InputMode property determines the type of data that is retrieved with the
  ' Input property.  If InputMode is set to Text then the Input property returns
  ' string data.  If InputMode is Binary then the Input property returns binary
  ' data in a array of bytes.
  Public ReadOnly Property Input() As String
    Get
      ' to see how many characters are in the buffer.
      GetState()
      If _inBufferCount > 0 Then
        If _inputLength = 0 Then
          ' read all of the characters in the buffer.
          Me.Read(_inBufferCount)
        Else
          ' only return data if there are at least the
          ' number of bytes requested.
          If _inputLength <= _inBufferCount Then
            Me.Read(_inputLength)
          Else
            ' otherwise, return a zero-length string.
            Return ""
          End If
        End If

        Dim encode As New System.Text.ASCIIEncoding()
        Return encode.GetString(_receiveBuffer)

      Else
        ' there's nothing in the buffer, so return a zero-length string.
        Return ""
      End If
    End Get
  End Property

  Public ReadOnly Property InputBytes() As Byte()
    Get

      Dim noda() As Byte ' return an empty byte array if criteria isn't met.

      ' to see how many characters are in the buffer.
      GetState()
      If _inBufferCount > 0 Then
        If _inputLength = 0 Then
          ' read all of the characters in the buffer.
          Me.Read(_inBufferCount)
        Else
          ' only return data if there are at least the
          ' number of bytes requested.
          If _inputLength <= _inBufferCount Then
            Me.Read(_inputLength)
          Else
            ' otherwise, return a zero-length string.
            Return noda
          End If
        End If

        Return _receiveBuffer

      Else
        ' there's nothing in the buffer, so return a zero-length string.
        Return noda
      End If
    End Get

  End Property

  ' Sets or returns the number of characters the Input property 
  ' reads from the receive buffer.
  ' The default value for the InputLength property is 0.  Setting InputLength
  ' to 0 causes the control to read the entire contents of the receive buffer
  ' when Input is used.
  ' If InputLength characters are not available in the receive buffer, the Input
  ' property returns a zero-length string ("").  The user can optionally check
  ' the InBufferCount property to determine if the required number of characters
  ' are present before using Input.
  ' This property is useful when reading data from a machine whose output
  ' is formatted in fixed-length blocks of data.
  Public Property InputLength() As Integer
    Get
      Return _inputLength
    End Get
    Set(ByVal Value As Integer)
      _inputLength = Value
    End Set
  End Property

  ' Sets or returns the type of data retrieved by the Input property.
  ' The InputMode property determines how data will be retrieved through
  ' the Input property.  The data will either be retrieved as string or
  ' as binary data in a byte array.
  ' Use Text for data that uses ANSI character set.  Use Binary for all
  ' other data such as data that has embedded control characters, Nulls, etc.
  Public Property InputMode() As InputModes
    Get
      Return _inputMode
    End Get
    Set(ByVal Value As InputModes)
      _inputMode = Value
    End Set
  End Property

  '' This read-only property returns the status of the modem.
  'Public ReadOnly Property ModemStatus() As ModemStatusBits
  '  Get
  '    If _CommID = -1 Then
  '      Throw New ApplicationException("Please initialize and open " + _
  '          "port before using this method")
  '    Else
  '      ' Retrieve modem status
  '      Dim lpModemStatus As Integer
  '      If Not GetCommModemStatus(_CommID, lpModemStatus) Then
  '        Throw New ApplicationException("Unable to get modem status")
  '      Else
  '        Return CType(lpModemStatus, ModemStatusBits)
  '      End If
  '    End If
  '  End Get
  'End Property

  ' Determines whether null characters are transferred from the port 
  ' to the receive buffer.
  ' The null character is defined as ASCII character 0, Chr(0).
  Public Property NullDiscard() As Boolean
    Get
      Return _nullDiscard
    End Get
    Set(ByVal Value As Boolean)
      _nullDiscard = Value
    End Set
  End Property

  ' Returns the number of characters waiting in the trasmit buffer.
  ' You can clear the transmit buffer by setting the OutBufferCount 
  ' property to 0.
  Public Property OutBufferCount() As Integer
    Get
      'TODO: Get number of characters in the buffer.
      Return _outBufferCount
    End Get
    Set(ByVal Value As Integer)
      If Value = 0 Then
        'TODO: clear the output buffer.
        _outBufferCount = 0
      Else
        Throw New ApplicationException("You can only set OutBufferCount to 0 (reset).")
      End If
    End Set
  End Property

  ' Sets and returns the size, in bytes, of the transmit buffer.
  ' OutBufferSize refers to the total size of the transmit buffer.
  ' The default size is 512 bytes.
  Public Property OutBufferSize() As Integer
    Get
      Return _outBufferSize
    End Get
    Set(ByVal Value As Integer)
      _outBufferSize = Value
    End Set
  End Property

  ' Writes a stream of data to the transmit buffer.
  ' The Output property can transmit text data or binary data.
  ' To send text data using the Output property, you must specify
  ' a string value.  To send binary data, you must pass the value ' 
  ' in as an array of bytes.
  Public Overloads Sub Output(ByVal buffer As String)
    If _portOpen Then
      Dim encode As New System.Text.ASCIIEncoding()
      Dim bytes() As Byte = encode.GetBytes(buffer)
      Write(bytes)
    Else
      Throw New ApplicationException("You must call PortOpen before using this method.")
    End If
  End Sub

  Public Overloads Sub Output(ByVal buffer() As Byte)
    If _portOpen Then
      Write(buffer)
    Else
      Throw New ApplicationException("You must call PortOpen before using this method.")
    End If
  End Sub

  ' Sets and returns the character that replaces an invalid character
  ' in the data stream whan a parity error occurs.
  ' The parity bit refers to a bit that is transmitted along with a 
  ' specified number of data bits to provide a small amount of error 
  ' checking.  When you use a parity bit, the component adds up all the 
  ' bits that are set (having a value of 1) in the data and tests the sum 
  ' as bing odd or even (according to the parity setting used when the port 
  ' was opened).
  ' By default, the control uses a question mark (?) character for replacing
  ' invalid characters.  Setting ParityReplace to an empty string ("") disables
  ' replacement of the character where the parity error occurs.  The OnComm
  ' event is still fired and the CommEvent property is set to EventRXParity.
  ' ParityReplace character is used in a byte-oriented operation, and must be 
  ' a single-byte character.  You can specify an ANSI character code with a 
  ' value from 0 to 255
  Public Property ParityReplace() As String
    Get
      Return _parityReplace
    End Get
    Set(ByVal Value As String)
      _parityReplace = Value
    End Set
  End Property

  ' Sets and returns the state of the communications port (open or closed).
  ' Setting the PortOpen property to True opens the port. Setting it to False
  ' closes the port and clears the receive and transmit buffers.  The component
  ' automatically closes the serial port when your application is terminated.
  ' Make sure the CommPort property is set to a valid port number before opening
  ' the port.  If you the CommPort property is set to an invalid port number when
  ' you trie to open the port, the control generates error (Device unavailable).
  ' In addition, your serial port device must support the current values in the
  ' Settings property.  If the Settings property contains communications settings
  ' that your hardware does not support, your hardware may not work correctly.
  ' If either the DTREnable or the RTSEnable properties is set to True before the
  ' port is opened, the properties are set to False when the port is closed.
  ' Otherwise, the DTR and RTS lines remain in their previous state.
  Public Property PortOpen() As Boolean
    Get
      Return _portOpen
    End Get
    Set(ByVal Value As Boolean)
      If Value AndAlso Not _portOpen Then
        Open()
        _timer.Enabled = True
      ElseIf Not Value AndAlso _portOpen Then
        _timer.Enabled = False
        Close()
        _dtrEnable = False
        _rtsEnable = False
      Else
        If _portOpen Then
          Throw New ApplicationException("Port already open.")
        Else
          Throw New ApplicationException("Port already closed.")
        End If
      End If
    End Set
  End Property

  ' Sets and returns the number of characters to receive before the control 
  ' sets the CommEvent property to EvReceive and generates the OnComm event.
  ' Setting the ReceiveThreshold property to 0 (the default) disables 
  ' generating the OnComm event when characters are received.
  ' Setting ReceiveThreshold to 1, for example, causes the control to generate
  ' the OnComm event every time a single character is placed in the receive buffer.
  Public Property ReceiveThreshold() As Integer
    Get
      Return _receiveThreshold
    End Get
    Set(ByVal Value As Integer)
      _receiveThreshold = Value
    End Set
  End Property

  ' Determines whether to enable the Request To Send (RTS) line.  Typically,
  ' the RTS signal that requests permission to transmit data is sent from a 
  ' computer to it's attached modem.
  ' When RTSEnable is set to True, the RTS line is set to high (on) when the
  ' port is opened, and low (off) when the port is closed.
  ' The RTS line is used in RTS/CTS hardware handshaking.  The RTSEnable
  ' property allows you to manually poll the RTS line if you need to determine
  ' its state.
  Public Property RTSEnable() As Boolean
    Get
      Return _rtsEnable
    End Get
    Set(ByVal Value As Boolean)
      If _portOpen Then
        If Value Then
          EscapeCommFunction(_commID, Lines.SetRts)
        Else
          EscapeCommFunction(_commID, Lines.ClearRts)
        End If
        _rtsEnable = Value
      Else
        Throw New ApplicationException("You must call PortOpen before using this method.")
      End If
    End Set
  End Property

  ' Sets and returns the baude rate, parity, data bit, and stop bit parameters.
  ' Must be in the format "baud, parity, data size, stop bits"... for example,
  ' "9600, N, 8, 1"
  Public Property Settings() As String
    Get
      Dim result As String
      result = CInt(_baudRate).ToString & ", "
      Select Case _parity
        Case Parities.Even : result &= "E, "
        Case Parities.Mark : result &= "M, "
        Case Parities.None : result &= "N, "
        Case Parities.Odd : result &= "O, "
        Case Parities.Space : result &= "S, "
        Case Else : result &= "?, "
      End Select
      result &= CInt(_dataSize).ToString & ", "
      result &= CInt(_stopBit).ToString
      Return result
    End Get
    Set(ByVal Value As String)
      ' using the value passed in, split it up and set appropriate values.
      If Value.Length > 8 Then
        Dim values() As String = Split(Replace(Value, Chr(32), ""), ",")
        If values.Length = 4 Then ' 4 elements.

          If IsNumeric(values(0)) Then
            Select Case CInt(values(0))
              Case 110, 300, 600, 1200, 2400, 4800, 9600, 14400, 19200, 28800, 38400, 56000, 128000, 256000
                _baudRate = CType(CInt(values(0)), BaudRates)

                If Not IsNumeric(values(1)) AndAlso values(1).Length = 1 Then
                  Select Case values(1).ToUpper
                    Case "E" : _parity = Parities.Even
                    Case "M" : _parity = Parities.Mark
                    Case "N" : _parity = Parities.None
                    Case "O" : _parity = Parities.Odd
                    Case "S" : _parity = Parities.Space
                    Case Else
                      Throw New ApplicationException("Invalid parity value specified.")
                  End Select

                  If IsNumeric(values(2)) AndAlso values(2).Length = 1 Then
                    If CInt(values(2)) > 3 AndAlso CInt(values(2)) < 9 Then
                      _dataSize = CType(CInt(values(2)), DataSizes)

                      If IsNumeric(values(3)) Then
                        Select Case CSng(values(3))
                          Case 1, 2 ', 1.5 
                            _stopBit = CType(CInt(values(3)), StopBits)
                          Case Else
                            Throw New ApplicationException("Invalid stopbit value specified.")
                        End Select
                      End If

                    Else
                      Throw New ApplicationException("Invalid datasize value specified.")
                    End If
                  End If

                Else
                  Throw New ApplicationException("Invalid parity value specified.")
                End If

              Case Else
                Throw New ApplicationException("Invalid baudrate value specified.")
            End Select
          End If
        Else
          Throw New ApplicationException("The parameter is formatted incorrectly.")
        End If
      Else
        Throw New ApplicationException("The parameter is formatted incorrectly.")
      End If
    End Set
  End Property

  Public Property BaudRate() As BaudRates
    Get
      Return _baudRate
    End Get
    Set(ByVal Value As BaudRates)
      _baudRate = Value
    End Set
  End Property

  Public Property Parity() As Parities
    Get
      Return _parity
    End Get
    Set(ByVal Value As Parities)
      _parity = Value
    End Set
  End Property

  Public Property DataSize() As DataSizes
    Get
      Return _dataSize
    End Get
    Set(ByVal Value As DataSizes)
      _dataSize = Value
    End Set
  End Property

  Public Property StopBit() As StopBits
    Get
      Return _stopBit
    End Get
    Set(ByVal Value As StopBits)
      _stopBit = Value
    End Set
  End Property

  ' Sets and returns the minimum number of characters allowable in the transmit
  ' buffer before the control sets the CommEvent property to EvSend and generates
  ' the OnComm event.
  ' Setting the SendThreshold property to 0 (the default) disables generating the 
  ' OnComm event for data transmission events.  Setting the SendThreshold property
  ' to 1 causes the control to generate the OnComm event when the transmission 
  ' buffer is completely empty.
  ' If the number of characters in the transmit buffer is less than value, the 
  ' CommEvent property is set to EvSend, and the OnComm event is generated.  The 
  ' EvSend even is only fired once, when the number of characters crosses the 
  ' SendThreshold.  For example, if SendThreshold equals five, the EvSend event 
  ' occurs only when the number of characters drops from five to four in the
  ' output queue.  If there are never more than SendThreshold characters in 
  ' the output queue, the even is never fired.
  Public Property SendThreshold() As Integer
    Get
      Return _sendThreshold
    End Get
    Set(ByVal Value As Integer)
      _sendThreshold = Value
    End Set
  End Property

  Public Overridable Property Timeout() As Integer
    Get
      Return _timeout
    End Get
    Set(ByVal Value As Integer)
      _timeout = CInt(IIf(Value = 0, 500, Value))
      ' If Port is open updates it on the fly
      SetTimeout()
    End Set
  End Property

  '' This property gets or sets the working mode to overlapped or non-overlapped.
  'Public Property WorkingMode() As Mode
  '  Get
  '    Return meMode
  '  End Get
  '  Set(ByVal Value As Mode)
  '    meMode = Value
  '  End Set
  'End Property

  ' This function takes the ModemStatusBits and returns a boolean value
  '   signifying whether the Modem is active.
  'Private Function CheckLineStatus(ByVal Line As ModemStatusBits) As Boolean
  '  Return Convert.ToBoolean(ModemStatus And Line)
  'End Function

#End Region

#Region "Custom Exceptions"

  ' This class defines a customized channel exception. This exception is
  '   raised when a NACK is raised.
  Public Class CIOChannelException : Inherits ApplicationException
    Sub New(ByVal Message As String)
      MyBase.New(Message)
    End Sub
    Sub New(ByVal Message As String, ByVal InnerException As Exception)
      MyBase.New(Message, InnerException)
    End Sub
  End Class

  ' This class defines a customized timeout exception.
  Public Class IOTimeoutException : Inherits CIOChannelException
    Sub New(ByVal Message As String)
      MyBase.New(Message)
    End Sub
    Sub New(ByVal Message As String, ByVal InnerException As Exception)
      MyBase.New(Message, InnerException)
    End Sub
  End Class

#End Region

  Private Sub _timer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs)
    GetState()
  End Sub

End Class

Public Class CommEventArgs
  Inherits EventArgs

  Public Enum CommEvents
    None = 0
    Send = 1
    Receive
    CTS
    DSR
    CD
    Ring
    EOF
    Break = 1001
    CTSTO
    DSRTO
    Frame
    Overrun
    CDTO
    RxOver
    TxFull
    DCB
    IOE
    Mode
    RxParity
  End Enum

  Private _event As CommEvents = CommEvents.None
  Private _length As Integer = 0

  Public Sub New(ByVal [event] As CommEvents)
    _event = [event]
  End Sub

  Public Sub New(ByVal [event] As CommEvents, ByVal length As Integer)
    _event = [event]
    _length = length
  End Sub

  Public ReadOnly Property [Event]() As CommEvents
    Get
      Return _event
    End Get
  End Property

  Public ReadOnly Property Length() As Integer
    Get
      Return _length
    End Get
  End Property

End Class
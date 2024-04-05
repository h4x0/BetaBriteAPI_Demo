Imports BetaBrite.RS232
Imports BetaBrite.Protocol

''' <summary>
''' Allows programmatic control of most BetaBrite sign features
''' </summary>
''' <remarks>
'''   Jeff Atwood
'''   http://www.codinghorror.com/
''' </remarks>
Public Class Sign

    Private _RS232 As BetaBrite.RS232
    Private _Address As String
    Private _IsDebug As Boolean = Diagnostics.Debugger.IsAttached
    Private _MemoryCommands As New SetMemoryCommands
    Private _TextFileCommands As New WriteTextFileCommands

    ''' <summary>
    ''' Sets the address and communication port of the BetaBrite sign you wish to control
    ''' communication will not be opened until the first command is sent
    ''' </summary>
    Public Sub New(Optional ByVal comPort As Integer = 1, Optional ByVal address As String = "00")
        _RS232 = New BetaBrite.RS232
        _RS232.CommPort = comPort
        _Address = address
    End Sub

    ''' <summary>
    ''' In debug mode, all commands are 'pretty printed' to Debug.Trace 
    ''' when they are sent to the sign
    ''' </summary>
    Public Property DebugMode() As Boolean
        Get
            Return _IsDebug
        End Get
        Set(ByVal Value As Boolean)
            _IsDebug = Value
        End Set
    End Property

    ''' <summary>
    ''' Returns true if the sign is Open and ready to accept commands.
    ''' note that the sign will be opened automatically when the 
    ''' first command is sent
    ''' </summary>
    Public ReadOnly Property IsOpen() As Boolean
        Get
            If _RS232 Is Nothing Then
                Return False
            Else
                Return _RS232.PortOpen
            End If
        End Get
    End Property

    ''' <summary>
    ''' Queue a request to use memory for a text element in this file label
    ''' Call AllocateMemory to perform your queued allocations
    ''' </summary>
    Public Sub UseMemoryText(ByVal fileLabel As Char, ByVal sizeBytes As Integer)
        _MemoryCommands.AllocateTextFile(fileLabel, Protection.Locked, sizeBytes)
    End Sub

    ''' <summary>
    ''' Queue a request to use memory for a string element in this file label
    ''' Call AllocateMemory to perform your queued allocations
    ''' </summary>
    Public Sub UseMemoryString(ByVal fileLabel As Char, ByVal sizeBytes As Integer)
        _MemoryCommands.AllocateStringFile(fileLabel, sizeBytes)
    End Sub

    ''' <summary>
    ''' Queue a request to use memory for a picture of a default (80x7) size in this file label
    ''' Call AllocateMemory to perform your queued allocations
    ''' </summary>
    Public Sub UseMemoryPicture(ByVal fileLabel As Char)
        _MemoryCommands.AllocatePictureFile(fileLabel)
    End Sub

    ''' <summary>
    ''' Queue a request to use memory for a picture of a specific size in this file label
    ''' Call AllocateMemory to perform your queued allocations
    ''' </summary>
    Public Sub UseMemoryPicture(ByVal fileLabel As Char, ByVal width As Integer, ByVal height As Integer)
        _MemoryCommands.AllocatePictureFile(fileLabel, Protection.Locked, width, height)
    End Sub

    ''' <summary>
    ''' Allocates all queued memory requests in the sign's memory. This is always destructive!
    ''' </summary>
    Public Sub AllocateMemory()
        If _MemoryCommands.Count = 0 Then
            Throw New Exception("No memory to allocate; use the 'UseMemory' commands to specify what type of memory you need first.")
        End If
        SendCommand(_MemoryCommands)
        '-- clear for next call
        _MemoryCommands = New SetMemoryCommands
    End Sub

    ''' <summary>
    ''' Calculates the exact amount of memory storage required for
    ''' a fully expanded message with control codes and/or international characters
    ''' </summary>
    Public Function CalculateMessageLength(ByVal message As String) As Integer
        Return ExpandedMessageLength(message)
    End Function

    ''' <summary>
    ''' Sets the date and time on the sign to the current system date/time
    ''' </summary>
    Public Sub SetDateAndTime()
        SetDateAndTime(DateTime.Now)
    End Sub

    ''' <summary>
    ''' Sets the date and time on the sign to any arbitrary date/time
    ''' </summary>
    Public Sub SetDateAndTime(ByVal dt As DateTime)
        Dim dc As New BetaBrite.Protocol.SetDateTimeCommand(dt)
        If _IsDebug Then
            Debug.Write(dc.ToString)
        End If
        Write(dc.ToBytes)
    End Sub

    ''' <summary>
    ''' Sets a run sequence 1 to 128 file labels (note: text files only)
    ''' eg, "DEBC" would display text files D, E, B, and C.
    ''' </summary>
    Sub SetRunSequence(ByVal fileLabels As String)
        Dim rc As New SetRunSequenceCommand
        For Each c As Char In fileLabels
            rc.AddFile(c)
        Next
        SendCommand(rc)
    End Sub

    ''' <summary>
    ''' Displays a single message on the sign and holds it there.
    ''' This basic command not require allocating memory, but can only display one message in file label "A".
    ''' HTML-style formatting codes can be used to specify various display options. 
    ''' </summary>
    Public Sub Display(ByVal message As String)
        '-- allocate memory for the single message in the first File Label "A"
        Dim mc As New SetMemoryCommand("A"c, FileType.Text, Protection.Unlocked, message)
        Console.WriteLine(mc.ToString)
        Write(mc.ToBytes)

        '-- fill the first file label "A" with our message
        Dim tc As New WriteTextFileCommand("A"c, message, Transition.Hold)
        SendCommand(tc)
    End Sub

    ''' <summary>
    ''' Sets a single text message in the specified file label.
    ''' Once set, a particular file label can be displayed by setting the RunOrder sequence.
    ''' HTML-style markup can be used to specify various display and visual options within the message.
    ''' </summary>
    Public Sub SetText(ByVal fileLabel As Char, ByVal message As String, _
            Optional ByVal t As Transition = Transition.Auto, Optional ByVal sm As Special = Special.None)
        SendCommand(New WriteTextFileCommand(fileLabel, message, t, sm))
    End Sub

    ''' <summary>
    ''' Sets multiple text messages in the specified file label.
    ''' </summary>
    Public Sub SetTextMultiple(ByVal fileLabel As Char)
        If _TextFileCommands.Count = 0 Then
            Throw New Exception("No text commands were specified; use SetTextMultiple to specify at least one text message first.")
        End If
        SendCommand(_TextFileCommands)
        '-- clear after sending
        _TextFileCommands = New WriteTextFileCommands
    End Sub

    ''' <summary>
    ''' Specifies a text message to be combined with other text messages.
    ''' Once set, a particular file label can be displayed by setting the RunOrder sequence.
    ''' HTML-style markup can be used to specify various display and visual options within the message.
    ''' </summary>
    Public Sub SetTextMultiple(Optional ByVal message As String = "", _
            Optional ByVal t As Transition = Transition.Auto, Optional ByVal sm As Special = Special.None)
        _TextFileCommands.AddTextFile(message, t, sm)
    End Sub

    ''' <summary>
    ''' Sets a string message in the sign's memory.
    ''' Once set, strings can be displayed via the &lt;CallString=(filelabel)&gt; message markup command
    ''' strings can be overwritten in memory without making the sign 'flash', but only support
    ''' a subset of the full message markup commands.
    ''' </summary>
    Public Sub SetString(ByVal fileLabel As Char, ByVal message As String)
        SendCommand(New WriteStringFileCommand(fileLabel, message))
    End Sub

    ''' <summary>
    ''' Loads a 80x7 8-color graphic file from disk, using any valid format, into the sign's memory.
    ''' Once loaded, pictures can be displayed via the &lt;CallPic=(filelabel)&gt; message markup command
    ''' </summary>
    Public Sub SetPicture(ByVal fileLabel As Char, ByVal filename As String)
        If IO.Path.GetFileName(filename) = filename Then
            filename = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, filename)
        End If
        SetPicture(fileLabel, New Bitmap(filename))
    End Sub

    ''' <summary>
    ''' Loads a 80x7 8-color graphic bitmap object into the sign's memory.
    ''' Once loaded, pictures can be displayed via the &lt;CallPic=(filelabel)&gt; message markup command
    ''' </summary>
    Public Sub SetPicture(ByVal fileLabel As Char, ByVal b As Bitmap)
        Dim pc As New WritePictureFileCommand(fileLabel, b)
        SendCommand(pc)
        '-- it can take some time for the graphic to arrive..
        Threading.Thread.CurrentThread.Sleep(50)
    End Sub

    ''' <summary>
    ''' closes the communication channel between the PC and the sign.
    ''' This is NOT done automatically, so it should be called when
    ''' you're done with the sign.
    ''' </summary>
    Public Sub Close()
        If Not _RS232 Is Nothing Then
            _RS232.PortOpen = False
        End If
    End Sub

    ''' <summary>
    ''' opens the communication channel between the PC and the sign
    ''' this can be done explicitly, or it will automatically happen
    ''' when the first commands is sent to the sign
    ''' </summary>
    Public Sub Open(ByVal comPort As Integer)
        _RS232.CommPort = comPort
        Open()
    End Sub

    ''' <summary>
    ''' opens the communication channel between the PC and the sign
    ''' this can be done explicitly, or it will automatically happen
    ''' when the first commands is sent to the sign
    ''' </summary>
    Public Sub Open()
        If IsOpen Then Return
        '-- BetaBrite requires 9600,N,8,1 serial communication
        _RS232 = New BetaBrite.RS232
        With _RS232
            .BaudRate = BaudRates.Baud9600
            .DataSize = DataSizes.Size8
            .Parity = Parities.None
            .StopBit = StopBits.Bit1
            .InBufferSize = 4096
            .InputLength = 0
            .ReceiveThreshold = 1
            .PortOpen = True
        End With
    End Sub

    ''' <summary>
    ''' Clear the sign's memory completely; this also causes it to go into the 
    ''' default attract sequence (which is also a pretty good demo of everything
    ''' you can do programmatically using this class!)
    ''' </summary>
    Public Sub ClearMemory()
        SendCommand((New ClearMemoryCommand))
    End Sub

    ''' <summary>
    ''' Performs a non-destructive reset of the sign. Memory contents ARE retained.
    ''' </summary>
    Public Sub Reset()
        SendCommand((New ResetCommand))
    End Sub

    Private Sub Write(ByVal b As Byte())
        If Not IsOpen Then
            Open()
        End If
        _RS232.Output(b)
    End Sub

    Private Sub SendCommand(ByVal c As BetaBrite.Protocol.BaseCommand)
        If _IsDebug Then
            Debug.WriteLine(c.ToString)
        End If
        Write(c.ToBytes)
    End Sub

End Class
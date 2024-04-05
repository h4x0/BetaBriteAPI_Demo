Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Drawing

''' <summary>
''' returns Alpha Sign Communications Protocol commands to control the BetaBrite LED sign
''' adapted from Alpha Sign Communications Protocol document at http://www.ams-i.com/Pages/97088061.htm
''' </summary>
'''<remarks>
'''  Jeff Atwood
'''  http://www.codinghorror.com/
''' </remarks>
Public Class Protocol

    Private Enum Ascii As Byte
        NUL = &H0
        SOH = &H1
        STX = &H2
        EOT = &H4
        ESC = &H1B
        CR = &HD
    End Enum

    Private Enum SignType As Byte
        AllSigns = &H5A    ' "Z"
        BetaBrite = &H5E   ' "^"
        '-- dozens more in protocol doc (!)
    End Enum

    Protected Enum CommandCode As Byte
        WriteTextFile = 65                ' "A"
        ReadTextFile = 66                 ' "B"
        WriteSpecialFunction = 69         ' "E"
        ReadSpecialFunction = 70          ' "F"
        WriteStringFile = 71              ' "G"
        ReadStringFile = 72               ' "H"
        WritePictureFile = 73             ' "I"
        ReadPictureFile = 74              ' "J"
    End Enum

    Private Shared _Address As String = "00"
    Private Const _SignType As SignType = SignType.BetaBrite
    Private Const _NullHeaderCount As Integer = 1
    Private Const _MaxPictureWidth As Integer = 80
    Private Const _MaxPictureHeight As Integer = 7

    Public Sub New(Optional ByVal address As String = "00")
        Me.Address = address
    End Sub

    ''' <summary>
    ''' The identifier or “address” of the sign represented by two ASCII digits as a number between “00” and “FF” (0 to 255). 
    ''' Address "00" is reserved as a broadcast address. The wildcard character “?” (3FH) can be used to send messages to a 
    ''' range of addresses. For example, a Sign Address of “0?” will access signs with address between 01H and 0FH (1 and 15).
    ''' Defaults to "00"
    ''' </summary>
    Public Property Address() As String
        Get
            Return _Address
        End Get
        Set(ByVal Value As String)
            If Not Regex.IsMatch(Value, "^[0-9a-fA-F?]{2}$") Then
                Throw New ArgumentOutOfRangeException("Address", Value, "Address must be two ASCII hex digits or the wildcard character '?'")
            End If
            _Address = Value
        End Set
    End Property

#Region "  Base/Core Classes"

    ''' <summary>
    ''' character betweeh 20h (32) and 7fH (127)
    ''' </summary>
    ''' <remarks>
    '''  File Label 0 (30H) is used for a Priority TEXT file
    '''  File Label 0 (30H) and ? (3FH) cannot be used as STRING file labels.
    ''' </remarks>
    Public Class File
        Private _LabelByte As Byte = &H41 ' "A"

        Public Sub New()
        End Sub

        Public Sub New(ByVal c As Char)
            Me.Label = c
        End Sub

        Public Sub New(ByVal s As String)
            Me.LabelString = s
        End Sub

        Public Property LabelInt() As Integer
            Get
                Return Convert.ToInt32(_LabelByte)
            End Get
            Set(ByVal Value As Integer)
                If Value > 94 Or Value < 0 Then
                    Throw New ArgumentOutOfRangeException("LabelInt", "Label must be between 0 and 94")
                End If
                _LabelByte = Convert.ToByte(Value + 20)
            End Set
        End Property
        Public Property LabelString() As String
            Get
                Return Convert.ToString(Me.Label)
            End Get
            Set(ByVal Value As String)
                If Value.Length > 1 Then
                    Throw New ArgumentOutOfRangeException("LabelString", "Label must be a single ASCII character.")
                End If
                Me.Label = Convert.ToChar(Value)
            End Set
        End Property
        Public Property Label() As Char
            Get
                Return Convert.ToChar(_LabelByte)
            End Get
            Set(ByVal Value As Char)
                If Not Regex.IsMatch(Value, "[\x20-\x7e]") Then
                    Throw New ArgumentOutOfRangeException("Label", "Label must be a single ASCII character between 'space' (20h) and 'half-space' (7eh)")
                End If
                _LabelByte = Encoding.ASCII.GetBytes(Value)(0)
            End Set
        End Property
        Public Property [Byte]() As Byte
            Get
                Return _LabelByte
            End Get
            Set(ByVal Value As Byte)
                If Value > 94 Or Value < 0 Then
                    Throw New ArgumentOutOfRangeException("FileLabel", "FileLabel must be between 20h and 7eh")
                End If
                _LabelByte = Value
            End Set
        End Property
    End Class

    ''' <summary>
    ''' Base class for an Alpha Communications Protocol command
    ''' </summary>
    Public MustInherit Class BaseCommand
        Protected CommandCode As CommandCode
        Public IsNested As Boolean = False

        Public Overridable Shadows Function ToString() As String
            Return BytesToString(Me.ToBytes())
        End Function

        Public Overridable Function ToBytes() As Byte()
            If CommandCode = 0 Then
                Throw New ArgumentNullException("CommandCode", "The CommandCode was not specified")
            End If

            Dim ms As New IO.MemoryStream

            If Not IsNested Then
                '-- A minimum of 5 <NUL>s must be transmitted;
                '-- the sign uses these to establish the baud rate
                For i As Integer = 0 To _NullHeaderCount - 1
                    ms.WriteByte(Ascii.NUL)
                Next

                '-- the <SOH> is the "Start of Header" ASCII character
                ms.WriteByte(Ascii.SOH)

                '-- sign type we're addressing with this command
                ms.WriteByte(_SignType)

                '-- sign address, if you have multiple signs
                '-- this is set via dip switches or firmware
                ms.Write(Encoding.ASCII.GetBytes(_Address), 0, 2)
            End If

            '-- The <STX> is the "Start of Text" ASCII character; it always precedes a command code
            '-- NOTE: When nesting packets, there must be at least a 100-millisecond delay after the <STX>.
            ms.WriteByte(Ascii.STX)

            '-- one ASCII character defines the command
            ms.WriteByte(CommandCode)

            '-- this calls the method on the inherited child
            Dim DataField As String = FormDataField()

            '-- made up of ASCII characters which are dependent on the preceding Command Code
            ms.Write(Encoding.ASCII.GetBytes(DataField), 0, DataField.Length)

            '-- The <EOT> is the "End of Text" ASCII character
            ms.WriteByte(Ascii.EOT)

            ms.Close()
            Return ms.ToArray
        End Function

        '-- this calls the method on the inherited child
        Protected MustOverride Function FormDataField() As String
    End Class

#End Region

#Region "  TextFileCommand"

    Public Enum LinePosition As Byte
        Middle = &H20   ' space
        '-- betabrite is a single line sign
        'Top = &H22      ' double quote
        'Bottom = &H26   ' &
        'Fill = &H30     ' 0
        'Left = &H31     ' 1
        'Right = &H32    ' 2
    End Enum

    Public Enum Transition As Byte
        ''' <summary>
        ''' Message travels right to left
        ''' </summary>
        Rotate = &H61           ' a
        ''' <summary>
        ''' Message remains stationary
        ''' </summary>
        Hold = &H62             ' b
        ''' <summary>
        ''' Message remains stationary and flashes
        ''' </summary>
        Flash = &H63            ' c
        ''' <summary>
        ''' Previous message is pushed up by a new message
        ''' </summary>
        RollUp = &H65           ' e
        ''' <summary>
        ''' Previous message is pushed down by a new message
        ''' </summary>
        RollDown = &H66         ' f
        ''' <summary>
        ''' Previous message is pushed left by a new message
        ''' </summary>
        RollLeft = &H67         ' g
        ''' <summary>
        ''' Previous message is pushed right by a new message
        ''' </summary>
        RollRight = &H68        ' h
        ''' <summary>
        ''' New message is wiped over the previous message from bottom to top
        ''' </summary>
        WipeUp = &H69           ' i
        ''' <summary>
        ''' New message is wiped over the previous message from top to bottom
        ''' </summary>
        WipeDown = &H6A         ' j
        ''' <summary>
        ''' New message is wiped over the previous message from right to left
        ''' </summary>
        WipeLeft = &H6B         ' k
        ''' <summary>
        ''' New message is wiped over the previous message from left to right
        ''' </summary>
        WipeRight = &H6C        ' l
        ''' <summary>
        ''' New message line pushes the bottom line to the top line if 2-line sign
        ''' </summary>
        Scroll = &H6D           ' m
        ''' <summary>
        ''' Various Modes are called upon to display the message automatically
        ''' </summary>
        [Auto] = &H6F           ' n
        ''' <summary>
        ''' Previous message is pushed toward the center of the display by the new message
        ''' </summary>
        RollIn = &H70           ' o
        ''' <summary>
        ''' Previous message is pushed outward from the center by the new message
        ''' </summary>
        RollOut = &H71          ' p
        ''' <summary>
        ''' New message is wiped over the previous message in an inward motion
        ''' </summary>
        WipeIn = &H72           ' q
        ''' <summary>
        ''' New message is wiped over the previous message in an outward motion
        ''' </summary>
        WipeOut = &H73          ' s
        ''' <summary>
        ''' Message travels right to left. Characters are approximately one half their normal width
        ''' </summary>
        CompressedRotate = &H74 ' t
        ''' <summary>
        ''' Special transition defined in the Special enumeration
        ''' </summary>
        Special = &H6E          ' n
    End Enum

    Public Enum Special As Byte
        ''' <summary>
        ''' no special mode selected; used as a default only
        ''' </summary>
        None = &H0          ' NUL
        ''' <summary>
        ''' Message will twinkle on the sign (all LEDs will flash rapidly)
        ''' </summary>
        Twinkle = &H30      ' 0
        ''' <summary>
        ''' New message will sparkle over the current message
        ''' </summary>
        Sparkle = &H31      ' 1
        ''' <summary>
        ''' Message will drift on to the display like falling snow
        ''' </summary>
        Snow = &H32         ' 2
        ''' <summary>
        ''' New message will interlock over the current message in alternating rows of dots from each end
        ''' </summary>
        Interlock = &H33    ' 3
        ''' <summary>
        ''' Alternating characters “switch” off the sign up and down. New message “switches” on in a similar manner
        ''' </summary>
        Switch = &H34       ' 4
        ''' <summary>
        ''' New message sprays across and onto the sign from right to left
        ''' </summary>
        Spray = &H36        ' 6
        ''' <summary>
        ''' small explosions blast the new message onto the sign
        ''' </summary>
        Starburst = &H37    ' 7
        ''' <summary>
        ''' welcome is written in script across the sign and changes multiple colors
        ''' </summary>
        Welcome = &H38      ' 8
        ''' <summary>
        ''' Slot machine symbols appear randomly across the sign
        ''' </summary>
        SlotMachine = &H39  ' 9
        ''' <summary>
        ''' Satellite dish broadcasts the words News Flash on the sign
        ''' </summary>
        NewsFlash = &H41    ' A
        ''' <summary>
        ''' Animated trumpet blows multicolored notes across the sign
        ''' </summary>
        Trumpet = &H42      ' B
        ''' <summary>
        ''' color changes from one color to another
        ''' </summary>
        CycleColors = &H43  ' C
        ''' <summary>
        ''' Thank You is written in script across the sign and changes multiple colors
        ''' </summary>
        ThankYou = &H53     ' S
        ''' <summary>
        ''' A cigarette image appears, is then extinguished and replaced with a no smoking symbol
        ''' </summary>
        NoSmoking = &H55    ' U
        ''' <summary>
        ''' A car runs into a cocktail glass and is replaced with the text “Please don’t drink and drive”
        ''' </summary>
        DontDrink = &H56    ' V
        ''' <summary>
        ''' fish swim across the sign, then are chased back across it by a shark
        ''' </summary>
        Fish = &H57         ' W
        ''' <summary>
        ''' Large fireworks explode randomly across the sign
        ''' </summary>
        Fireworks = &H58    ' X
        ''' <summary>
        ''' Party baloons scroll up the display
        ''' </summary>
        Balloon = &H59      ' Y
        ''' <summary>
        ''' A bomb fuse burns down followed by a giant explosion
        ''' </summary>
        CherryBomb = &H5A   ' Z
    End Enum

    Public Enum FileType As Byte
        ''' <summary>
        ''' File represents displayable text
        ''' </summary>
        Text = &H41         ' A
        ''' <summary>
        ''' File represents a string variable
        ''' </summary>
        [String] = &H42     ' B 
        ''' <summary>
        ''' File represents a bitmap image
        ''' </summary>
        Picture = &H44      ' D
    End Enum

    Public Enum Protection As Byte
        ''' <summary>
        ''' Can be changed via the infrared remote
        ''' </summary>
        Unlocked = &H55     ' U
        ''' <summary>
        ''' Cannot be changed via the infrared remote
        ''' </summary>
        Locked = &H4C       ' L
    End Enum

    ''' <summary>
    ''' TextFile used in SetTextCommand and SetTextCommands
    ''' </summary>
    Private Class TextFile
        Private _Position As LinePosition = LinePosition.Middle
        Private _Message As String
        Public TransitionType As Transition = Transition.Auto
        Public SpecialMode As Special = Special.None

        Public Sub New()
        End Sub

        Public Sub New(ByVal message As String, Optional ByVal t As Transition = Transition.Auto, _
            Optional ByVal s As Special = Special.None)
            Me.Message = message
            Me.TransitionType = t
            Me.SpecialMode = s
        End Sub

        Public Property Message() As String
            Get
                Return _Message
            End Get
            Set(ByVal Value As String)
                Value = ExpandMessage(Value)
                If Not Regex.IsMatch(Value, "[\x00-\x7f]*") Then
                    Throw New ArgumentOutOfRangeException("Message", "Text messages can only contain ASCII characters in the range 00-7F.")
                End If
                _Message = Value
            End Set
        End Property

        Public Overrides Function ToString() As String
            Dim sb As New StringBuilder
            sb.Append(Convert.ToChar(Ascii.ESC))
            sb.Append(Convert.ToChar(_Position))
            sb.Append(Convert.ToChar(TransitionType))
            If Me.TransitionType = Transition.Special Then
                If SpecialMode = Special.None Then
                    Throw New ArgumentException("TransitionType was set to special, but SpecialMode was not specified (None).", "SpecialMode")
                End If
                sb.Append(Convert.ToChar(SpecialMode))
            End If
            '-- An ASCII Message cannot be displayed if the previous field (Special Specifier) is a Special
            '-- Graphic. To display text after a Special Graphic, another Mode Field must be used.
            sb.Append(_Message)
            Return sb.ToString
        End Function
    End Class

    ''' <summary>
    ''' Write a single text command to a File
    ''' </summary>
    Public Class WriteTextFileCommand
        Inherits BaseCommand
        Private _TextFile As TextFile
        Public File As New File

        Public Sub New()
            Me.CommandCode = CommandCode.WriteTextFile
        End Sub

        Public Sub New(ByVal fileLabel As Char, Optional ByVal message As String = "", _
            Optional ByVal t As Transition = Transition.Auto, Optional ByVal sm As Special = Special.None)
            MyClass.new()
            Me.File.Label = fileLabel
            _TextFile = New TextFile
            With _TextFile
                .Message = message
                .TransitionType = t
                .SpecialMode = sm
            End With
        End Sub

        Protected Overrides Function FormDataField() As String
            Dim sb As New StringBuilder
            sb.Append(File.Label)
            sb.Append(_TextFile.ToString)
            Return sb.ToString
        End Function
    End Class

    ''' <summary>
    ''' Writes multiple TEXT commands to a File
    ''' </summary>
    Public Class WriteTextFileCommands
        Inherits BaseCommand

        Public File As New File
        Private _TextFiles As New ArrayList

        Public Sub New()
            Me.CommandCode = CommandCode.WriteTextFile
        End Sub

        Public Sub New(ByVal fileLabel As Char)
            MyClass.New()
            Me.File.Label = fileLabel
        End Sub

        Public ReadOnly Property Count()
            Get
                Return _TextFiles.Count
            End Get
        End Property

        Public Sub AddTextFile(ByVal message As String, _
            Optional ByVal t As Transition = Transition.Auto, Optional ByVal sm As Special = Special.None)
            Dim tf As New TextFile(message, t, sm)
            _TextFiles.Add(tf)
        End Sub

        Protected Overrides Function FormDataField() As String
            Dim sb As New StringBuilder
            sb.Append(File.Label)
            If _TextFiles.Count = 0 Then
                Throw New ArgumentException("TextCommand command contains no TextFiles; use AddTextFile to add items first")
            End If
            For Each tf As TextFile In _TextFiles
                sb.Append(tf.ToString)
            Next
            Return sb.ToString
        End Function

    End Class

#End Region

#Region "  PictureFileCommand"

    ''' <summary>
    ''' Converts an image file from disk into a BetaBrite picture file
    ''' dimensions can be maximum of 255 width x 31 height, and up to 8 colors
    ''' </summary>
    Public Class WritePictureFileCommand
        Inherits BaseCommand
        Public File As New File
        Private _Height As Integer = _MaxPictureHeight
        Private _Width As Integer = _MaxPictureWidth
        Private _Picture As Bitmap
        Private _ColorArray(8) As Drawing.Color
        Private _ColorNames(8) As String

        Private Sub New()
            MyBase.CommandCode = CommandCode.WritePictureFile
            _ColorArray(0) = Drawing.Color.FromArgb(0, 0, 0)
            _ColorNames(0) = "Black"
            _ColorArray(1) = Drawing.Color.FromArgb(255, 0, 0)
            _ColorNames(1) = "Red"
            _ColorArray(2) = Drawing.Color.FromArgb(0, 255, 0)
            _ColorNames(2) = "Green"
            _ColorArray(3) = Drawing.Color.FromArgb(254, 204, 2)
            _ColorNames(3) = "Amber"
            _ColorArray(4) = Drawing.Color.FromArgb(128, 0, 0)
            _ColorNames(4) = "DarkRed"
            _ColorArray(5) = Drawing.Color.FromArgb(0, 128, 0)
            _ColorNames(5) = "DarkGreen"
            _ColorArray(6) = Drawing.Color.FromArgb(202, 134, 54)
            _ColorNames(6) = "Brown"
            _ColorArray(7) = Drawing.Color.FromArgb(255, 165, 0)
            _ColorNames(7) = "Orange"
            _ColorArray(8) = Drawing.Color.FromArgb(255, 255, 0)
            _ColorNames(8) = "Yellow"
        End Sub

        Public Sub New(ByVal fileLabel As Char, ByVal b As Bitmap)
            Me.New()
            File.Label = fileLabel
            _Picture = b
            Me.Width = _Picture.Width
            Me.Height = _Picture.Height
        End Sub

        Public Sub New(ByVal fileLabel As Char, ByVal bitmapFilePath As String)
            Me.New(fileLabel, New Bitmap(bitmapFilePath))
        End Sub

        Private Property Width() As Integer
            Get
                Return _Width
            End Get
            Set(ByVal Value As Integer)
                If Value <= 0 Or Value > _MaxPictureWidth Then
                    Throw New ArgumentException("Picture file width is " & Value & ", but should be between 1 and " & _MaxPictureWidth, "Width")
                End If
                _Width = Value
            End Set
        End Property

        Private Property Height() As Integer
            Get
                Return _Height
            End Get
            Set(ByVal Value As Integer)
                If Value <= 0 Or Value > _MaxPictureHeight Then
                    Throw New ArgumentException("Picture file height is " & Value & ", but should be between 1 and " & _MaxPictureHeight, "Width")
                End If
                _Height = Value
            End Set
        End Property

        Private Function MapToBetaBriteColor(ByVal pc As System.Drawing.Color) As Char

            Dim d As Double = 0
            Dim dmin As Double = Double.MaxValue
            Dim match As Integer = 0

            For i As Integer = 0 To 8
                d = ColorDifference(pc, _ColorArray(i))
                If d = 0 Then
                    '-- perfect match!
                    match = i
                    Exit For
                Else
                    If d < dmin Then
                        match = i
                        dmin = d
                    End If
                End If
            Next

            'Debug.WriteLine("best match:" & _ColorNames(match) & " m=" & minm)
            Return Convert.ToChar(48 + match)
        End Function

        Private Function ColorDifference(ByVal a As System.Drawing.Color, ByVal b As System.Drawing.Color) As Double
            Dim i As Integer = _
                Convert.ToInt32(Math.Abs(Convert.ToInt32(a.R) - Convert.ToInt32(b.R)) ^ 2) + _
                Convert.ToInt32(Math.Abs(Convert.ToInt32(a.G) - Convert.ToInt32(b.G)) ^ 2) + _
                Convert.ToInt32(Math.Abs(Convert.ToInt32(a.B) - Convert.ToInt32(b.B)) ^ 2)
            If i > 0 Then
                Return Math.Sqrt(i)
            Else
                Return 0
            End If
        End Function

        Protected Overrides Function FormDataField() As String
            Dim sb As New StringBuilder(_Height * _Width + 5)
            sb.Append(File.Label)
            sb.Append(String.Format("{0:x2}", _Height))
            sb.Append(String.Format("{0:x2}", _Width))

            Dim q As Byte = &H11

            For row As Integer = 0 To _Height - 1
                For col As Integer = 0 To _Width - 1
                    sb.Append(MapToBetaBriteColor(_Picture.GetPixel(col, row)))
                Next
                sb.Append(Convert.ToChar(Ascii.CR))
            Next
            Return sb.ToString.ToUpper
        End Function

    End Class
#End Region

#Region "  StringFileCommand"

    Public Class WriteStringFileCommand
        Inherits BaseCommand

        Public File As New File
        Private _FileData As String

        Public Property FileData() As String
            Get
                Return _FileData
            End Get
            Set(ByVal Value As String)
                If Not Regex.IsMatch(Value, "[\x20-\x7f\x09\x0d\x11-\x13\x15-\x19\x1a\x1c\1xe]*") Then
                    Throw New ArgumentOutOfRangeException("StringData", "Strings can only contain ASCII characters 20-7F or a subset of the formatting codes.")
                End If
                _FileData = Value
            End Set
        End Property

        Public Sub New(ByVal fileLabel As Char, Optional ByVal fileData As String = "")
            MyBase.CommandCode = CommandCode.WriteStringFile
            Me.FileData = fileData
            Me.File.Label = fileLabel
        End Sub

        Protected Overrides Function FormDataField() As String
            Return File.Label & _FileData
        End Function
    End Class

#End Region

#Region "  SpecialCommand"

    Public Enum SpecialFunction As Byte
        TimeOfDay = &H20        ' space
        Memory = &H24           ' $
        DayOfWeek = &H26        ' &
        TimeFormat = &H27       ' double quote
        RunTime = &H29          ' )
        SoftReset = &H2C        ' comma
        RunSequence = &H2E      ' period
        RunDay = &H32           '  2
        ClearSerialError = &H34 ' 4
        Address = &H37          ' 7
        [Date] = &H3B           ' ;
    End Enum

    ''' <summary>
    ''' MemoryItem used in SetMemoryCommand and SetMemoryCommands
    ''' </summary>
    Private Class MemoryItem
        Public File As New File
        Public FileType As FileType = FileType.Text
        Public Protection As Protection = Protection.Unlocked
        Public SizeBytes As Integer = 0
        Public StartTime As DateTime = Nothing
        Public StopTime As DateTime = Nothing
        Public Width As Integer = 0
        Public Height As Integer = 0

        Public Overrides Function ToString() As String
            Dim sb As New StringBuilder
            sb.Append(File.Label)
            sb.Append(Convert.ToChar(FileType))
            '-- protection type
            If FileType = FileType.String Then
                '-- string files MUST be locked
                sb.Append(Convert.ToChar(Protection.Locked))
            Else
                sb.Append(Convert.ToChar(Me.Protection))
            End If
            '-- 4-digit ASCII hex size
            Select Case FileType
                Case FileType.String, FileType.Text
                    If SizeBytes = 0 Then
                        Throw New ArgumentOutOfRangeException("SizeBytes", "File size, in bytes, must be specified.")
                    End If
                    sb.Append(String.Format("{0:x4}", SizeBytes))
                Case FileType.Picture
                    If Width = 0 Then
                        Throw New ArgumentOutOfRangeException("Width", "Picture width, in pixels, must be specified.")
                    End If
                    If Width > _MaxPictureWidth Then
                        Throw New ArgumentOutOfRangeException("Width", "Picture width, in pixels, must be less than " & _MaxPictureWidth)
                    End If
                    If Height = 0 Then
                        Throw New ArgumentOutOfRangeException("Height", "Picture height, in pixels, must be specified.")
                    End If
                    If Height > _MaxPictureHeight Then
                        Throw New ArgumentOutOfRangeException("Height", "Picture height, in pixels, must be less than " & _MaxPictureHeight)
                    End If
                    sb.Append(String.Format("{0:x2}", Height))
                    sb.Append(String.Format("{0:x2}", Width))
            End Select
            '-- misc data per-type
            Select Case FileType
                Case FileType.Text
                    If StartTime = Nothing Or StopTime = Nothing Then
                        sb.Append(TimeToByteAscii(TimeConstant.Always))
                        sb.Append(TimeToByteAscii(TimeConstant.StartOfDay))
                    Else
                        sb.Append(TimeToByteAscii(StartTime))
                        sb.Append(TimeToByteAscii(StopTime))
                    End If
                Case FileType.String
                    '-- constant padding; not used
                    sb.Append("0000")
                Case FileType.Picture
                    'Valid entries are “1000” = monochrome, “2000” = 3-color, “4000”= 8-color (RGB)
                    '-- always 4000 for betabrite..
                    sb.Append("4000")
            End Select
            Return sb.ToString.ToUpper
        End Function
    End Class

    ''' <summary>
    ''' Stores a single item in the sign's memory
    ''' (via constructor only)
    ''' </summary>
    Public Class SetMemoryCommand
        Inherits BaseCommand
        Private _MemoryItem As MemoryItem

        Public Sub New()
            MyBase.CommandCode = CommandCode.WriteSpecialFunction
        End Sub

        Public Sub New(ByVal fileLabel As Char, ByVal ft As FileType, ByVal p As Protection, _
            ByVal sizeBytes As Integer, ByVal startTime As DateTime, ByVal stopTime As DateTime)
            MyClass.new()
            _MemoryItem = New MemoryItem
            With _MemoryItem
                .File.Label = fileLabel
                .FileType = ft
                .Protection = p
                .SizeBytes = sizeBytes
                .StartTime = startTime
                .StopTime = stopTime
            End With
        End Sub

        Public Sub New(ByVal fileLabel As Char, ByVal ft As FileType, ByVal p As Protection, ByVal sizeBytes As Integer)
            MyClass.New(fileLabel, ft, p, sizeBytes, Nothing, Nothing)
        End Sub

        Public Sub New(ByVal fileLabel As Char, ByVal ft As FileType, ByVal p As Protection, ByVal message As String)
            MyClass.New(fileLabel, ft, p, ExpandedMessageLength(message), Nothing, Nothing)
        End Sub

        Public Sub New(ByVal fileLabel As Char, ByVal ft As FileType, ByVal p As Protection, ByVal message As String, _
            ByVal startTime As DateTime, ByVal stopTime As DateTime)
            MyClass.New(fileLabel, ft, p, ExpandedMessageLength(message), startTime, stopTime)
        End Sub

        Protected Overrides Function FormDataField() As String
            Dim sb As New StringBuilder
            sb.Append(Convert.ToChar(SpecialFunction.Memory))
            sb.Append(_MemoryItem.ToString)
            Return sb.ToString
        End Function
    End Class

    ''' <summary>
    ''' Stores multiple items in the sign's memory
    ''' </summary>
    Public Class SetMemoryCommands
        Inherits BaseCommand

        Private _MemoryItems As New ArrayList

        Public Sub New()
            MyBase.CommandCode = CommandCode.WriteSpecialFunction
        End Sub

        Public Sub AllocateTextFile(ByVal fileLabel As Char, ByVal p As Protection, ByVal sizeBytes As Integer)
            Dim mi As New MemoryItem
            mi.File.Label = fileLabel
            mi.FileType = FileType.Text
            mi.Protection = p
            mi.SizeBytes = sizeBytes
            _MemoryItems.Add(mi)
        End Sub

        Public ReadOnly Property Count()
            Get
                Return _MemoryItems.Count
            End Get
        End Property

        Public Sub AllocateTextFile(ByVal fileLabel As Char, ByVal p As Protection, ByVal message As String)
            AllocateTextFile(fileLabel, p, ExpandedMessageLength(message))
        End Sub

        Public Sub AllocateStringFile(ByVal fileLabel As Char, ByVal sizeBytes As Integer)
            Dim mi As New MemoryItem
            mi.File.Label = fileLabel
            mi.FileType = FileType.String
            mi.Protection = Protection.Locked
            mi.SizeBytes = sizeBytes
            _MemoryItems.Add(mi)
        End Sub

        Public Sub AllocateStringFile(ByVal fileLabel As Char, ByVal message As String)
            AllocateStringFile(fileLabel, ExpandedMessageLength(message))
        End Sub

        Public Sub AllocatePictureFile(ByVal fileLabel As Char, ByVal p As Protection, ByVal width As Integer, ByVal height As Integer)
            Dim mi As New MemoryItem
            mi.File.Label = fileLabel
            mi.FileType = FileType.Picture
            mi.Protection = p
            mi.Width = width
            mi.Height = height
            _MemoryItems.Add(mi)
        End Sub

        Public Sub AllocatePictureFile(ByVal fileLabel As Char)
            AllocatePictureFile(fileLabel, Protection.Unlocked, 80, 7)
        End Sub

        Protected Overrides Function FormDataField() As String
            If _MemoryItems.Count = 0 Then
                Throw New ArgumentException("Memory command contains no items; use AddMemoryItem to add items first")
            End If
            Dim sb As New StringBuilder
            sb.Append(Convert.ToChar(SpecialFunction.Memory))
            For Each mi As MemoryItem In _MemoryItems
                sb.Append(mi.ToString)
            Next
            Return sb.ToString
        End Function
    End Class

    Public Enum RunSequenceOrder
        ByTime = &H54
        ByOrder = &H53
        ByTimeThenDelete = &H44
    End Enum

    ''' <summary>
    ''' Sets the run (display) sequence for a series of 1-128 text files. If a file label is invalid or 
    ''' does not exist, the next one in the sequence will run.
    ''' </summary>
    Public Class SetRunSequenceCommand
        Inherits BaseCommand

        Private _Files As New ArrayList
        Public Order As RunSequenceOrder = RunSequenceOrder.ByOrder
        Public Protection As Protection = Protection.Unlocked

        Public Sub New()
            MyBase.CommandCode = CommandCode.WriteSpecialFunction
        End Sub

        Public Sub New(ByVal rso As RunSequenceOrder, ByVal p As Protection)
            MyClass.new()
            Me.Order = rso
            Me.Protection = p
        End Sub

        Public Sub AddFile(ByVal fileLabel As Char)
            If _Files.Count > 128 Then
                Throw New ArgumentOutOfRangeException("AddFile", "Run sequences cannot exceed 128 text files.")
            End If
            _Files.Add(New File(fileLabel))
        End Sub

        Protected Overrides Function FormDataField() As String
            If _Files.Count = 0 Then
                Throw New ArgumentException("Run sequence contains no files; use AddFile to add files.")
            End If
            Dim sb As New StringBuilder
            sb.Append(Convert.ToChar(SpecialFunction.RunSequence))
            sb.Append(Convert.ToChar(Order))
            sb.Append(Convert.ToChar(Protection))
            For Each f As File In _Files
                sb.Append(f.Label)
            Next
            Return sb.ToString
        End Function
    End Class

    ''' <summary>
    ''' Set the times to start and stop displaying a particular Text file
    ''' </summary>
    Public Class SetRunTimeCommand
        Inherits BaseCommand
        Public Start As DateTime = DateTime.MinValue
        Public [Stop] As DateTime = DateTime.MaxValue
        Public File As File

        Public Sub New()
            MyBase.CommandCode = CommandCode.WriteSpecialFunction
        End Sub

        Public Sub New(ByVal fileLabel As Char, ByVal startTime As DateTime, ByVal stopTime As DateTime)
            MyClass.New()
            Me.File.Label = fileLabel
            Me.Start = startTime
            Me.Stop = stopTime
        End Sub

        Protected Overrides Function FormDataField() As String
            Dim sb As New StringBuilder
            sb.Append(Convert.ToChar(SpecialFunction.RunTime))
            sb.Append(File.Label)
            sb.Append(TimeToByteAscii(Start))
            sb.Append(TimeToByteAscii([Stop]))
            Return sb.ToString
        End Function
    End Class

    Public Enum RunDay As Byte
        Daily = &H30
        Sun = &H31
        Mon = &H32
        Tue = &H33
        Wed = &H34
        Thu = &H35
        Fri = &H36
        Sat = &H37
        MonToFri = &H38
        SatSun = &H39
        Always = &H41
        Never = &H42
    End Enum

    ''' <summary>
    ''' Set the days to start and stop displaying a particular Text file
    ''' </summary>
    Public Class SetRunDayCommand
        Inherits BaseCommand

        Public File As File
        Public Start As RunDay = RunDay.Daily
        Public [Stop] As RunDay = RunDay.Always

        Public Sub New()
            MyBase.CommandCode = CommandCode.WriteSpecialFunction
        End Sub

        Public Sub New(ByVal fileLabel As Char, ByVal start As RunDay, ByVal [stop] As RunDay)
            MyClass.New()
            Me.File.Label = fileLabel
            Me.Start = start
            Me.Stop = [stop]
        End Sub

        Protected Overrides Function FormDataField() As String
            Dim sb As New StringBuilder
            sb.Append(Convert.ToChar(SpecialFunction.RunDay))
            sb.Append(File.Label)
            sb.Append(Convert.ToChar(Start))
            sb.Append(Convert.ToChar([Stop]))
            Return sb.ToString
        End Function
    End Class

    ''' <summary>
    ''' Clears all sign memory
    ''' </summary>
    Public Class ClearMemoryCommand
        Inherits BaseCommand
        Public Sub New()
            MyBase.CommandCode = CommandCode.WriteSpecialFunction
        End Sub
        Protected Overrides Function FormDataField() As String
            Return Convert.ToChar(SpecialFunction.Memory)
        End Function
    End Class

    ''' <summary>
    ''' Performs a soft (non-destructive) reset of the sign; memory is retained
    ''' </summary>
    Public Class ResetCommand
        Inherits BaseCommand
        Public Sub New()
            MyBase.CommandCode = CommandCode.WriteSpecialFunction
        End Sub
        Protected Overrides Function FormDataField() As String
            Return Convert.ToChar(SpecialFunction.SoftReset)
        End Function
    End Class

#Region "  Date/Time Commands"

    Public Enum TimeFormat As Byte
        Standard = &H53
        Military = &H4D
    End Enum

    ''' <summary>
    ''' Set format of time displayed on the sign, either
    ''' standard or 24-hour (military)
    ''' </summary>
    Public Class SetTimeFormatCommand
        Inherits BaseCommand
        Private _tf As TimeFormat = TimeFormat.Standard
        Public Sub New(ByVal tf As TimeFormat)
            MyBase.CommandCode = CommandCode.WriteSpecialFunction
        End Sub
        Protected Overrides Function FormDataField() As String
            Return Convert.ToChar(SpecialFunction.TimeFormat) & Convert.ToChar(_tf)
        End Function
    End Class


    ''' <summary>
    ''' Set the date, time, and day of week on the display
    ''' </summary>
    Friend Class SetDateTimeCommand
        Private _TimeCommand As SetTimeCommand
        Private _DayCommand As SetDayCommand
        Private _DateCommand As SetDateCommand
        Private _dt As DateTime

        Public Property [Date]() As DateTime
            Get
                Return _dt
            End Get
            Set(ByVal Value As DateTime)
                _dt = Value
            End Set
        End Property

        Public Sub New(ByVal dt As DateTime)
            _dt = dt
        End Sub

        Public Overrides Function ToString() As String
            Return BytesToString(Me.ToBytes)
        End Function

        Public Function ToBytes() As Byte()
            Dim ms As New MemoryStream

            _TimeCommand = New SetTimeCommand(_dt)
            _DayCommand = New SetDayCommand(_dt)
            _DateCommand = New SetDateCommand(_dt)
            Dim b() As Byte
            b = _TimeCommand.ToBytes
            ms.Write(b, 0, b.Length)
            b = _DayCommand.ToBytes
            ms.Write(b, 0, b.Length)
            b = _DateCommand.ToBytes
            ms.Write(b, 0, b.Length)
            Return ms.ToArray
        End Function
    End Class

    ''' <summary>
    ''' Sets the time on the display
    ''' </summary>
    Private Class SetTimeCommand
        Inherits BaseCommand
        Private _dt As DateTime
        Public Sub New()
            MyBase.CommandCode = CommandCode.WriteSpecialFunction
            _dt = Now
        End Sub
        Public Sub New(ByVal dt As DateTime)
            MyClass.New()
            _dt = dt
        End Sub
        Protected Overrides Function FormDataField() As String
            Return Convert.ToChar(SpecialFunction.TimeOfDay) & _dt.ToString("HHmm")
        End Function
    End Class

    ''' <summary>
    ''' Sets the day of week on the display
    ''' </summary>
    Private Class SetDayCommand
        Inherits BaseCommand
        Private _dt As DateTime
        Public Sub New()
            MyBase.CommandCode = CommandCode.WriteSpecialFunction
            _dt = Now
        End Sub
        Public Sub New(ByVal dt As DateTime)
            MyClass.New()
            _dt = dt
        End Sub
        Protected Overrides Function FormDataField() As String
            Return Convert.ToChar(SpecialFunction.DayOfWeek) & Convert.ToString(Convert.ToInt32(_dt.DayOfWeek) + 1)
        End Function
    End Class

    ''' <summary>
    ''' Sets the date on the display
    ''' </summary>
    Private Class SetDateCommand
        Inherits BaseCommand
        Private _dt As DateTime
        Public Sub New()
            MyBase.CommandCode = CommandCode.WriteSpecialFunction
            _dt = Now
        End Sub
        Public Sub New(ByVal dt As DateTime)
            MyClass.New()
            _dt = dt
        End Sub
        Protected Overrides Function FormDataField() As String
            Return Convert.ToChar(SpecialFunction.Date) & _dt.ToString("MMddyy")
        End Function
    End Class

#End Region

#End Region

#Region "  Text Formatting"

    Private Enum ControlChars As Byte
        Flash = &H7
        ExtChar = &H8
        CallDate = &HB
        NewLine = &HD
        NoHold = &H9
        CallString = &H10
        WideOff = &H11
        WideOn = &H12
        CallTime = &H13
        CallPic = &H14
        Speed1 = &H15
        Speed2 = &H16
        Speed3 = &H17
        Speed4 = &H18
        Speed5 = &H19
        Font = &H1A
        Color = &H1C
        CharAttrib = &H1D
        CharSpacing = &H1E
    End Enum

    Private Enum Color As Byte
        Red = &H31
        Green = &H32
        Amber = &H33
        DimRed = &H34
        DimGreen = &H35
        Brown = &H36
        Orange = &H37
        Yellow = &H38
        Rainbow1 = &H39
        Rainbow2 = &H41
        ColorMix = &H42
        [Auto] = &H43
    End Enum

    Private Enum Font As Byte
        Five = &H31
        FiveBold = &H32
        FiveWide = &H3B
        FiveWideBold = &H3E
        Seven = &H33
        SevenSerif = &H35
        SevenBold = &H34
        SevenBoldSerif = &H36
        SevenShadow = &H37
        SevenShadowSerif = &H3A
        SevenWide = &H3C
        SevenWideBold = &H39
        SevenWideBoldSerif = &H38
        SevenWideSerif = &H3D
    End Enum

    Private Enum ExtChar
        EnterKey = &H62
        YKey = &H63
        UpArrow = &H64
        DownArrow = &H65
        LeftArrow = &H66
        RightArrow = &H67
        Pacman = &H68
        Sailboat = &H69
        Baseball = &H6A
        Telephone = &H6B
        Heart = &H6C
        Car = &H6D
        Handicap = &H6E
        Rhino = &H6F
        Mug = &H70
        Satellite = &H71
        Copyright = &H72
    End Enum

    Private Enum CharAttrib
        Wide = &H30
        DoubleWide = &H31
        '-- these don't appear to work on the BetaBrite
        'DoubleHigh = &H32
        'TrueDescenders = &H33
        'FixedWidth = &H34
        'Fancy = &H35
        'Shadow = &H1D
    End Enum

    ''' <summary>
    ''' Returns the number of bytes in a fully expanded message that contains
    ''' control code tags and/or international characters
    ''' </summary>
    Public Shared Function ExpandedMessageLength(ByVal message As String) As Integer
        Return ExpandMessage(message).Length + 1
    End Function

    ''' <summary>
    ''' Expands message control code tags and/or international characters, if present
    ''' </summary>
    Private Shared Function ExpandMessage(ByVal message As String) As String
        message = ExpandControlCodes(message)
        message = ExpandInternationalChars(message)
        Return message
    End Function

    ''' <summary>
    ''' returns a list of vaguely HTML-style commands and parameters, eg
    ''' &lt;font=fiveslim&gt;Hello &lt;color=red&gt;World
    ''' </summary>
    Private Shared Function ExpandControlCodes(ByVal s As String) As String

        '-- no pseudo-HTML found? nothing to expand; exit
        If Not Regex.IsMatch(s, "<[^>]+?>") Then
            Return s
        End If

        Dim CommandCode As String
        Dim Parameter As String
        Dim Expansion As String
        For Each m As Match In Regex.Matches(s, "<(?<Command>[^=>]+)=*(?<Parameter>[^=>]+)*.*?>")
            CommandCode = m.Groups("Command").Value
            Parameter = m.Groups("Parameter").Value
            Expansion = ExpandControlChar(DirectCast(MapCommandToEnum(CommandCode), ControlChars), Parameter)
            s = s.Replace(m.Value, Expansion)
        Next

        Return s
    End Function

    ''' <summary>
    ''' high ascii chars must be expressed as double-byte chars in a specific BetaBrite format
    ''' </summary>
    Private Shared Function ExpandInternationalChars(ByVal s As String) As String

        '-- no high ascii found? nothing to expand; exit
        If Not Regex.IsMatch(s, "[\x80-\xFF]") Then Return s

        Dim b As String = System.Text.Encoding.ASCII.GetString(System.Text.Encoding.ASCII.GetBytes(s))
        For Each m As Match In Regex.Matches(s, "[\x80-\xFF]")
            s = s.Replace(m.Value, ExpandInternationalChar(Convert.ToChar(m.Value)))
        Next

        Return s
    End Function

    Private Shared Function ExpandInternationalChar(ByVal c As Char) As String
        Dim asciimap As Integer
        Select Case c
            Case "Ç"
                asciimap = &H20
            Case "Ü"
                asciimap = &H21
            Case "é"
                asciimap = &H22
            Case "â"
                asciimap = &H23
            Case "ä"
                asciimap = &H24
            Case "à"
                asciimap = &H25
            Case "å"
                asciimap = &H26
            Case "ç"
                asciimap = &H27
            Case "ê"
                asciimap = &H28
            Case "ë"
                asciimap = &H29
            Case "è"
                asciimap = &H2A
            Case "Ï"
                asciimap = &H2B
            Case "Î"
                asciimap = &H2C
            Case "Ì"
                asciimap = &H2D
            Case "Ä"
                asciimap = &H2E
            Case "Å"
                asciimap = &H2F
            Case "É"
                asciimap = &H30
            Case "æ"
                asciimap = &H31
            Case "Æ"
                asciimap = &H32
            Case "ô"
                asciimap = &H33
            Case "ö"
                asciimap = &H34
            Case "ò"
                asciimap = &H35
            Case "Û"
                asciimap = &H36
            Case "ù"
                asciimap = &H37
            Case "Ÿ"
                asciimap = &H38
            Case "Ö"
                asciimap = &H39
            Case "Ü"
                asciimap = &H3A
            Case "¢"
                asciimap = &H3B
            Case "£"
                asciimap = &H3C
            Case "¥"
                asciimap = &H3D
            Case "ƒ"
                asciimap = &H3F
            Case "á"
                asciimap = &H40
            Case "í"
                asciimap = &H41
            Case "ó"
                asciimap = &H42
            Case "ú"
                asciimap = &H43
            Case "ñ"
                asciimap = &H44
            Case "Ñ"
                asciimap = &H45
            Case "¿"
                asciimap = &H48
            Case "°"
                asciimap = &H49
            Case "¡"
                asciimap = &H4A
            Case " "
                asciimap = &H4B
            Case "ø"
                asciimap = &H4D
            Case "Ø"
                asciimap = &H4C
            Case "c"
                asciimap = &H4E
            Case "C"
                asciimap = &H4F
            Case "c"
                asciimap = &H50
            Case "C"
                asciimap = &H51
            Case "d"
                asciimap = &H52
            Case "Ð"
                asciimap = &H53
            Case "Š"
                asciimap = &H54
            Case "Ž"
                asciimap = &H55
            Case "ž"
                asciimap = &H56
            Case "ß"
                asciimap = &H57
            Case "š"
                asciimap = &H58
            Case "Á"
                asciimap = &H5A
            Case "À"
                asciimap = &H5B
            Case "Ã"
                asciimap = &H5C
            Case "ã"
                asciimap = &H5D
            Case "Ê"
                asciimap = &H5E
            Case "Í"
                asciimap = &H5F
            Case "Õ"
                asciimap = &H60
            Case "õ"
                asciimap = &H61
            Case Else
                '-- remove anything we can't map; it'll be illegal anyway
                Return ""
        End Select

        Return Convert.ToChar(ControlChars.ExtChar) & Convert.ToChar(asciimap)
    End Function

    Private Shared Function ExpandCharAttrib(ByVal param As String) As String
        If param.IndexOf(",") = -1 Then
            Throw New ArgumentException("The character attribute parameter must include a comma followed by a boolean value.")
        End If

        '-- extract the boolean trailer
        Dim parameterBool As String = Regex.Match(param, "[^,]+$").Value
        param = Regex.Match(param, "^[^,]+").Value

        Dim ca As CharAttrib
        Dim o As Object = ParseEnum(param, GetType(CharAttrib))
        If o Is Nothing Then
            Throw New ArgumentException("The character attribute '" & param & "' is not recognized.")
        Else
            ca = DirectCast(o, CharAttrib)
        End If
        Return Convert.ToChar(ca) & ExpandBool(parameterBool)
    End Function

    Private Shared Function ExpandColor(ByVal param As String) As String
        Dim c As Color
        Dim o As Object = ParseEnum(param, GetType(Color))
        If o Is Nothing Then
            Throw New ArgumentException("The color '" & param & "' is not recognized.")
        Else
            c = DirectCast(o, Color)
        End If
        Return Convert.ToChar(c)
    End Function

    Private Shared Function ExpandFont(ByVal param As String) As String
        Dim f As Font
        Dim o As Object = ParseEnum(param, GetType(Font))
        If o Is Nothing Then
            Throw New ArgumentException("The font '" & param & "' is not recognized.")
        Else
            f = DirectCast(o, Font)
        End If
        Return Convert.ToChar(f)
    End Function

    Private Shared Function ExpandBool(ByVal param As String) As String
        If Not Regex.IsMatch(param, "^on|^1|^true|^yes|^off|^0|^false|^no", RegexOptions.IgnoreCase) Then
            Throw New ArgumentException("The parameter '" & param & "' should be a boolean.")
        End If

        If Regex.IsMatch(param, "^on|^1|^true|^yes", RegexOptions.IgnoreCase) Then
            Return "1"
        Else
            Return "0"
        End If
    End Function

    Private Shared Function ExpandDateAttrib(ByVal param As String) As String
        Select Case param.ToUpper
            Case "MM/DD/YY"
                Return "0"
            Case "DD/MM/YY"
                Return "1"
            Case "MM-DD-YY"
                Return "2"
            Case "DD-MM-YY"
                Return "3"
            Case "MM.DD.YY"
                Return "4"
            Case "DD.MM.YY"
                Return "5"
            Case "MM DD YY"
                Return "6"
            Case "DD MM YY"
                Return "7"
            Case "MMM.DD, YYYY", "MMM.DD,YYYY"
                Return "8"
            Case "DDD"
                Return "9"
            Case Else
                Throw New ArgumentException("The date format '" & param & "' is not recognized.")
        End Select
    End Function

    Private Shared Function ExpandExtendedChar(ByVal param As String) As String
        Dim ec As ExtChar
        Dim o As Object = ParseEnum(param, GetType(ExtChar))
        If o Is Nothing Then
            Throw New ArgumentException("The extended character '" & param & "' is not recognized.")
        Else
            ec = DirectCast(o, ExtChar)
        End If
        Return Convert.ToChar(ec)
    End Function

    Private Shared Function ExpandFileLabel(ByVal param As String) As Char
        '-- let the filelabel do our validation for us
        Dim fl As New File(param)
        Return fl.Label
    End Function

    Private Shared Function ExpandControlChar(ByVal c As ControlChars, ByVal param As String) As String

        Dim s As String = Convert.ToChar(c)

        Select Case c
            Case ControlChars.CallDate
                s &= ExpandDateAttrib(param)
            Case ControlChars.CallPic
                s &= ExpandFileLabel(param)
            Case ControlChars.CallString
                s &= ExpandFileLabel(param)
            Case ControlChars.CallTime
            Case ControlChars.CharAttrib
                s &= ExpandCharAttrib(param)
            Case ControlChars.CharSpacing
            Case ControlChars.Color
                s &= ExpandColor(param)
            Case ControlChars.ExtChar
                s &= ExpandExtendedChar(param)
            Case ControlChars.Flash
                s &= ExpandBool(param)
            Case ControlChars.Font
                s &= ExpandFont(param)
            Case ControlChars.NewLine
            Case ControlChars.Speed1
            Case ControlChars.Speed2
            Case ControlChars.Speed3
            Case ControlChars.Speed4
            Case ControlChars.Speed5
            Case ControlChars.WideOff
            Case ControlChars.WideOn
            Case Else
        End Select

        Return s
    End Function

    Private Shared Function MapCommandToEnum(ByVal CommandCode As String) As [Enum]
        Dim o As Object = ParseEnum(CommandCode, GetType(ControlChars))
        If Not o Is Nothing Then
            Return DirectCast(o, ControlChars)
        End If
        Throw New ArgumentException("The command '" & CommandCode & "' is not recognized.")
    End Function


    Private Shared Function ParseEnum(ByVal s As String, ByVal t As Type) As Object
        If s = "" Then
            Return Nothing
        End If
        Dim o As Object = Nothing
        Try
            o = System.Enum.Parse(t, s, True)
        Catch ex As System.ArgumentException
            '-- if the string representation provided doesn't match 
            '-- any known enum case, we'll get this exception
        End Try
        Return o
    End Function

#End Region

    ''' <summary>
    ''' pretty-printer for showing ASCII bytestreams in the console
    ''' </summary>
    Private Shared Function BytesToString(ByVal ba() As Byte) As String
        Dim sb As New StringBuilder
        For Each b As Byte In ba
            If Convert.ToInt32(b) < 32 Then
                If [Enum].IsDefined(GetType(Ascii), b) Then
                    sb.Append("<" & CType(b, Ascii).ToString & ">")
                    If b = Ascii.CR Then
                        sb.Append(Environment.NewLine)
                    End If
                Else
                    sb.Append(String.Format("<{0:x2}>", b))
                End If
            Else
                sb.Append(Convert.ToChar(b))
            End If
        Next
        Return sb.ToString
    End Function

    ''' <summary>
    ''' converts a time to hex code between 00-8F (10 minute intervals)
    ''' 00 equals 12:00-12:10am, 01 equals 12:10-12:20am, etc.
    ''' </summary>
    ''' <remarks>
    ''' Stop Time is ignored when Start Time is set to Always (FF)
    ''' </remarks>
    Private Shared Function TimeToByteAscii(ByVal dt As DateTime) As String
        If dt = DateTime.MinValue Then
            Return TimeToByteAscii(TimeConstant.StartOfDay)
        End If
        If dt = DateTime.MaxValue Then
            Return TimeToByteAscii(TimeConstant.Always)
        End If
        Dim i As Integer = Convert.ToInt32(Math.Floor(dt.TimeOfDay.TotalMinutes / 10))
        Return String.Format("{0:x2}", i)
    End Function

    Private Shared Function TimeToByteAscii(ByVal t As TimeConstant) As String
        Return String.Format("{0:x2}", Convert.ToByte(t))
    End Function

    Public Enum TimeConstant As Byte
        StartOfDay = &H0
        EndOfDay = &H8F
        AllDay = &HFD
        Never = &HFE
        Always = &HFF
    End Enum

End Class
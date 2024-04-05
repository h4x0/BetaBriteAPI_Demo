Imports BetaBrite.RS232
Imports BetaBrite.Protocol
Imports BetaBrite.Sign

Module Module1

    Const ComPort As Integer = 1

    Dim _bb As New BetaBrite.Sign(ComPort)

    Sub Main()
        '-- this causes Debug.Write to appear in console
        Dim t As New TextWriterTraceListener(Console.Out)
        Trace.Listeners.Add(t)

        _bb.Open()

        ' ** UNCOMMENT LINES BELOW TO TEST THOSE FEATURES **

        'TestHelloWorld()

        '_bb.Reset()
        '_bb.ClearMemory()

        'TestEnums()
        'TestTimeSet()
        'TestTransitions()
        'TestSpecials()
        'TestString()
        'TestRunSequence()
        'TestPicture()
        TestAnimation()
        'TestInternational()

        _bb.Close()
        Console.WriteLine("Press ENTER to continue..")
        Console.ReadLine()
    End Sub

    ''' <summary>
    ''' high ASCII characters are supported but must be internally mapped to BetaBrite's
    ''' proprietary double-byte representation
    ''' </summary>
    Sub TestInternational()
        _bb.Display("decisión <color=yellow>idénticas <color=green>más <color=amber>enseñanza")
    End Sub

    ''' <summary>
    ''' The betabrite doesn't support a real animation mode, however, you can simulate animation
    ''' by using the &lt;NoHold&gt; or &gt;speed5&lt; commands and displaying several pictures 
    ''' in sequence
    ''' </summary>
    Sub TestAnimation()

        With _bb
            .UseMemoryText("A"c, 256)
            .UseMemoryPicture("D"c)
            .UseMemoryPicture("E"c)
            .UseMemoryPicture("F"c)
            .UseMemoryPicture("G"c)
            .UseMemoryPicture("H"c)
            .UseMemoryPicture("I"c)
            .UseMemoryPicture("J"c)
            .UseMemoryPicture("K"c)
            .AllocateMemory()

            .SetPicture("D"c, "betabrite_picture_anim1.bmp")
            .SetPicture("E"c, "betabrite_picture_anim2.bmp")
            .SetPicture("F"c, "betabrite_picture_anim3.bmp")
            .SetPicture("G"c, "betabrite_picture_anim4.bmp")
            .SetPicture("H"c, "betabrite_picture_anim5.bmp")
            .SetPicture("I"c, "betabrite_picture_anim6.bmp")
            .SetPicture("J"c, "betabrite_picture_anim7.bmp")
            .SetPicture("K"c, "betabrite_picture_anim8.bmp")

            .SetText("A"c, "<nohold><callpic=D><newline><callpic=E><newline><callpic=F><newline><callpic=G><newline><callpic=H><newline><callpic=I><newline><callpic=J><newline><callpic=K><newline>", Transition.Hold)
        End With
    End Sub

    ''' <summary>
    ''' tests loading and displaying pictures of various sizes
    ''' see file templates in the .bin folder to create your own
    ''' </summary>
    Sub TestPicture()
        With _bb
            .UseMemoryText("A"c, 256)
            .UseMemoryPicture("B"c, 10, 7)
            .UseMemoryPicture("C"c, 35, 7)
            .UseMemoryPicture("D"c)
            .AllocateMemory()

            .SetPicture("B"c, "betabrite_picture_triangle.bmp")
            .SetPicture("C"c, "betabrite_picture_smiley.bmp")
            .SetPicture("D"c, "betabrite_picture_demo.bmp")

            .SetText("A"c, "<callpic=C><callpic=C><callpic=B><callpic=C><callpic=C><newline><callpic=D>", Transition.RollUp)
        End With
    End Sub

    ''' <summary>
    ''' tests run sequences, so multiple textfiles will display in a particular order
    ''' </summary>
    Sub TestRunSequence()
        With _bb
            .UseMemoryText("D"c, 256)
            .UseMemoryText("E"c, 256)
            .UseMemoryText("F"c, 256)
            .UseMemoryText("G"c, 256)
            .AllocateMemory()

            .SetText("D"c, "File D!")
            .SetText("E"c, "File E@")
            .SetText("F"c, "File F#")
            .SetText("G"c, "File G$")

            .SetRunSequence("GDFE")
        End With
    End Sub

    Sub Demo1()
        Dim bb As New BetaBrite.Sign(1)
        With bb
            .Open()
            .UseMemoryText("D"c, 256)
            .UseMemoryText("E"c, 256)
            .UseMemoryText("F"c, 256)
            .AllocateMemory()

            .SetText("D"c, _
                "<font=five><color=green>This is <font=seven>file D", _
                Transition.Rotate)
            .SetText("E"c, _
                "<font=five><color=yellow>This is <font=seven>file E", _
                Transition.WipeLeft)
            .SetText("F"c, _
                "<font=five><color=red>time is <calltime>", _
                Transition.RollDown)

            .SetRunSequence("GDFE")
            .Close()
        End With
    End Sub

    Sub TestHelloWorld()
        _bb.Display("I <extchar=heart> BetaBrite!")
        '_bb.SetSingleMessage("Hello World!")
    End Sub

    Sub TestTimeSet()
        _bb.SetDateAndTime()
        _bb.Display("Time<newline><calltime><newline>Date<newline><calldate=mmm.dd,yyyy><newline>")
    End Sub

    ''' <summary>
    ''' strings can be updated without making the sign "flash", 
    ''' which is the major reason to use them.. eg, they are like variables
    ''' note that strings only support a subset of the text formatting codes
    ''' </summary>
    Sub TestString()
        TestStringPrivate()
        Console.WriteLine("(note that the display will NOT flash when this happens!)")
        Console.WriteLine("Press ENTER to update the string..")
        Console.ReadLine()
        TestStringPrivate("678")
    End Sub
    Sub TestStringPrivate(Optional ByVal UpdateValue As String = "")
        If UpdateValue = "" Then
            _bb.UseMemoryText("A"c, 1024)
            _bb.UseMemoryString("B"c, 32)
            _bb.AllocateMemory()
            _bb.SetText("A"c, "Count is <callstring=B>")
            UpdateValue = "364"
        End If
        '-- possibly update existing string?
        _bb.SetString("B"c, UpdateValue)
    End Sub

    ''' <summary>
    ''' tests the special animation modes supported by the BetaBrite
    ''' some of these work fine as standard transitions (eg they will
    ''' display the message text), but some won't, so be careful
    ''' </summary>
    Sub TestSpecials()
        With _bb
            .UseMemoryText("A"c, 1024)
            .AllocateMemory()

            .SetTextMultiple("Balloon", Transition.Special, Special.Balloon)
            .SetTextMultiple("CherryBomb", Transition.Special, Special.CherryBomb)
            .SetTextMultiple("CycleColors", Transition.Special, Special.CycleColors)
            .SetTextMultiple("DontDrink", Transition.Special, Special.DontDrink)
            .SetTextMultiple("Fireworks", Transition.Special, Special.Fireworks)
            .SetTextMultiple("Fish", Transition.Special, Special.Fish)
            .SetTextMultiple("Interlock", Transition.Special, Special.Interlock)
            .SetTextMultiple("NewsFlash", Transition.Special, Special.NewsFlash)
            .SetTextMultiple("NoSmoking", Transition.Special, Special.NoSmoking)
            .SetTextMultiple("SlotMachine", Transition.Special, Special.SlotMachine)
            .SetTextMultiple("Snow", Transition.Special, Special.Snow)
            .SetTextMultiple("Sparkle", Transition.Special, Special.Sparkle)
            .SetTextMultiple("Spray", Transition.Special, Special.Spray)
            .SetTextMultiple("Starburst", Transition.Special, Special.Starburst)
            .SetTextMultiple("Switch", Transition.Special, Special.Switch)
            .SetTextMultiple("ThankYou", Transition.Special, Special.ThankYou)
            .SetTextMultiple("Twinkle", Transition.Special, Special.Twinkle)
            .SetTextMultiple("Welcome", Transition.Special, Special.Welcome)
            .SetTextMultiple("A"c)
        End With

    End Sub

    ''' <summary>
    ''' tests the various text transitions available 
    ''' </summary>
    Sub TestTransitions()
        With _bb
            .UseMemoryText("A"c, 1024)
            .AllocateMemory()

            .SetTextMultiple("Compressed Rotate", Transition.CompressedRotate)
            .SetTextMultiple("Flash", Transition.Flash)
            .SetTextMultiple("Hold", Transition.Hold)
            .SetTextMultiple("RollDown", Transition.RollDown)
            .SetTextMultiple("RollIn", Transition.RollIn)
            .SetTextMultiple("RollLeft", Transition.RollLeft)
            .SetTextMultiple("RollOut", Transition.RollOut)
            .SetTextMultiple("RollRight", Transition.RollRight)
            .SetTextMultiple("RollUp", Transition.RollUp)
            .SetTextMultiple("Rotate", Transition.Rotate)
            .SetTextMultiple("Scroll", Transition.Scroll)
            .SetTextMultiple("WipeDown", Transition.WipeDown)
            .SetTextMultiple("WipeIn", Transition.WipeIn)
            .SetTextMultiple("WipeLeft", Transition.WipeLeft)
            .SetTextMultiple("WipeOut", Transition.WipeOut)
            .SetTextMultiple("WipeRight", Transition.WipeRight)
            .SetTextMultiple("WipeUp", Transition.WipeUp)
            .SetTextMultiple("A"c)
        End With
    End Sub

    ''' <summary>
    ''' tests all the HTML-style tags supported in message text to specify font, 
    ''' color, and many other things. These are implemented using Enum lookups.
    ''' </summary>
    Sub TestEnums()
        '-- set to bland, basic settings
        Dim msg As String = "<color=red><font=seven>"

        '-- character attributes (only two work on the BetaBrite)
        msg &= _
            "<charattrib=wide,on>wide <charattrib=wide,off>off<newline>" & _
            "<charattrib=doublewide,on>dwd <charattrib=doublewide,off>off<newline>"

        '-- time (only one format)
        msg &= _
            "time: <calltime><newline>"

        '-- date formatting (multiple formats available)
        msg &= _
            "Today's date<newline>" & _
            "<calldate=mm/dd/yy><newline>" & _
            "<calldate=dd/mm/yy><newline>" & _
            "<calldate=mm-dd-yy><newline>" & _
            "<calldate=dd-mm-yy><newline>" & _
            "<calldate=mm.dd.yy><newline>" & _
            "<calldate=dd.mm.yy><newline>" & _
            "<calldate=mm dd yy><newline>" & _
            "<calldate=dd mm yy><newline>" & _
            "<calldate=mmm.dd, yyyy><newline>" & _
            "<calldate=ddd><newline>"

        '-- extended character set
        msg &= "<extchar=uparrow><extchar=downarrow><extchar=leftarrow><extchar=rightarrow>" & _
            "<extchar=pacman><extchar=sailboat><extchar=baseball><extchar=telephone>" & _
            "<extchar=heart><extchar=car><extchar=handicap><extchar=rhino>" & _
            "<extchar=mug><extchar=satellite><extchar=copyright><newline>"

        '-- colors
        msg &= _
            "<color=red>Red<newline><color=green>Green<newline><color=amber>Amber<newline>" & _
            "<color=dimred>DimRed<newline><color=dimgreen>DimGreen<newline><color=brown>Brown<newline>" & _
            "<color=brown>Brown<newline><color=orange>Orange<newline><color=yellow>Yellow<newline>" & _
            "<color=rainbow1>Rainbow1<newline><color=rainbow2>Rainbow2<newline>" & _
            "<color=auto>Automatic<newline><color=red>"


        '-- font demos (5 pixel height)
        msg &= _
            "<font=five>Five<newline>" & _
            "<font=fivebold>FiveBold<newline>" & _
            "<font=fivewide>FiveWide<newline>" & _
            "<font=fivewidebold>FiveWideBold<newline>"

        '-- font demos (7 pixel height)
        msg &= _
            "<font=seven>Seven<newline>" & _
            "<font=sevenserif>SevenSerif<newline>" & _
            "<font=sevenbold>SevenBold<newline>" & _
            "<font=sevenboldserif>SevenBoldSerif<newline>" & _
            "<font=sevenshadow>SevenShadow<newline>" & _
            "<font=sevenshadowserif>SevenShadowSerif<newline>" & _
            "<font=sevenwide>SevenWide<newline>" & _
            "<font=sevenwideserif>SevenWideSerif<newline>" & _
            "<font=sevenwidebold>SevenWideBold<newline>" & _
            "<font=sevenwideboldserif>SevenWideBoldSerif<newline>" & _
            "<font=seven><newline>"

        '-- control codes
        msg &= _
            "Line One<newline>Line Deux<newline>Line Tres" & _
            "<flash=1>Flash <flash=0>NoFlash<newline>" & _
            "<wideon>wide<wideoff>off<newline>" & _
            "<speed1>speed1<newline><speed2>speed2<newline>" & _
            "<speed3>speed3<newline><speed4>speed4<newline>" & _
            "<speed5>speed5<newline>"

        _bb.Display(msg)
    End Sub

End Module
option explicit
dim objShell, objFile
set objShell = CreateObject("wscript.shell")
set objFile = CreateObject("Scripting.FileSystemObject")
dim directory, n, readData, readElement, i, ii, iii, config, countdown, think, delay, das, dropThink, holdThink, sequence, lastDirection, natural

function funcRand(fMin, fMax)
    funcRand = int((fMax-fMin+1)*rnd+fMin)
end function

sub initialize()
    directory = objShell.CurrentDirectory
    objFile.CreateTextFile(directory & "\terminate.txt", true).WriteLine("0")

    objFile.CreateTextFile(directory & "\terminate.vbs", true).WriteLine( _
        "set objShell = CreateObject(""wscript.shell"")" & vbCrlf & _
        "set objFile = CreateObject(""Scripting.FileSystemObject"")" & vbCrlf & _
        "directory = objShell.CurrentDirectory" & vbCrlf & _
        "MsgBox(""Terminate bot"")" & vbCrlf & _
        "objFile.CreateTextFile(directory & ""\terminate.txt"", true).WriteLine(""1"")" & vbCrlf & _
        "call objShell.run(""taskkill /f /im wscript.exe"",,True)" )

    objShell.run("""" & directory & "\terminate.vbs" & """")
    body()
end sub

sub body()
    n = MsgBox("Start", 0+64)

    if objFile.FileExists(directory & "\config.txt") then
    else
        objFile.CreateTextFile(directory & "\config.txt", false).WriteLine("3000, 500, 200, 100, 0, 0")
    end if

    if objFile.FileExists(directory & "\sequence.txt") then
    else
        objFile.CreateTextFile(directory & "\sequence.txt", false).WriteLine()
    end if

    if objFile.FileExists(directory & "\terminate.txt") then
    else
        objFile.CreateTextFile(directory & "\terminate.txt", false).WriteLine("0")
    end if

    readData = objFile.OpenTextFile(directory & "\config.txt").ReadAll
    readData = replace(readData, vbCrlf, "")
    readData = replace(readData, " ", "")
    config = split(readData, ",")

    if ubound(config) < 6 then
        n = MsgBox("Invalid config parameters", 0+16)
        wscript.quit
    else
        if isNumeric(config(0)) and isNumeric(config(1)) and isNumeric(config(2)) and isNumeric(config(3)) and isNumeric(config(4)) and isNumeric(config(5)) and isNumeric(config(6)) then
        else
            n = MsgBox("Invalid config parameters", 0+16)
            wscript.quit
        end if
    end if

    countdown = config(0)
    think = config(1)
    delay = config(2)
    das = config(3)
    dropThink = config(4)
    holdThink = config(5)
    natural = config(6)

    if objFile.GetFile(directory & "\sequence.txt").size > 0 then
    else
        n = MsgBox("Sequence file is empty", 0+16)
        wscript.quit
    end if

    readData = objFile.OpenTextFile(directory & "\sequence.txt").ReadAll
    readData = replace(readData, vbCrlf, "")
    readData = replace(readData, " ", "")
    sequence = split(readData, ",")
    initiate()
end sub

sub initiate()
    wscript.sleep(countdown)
    lastDirection = "none"

    for i = 0 to ubound(sequence)
        readElement = sequence(i)

        if natural = 1 then
            if lastDirection = "l" and sequence(i) = "dasl" then
            elseif lastDirection = "r" and sequence(i) = "dasr" then
            else
                wscript.sleep(delay)
            end if
        else
            wscript.sleep(delay)
        end if

        if readElement = "left" then
            objShell.SendKeys("{LEFT}")
            lastDirection = "none"
        elseif readElement = "right" then
            objShell.SendKeys("{RIGHT}")
            lastDirection = "none"
        elseif readElement = "dasl" then
            objShell.SendKeys("{LEFT}")

            if natural = 1 and lastDirection = "l" then
            else
                wscript.sleep(das)
            end if

            for ii = 1 to 10
                objShell.SendKeys("{LEFT}")
            next

            lastDirection = "l"
        elseif readElement = "dasr" then
            objShell.SendKeys("{RIGHT}")
            
            if natural = 1 and lastDirection = "r" then
            else
                wscript.sleep(das)
            end if

            for ii = 1 to 10
                objShell.SendKeys("{RIGHT}")
            next

            lastDirection = "r"
        elseif readElement = "up" or readElement = "cw" then
            objShell.SendKeys("{UP}")
            lastDirection = "none"
        elseif readElement = "down" then
            for ii = 1 to 5
                for iii = 1 to 5
                    objShell.SendKeys("{DOWN}")
                next

                wscript.sleep(10)
            next
        elseif readElement = "space" or readElement = "drop" then
            objShell.SendKeys(" ")
            wscript.sleep(dropThink)
        elseif readElement = "c" or readElement = "hold" then
            objShell.SendKeys("c")
            wscript.sleep(holdThink)
        elseif readElement = "z" or readElement = "ccw" then
            objShell.SendKeys("z")
            lastDirection = "none"
        elseif readElement = "a" or readElement = "180" then
            objShell.SendKeys("a")
            lastDirection = "none"
        elseif readElement = "think" then
            wscript.sleep(think)
            lastDirection = "none"
        elseif readElement = "delay" then
            wscript.sleep(delay)
            lastDirection = "none"
        elseif readElement = "break" then
            n = MsgBox("Breakpoint reached", 0+48)
            body()
        end if

        readData = objFile.OpenTextFile(directory & "\terminate.txt").ReadAll

        if left(readData, 1) = 1 then
            n = MsgBox("Task has been terminated", 0+48)
            wscript.quit
        end if
    next

    body()
end sub

initialize()
wscript.echo("Done")
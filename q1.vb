Module Module1

    Sub Main()
        '2 D dater arrays'
        'pre defined average speed in each material'
        Dim average_speeds() As Double = {323.5, 180.8, 100.3}
        'pre defined materials'
        Dim material() As String = {"Wood    ", "Steel   ", "Titanium"}
        'user inputed depth of material'
        Dim Depth(2) As Double
        'Index'
        Dim i

        'asking the user for depths'
        For i = 0 To 2
            Console.WriteLine("Please enter the depth of " & material(i) & " you are testing on in meters")
            Depth(i) = Console.ReadLine
        Next

        'table titles'
        Console.WriteLine("/ Material /" & vbTab & "/ Depth (meters) /" & vbTab & "/ Average Speed (m/s) /" & vbTab & "/ Time Spent in the medium (s) /")

        'table rows'
        For i = 0 To 2
            Console.WriteLine(material(i) & vbTab & Space(3) & Format(Depth(i), "0.000") & vbTab & Space(5) & Format(average_speeds(i), "0.000") & vbTab & Space(11) & Format(Depth(i) / average_speeds(i), "0.000"))
        Next

        console.writeline("the average of the three times: ")

    End Sub

End Module

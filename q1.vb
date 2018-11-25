Module Module1
    Function Table_Headings() 'no inputs required'
        'making the headings of the table'
        Console.WriteLine("Material /" & vbTab & "Time Taken (s)")
        Console.WriteLine("---------------------------------")
    End Function

    Function Distance_In(Wood As Double, Steel As Double, Titanium As Double) 'assighning our array of depths to internal variables within the function'
        Dim Output
        'making a variable within the function = what we want it to output, so we state what we want our internal function variables to = and then get them to = the format we want to output'
        Output = "For distance in wood you entered: " & Wood & vbNewLine & "For distance in steel you entered: " & Steel & vbNewLine & "For distance in titanium you entered: " & Titanium
        Console.WriteLine(Output)
    End Function

    Function average_time_to_reach_targer(Wood As Integer, Steel As Integer, Titanium As Integer, WoodD As Integer, SteelD As Integer, TitaniumD As Integer) 'Wood being the average speed the bullet travels through wood and woodD being the distance the bullet was fired from the wood'
        Console.WriteLine("On average it took " & ((WoodD / Wood) + (SteelD / Steel) + (SteelD / Titanium)) / 3 & " seconds for the bullet to reach the targer")
    End Function

    Sub Main()

        'Pre given values placed in acording arrays'
        Dim average_speeds() As Double = {323.5, 180.8, 100.3}
        Dim Distance_from_material() As Integer = {800, 500, 250}
        Dim material() As String = {"Wood    ", "Steel   ", "Titanium"}
        'A variable for the user to input a depth for ecah material'
        Dim Depth(2)
        'a usfull index for out for loops to work'
        Dim i, i1


        'asking the user to input a depth and then assighning that value to a possition in an array'
        For i = 0 To 2
            Console.WriteLine("Please enter the depth of " & Trim(material(i)) & " you are testing on in meters")
            Depth(i) = Console.ReadLine
            'making sure the assighned value is a positive integer (not including 0 as the average speed infers it spends time in the medium)'
            For i1 = 0 To 2
                If Depth(i) <= 0 Then
                    'If it isnt, asking the user to input a value in again'
                    Console.WriteLine("please enter a possitive double for the depth of " & Trim(material(i)))
                    Depth(i) = Console.ReadLine()
                End If

            Next
        Next

        Console.WriteLine(vbNewLine)

        Distance_In(Depth(0), Depth(1), Depth(2))

        Console.WriteLine(vbNewLine)

        average_time_to_reach_targer(average_speeds(0), average_speeds(1), average_speeds(2), Distance_from_material(0), Distance_from_material(1), Distance_from_material(2))

        Console.WriteLine(vbNewLine)

        'outputing the table of results'
        Table_Headings()

        'outputing the rows'
        For i = 0 To 2
            Console.WriteLine(material(i) & " /" & vbTab & Format(Depth(i) / average_speeds(i), "0.000"))
        Next

    End Sub
End Module

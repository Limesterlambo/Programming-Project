'07/12/18.	Module Module1
'Nicholas Letchford-Jones.	
'This program is designed to be helpful to the scientists who will be using it, asking them to input depths, checking whether the are valid field results, aka above 0; and then going throw a secondary verification process of confirming with them the value they input. the program will then inform the scientists the period the bullet spent in the medium, and then the average times over the 3 materials the bullet spent in the medium.
Module Module1
    Dim t

    '  To seperate your work out in the command prompt.
    Function newlines(Number_of_newlines As Integer)
        For i = 1 To Number_of_newlines
            Console.Write(vbNewLine)
        Next
    End Function

    '  General up to 5 column table maker.
    Function Table_Heading(Number_of_columnns As Integer, Heading1 As String, Heading2 As String, Heading3 As String, Heading4 As String, Heading5 As String)
        Dim n
        n = Number_of_columnns
        If n > 0 Then
            Console.Write(Heading1 & vbTab & "|" & vbTab)
        Else
            Console.Write("")
        End If

        If n > 1 Then
            Console.Write(Heading2 & vbTab & "|" & vbTab)
        Else
            Console.Write("")
        End If

        If n > 2 Then
            Console.Write(Heading3 & vbTab & "|" & vbTab)
        Else
            Console.Write("")
        End If

        If n > 3 Then
            Console.Write(Heading4 & vbTab & "|" & vbTab)
        Else
            Console.Write("")
        End If

        If n > 4 Then
            Console.WriteLine(Heading5 & vbTab & "|" & vbTab)
        Else
            Console.WriteLine("")
        End If
        ' If the table devision is off, modify the number of times "-" string is duplicated.
        Console.WriteLine(StrDup((n * 17) + 7, "-"))
    End Function

    ' Time = distance/velocity rearranged from velocity = (distance/time), assuming SI units.
    Function Time_Taken(distance, speed) As Double
        Time_Taken = (distance / speed)
    End Function

    ' Average calculator.
    Function Average(Total, number_of_data_samples) As Double
        Average = Total / number_of_data_samples
    End Function

    Sub Main()
        Dim material = {"Wood    ", "Steel   ", "Titanium"}
        Dim Average_speed = {323.5, 180.8, 100.3}
        Dim depth(2) As Double
        Dim t, total_time As Double

        ' Q1 part c - fore user to input positive doubles. 
        For i = 0 To 2
            Console.WriteLine("Please enter the depth of " & Trim(material(i)) & " you are testing on in meters:")
            depth(i) = Console.ReadLine
            ' Making sure the assighned value is a positive integer (not including 0 as the average speed infers it spends time in the medium).
            While depth(i) <= 0
                ' If it isnt, asking the user to reinput a value
                Console.WriteLine("please enter a possitive double for the depth of " & Trim(material(i)))
                    depth(i) = Console.ReadLine()
            End While
        Next

        newlines(1)

        ' Q1 part a.
        ' -Part 1 - confirming users inputed depths with them.
        For i = 0 To 2
            Console.WriteLine("For distance in " & Trim(material(i)) & " you entered: " & depth(i) & " metres")
        Next

        newlines(1)
        ' -Part 2 - forming a table for the time taken by the bullet to reach the target.

        ' --Designing  the table columns.
        Table_Heading(2, "Material", "Time Taken (s)", "", "", "")

        ' --Designing  the table rows.
        For i = 0 To 2
            Console.Write(material(i) & vbTab & "|" & vbTab)
            t = Time_Taken(depth(i), Average_speed(i))
            Console.WriteLine(Format(t, "0.000") & vbTab & vbTab & "|")
            total_time = total_time + t
        Next

        newlines(1)

        ' Q1 part b.
        ' -Average calculation.
        Console.WriteLine("""The time taken by the bullet to reach the target on average"" was " & Format(Average(total_time, 3), "0.0000") & " seconds")
    End Sub

End Module

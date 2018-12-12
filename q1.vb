' 07/12/18.	
' Main contibutor: Nicholas Letchford-Jones.	
' This program is designed to be helpful to the experimental team, asking them to input depths of the materirial they are testing, checking whether they are valid field results, 
' aka above 0; The progam then confirmins with them the value they input. The program will Then inform the scientists the period the bullet 
' spent in the medium. Finally the program prints out the average times the bullet spent in the medium over the 3 materials.
Module Module1
    Dim t

    '  To seperate your work out in the command prompt.
    Function Newlines(Number_of_newlines As Integer)
        For i = 1 To Number_of_newlines
            Console.Write(vbNewLine)
        Next
    End Function

    '  Fuction for generating table column heading for up to 5 columns
    Function Table_Heading(Number_of_columnns As Integer, Heading0 As String, Heading1 As String, Heading2 As String, Heading3 As String, Heading4 As String)
        Dim Headings(4) As String
        Headings(0) = Heading0
        Headings(1) = Heading1
        Headings(2) = Heading2
        Headings(3) = Heading3
        Headings(4) = Heading4

        For I = 0 To 4
            If Number_of_columnns > I Then
                Console.Write(Headings(I) & vbTab & "|" & vbTab)
            Else
                Console.Write("")
            End If
        Next

        ' If the table devision is off, modify the number of times "-" string is duplicated.
        Console.WriteLine(vbNewLine & StrDup((Number_of_columnns * 17) + 7, "-"))
    End Function

    ' Time = distance/velocity (rearranged from velocity = (distance/time), assuming SI units.)
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
        Dim depth(2)
        Dim t, total_time As Double
        Dim j

        ' Q1 part c - fourcing the user to input positive doubles. 
        For i = 0 To 2
            Console.WriteLine("Please enter the depth of " & LCase(Trim(material(i))) & " you are testing on in meters:")
            depth(i) = Console.ReadLine

            ' Loop Making sure the assighned value is a number
            Int32.TryParse(depth(i), j)
            While j = False
                Console.WriteLine("ERROR: depth must be a numerical! " & "Please re-enter the depth of " & LCase(Trim(material(i))) & " you are testing on in meters:")
                depth(i) = Console.ReadLine()
                Int32.TryParse(depth(i), j)
            End While
            depth(i) = CDbl(depth(i))

            ' Making sure the value is positive. (not including 0 as the average speed infers it spends time in the medium).
            While depth(i) <= 0
                ' If it isnt positve ie. depth<=0 , asking the user to reinput a value
                Console.WriteLine("ERROR: depth must be a positive value! " & "Please re-enter the depth of " & LCase(Trim(material(i))) & " you are testing on in meters:")
                depth(i) = Console.ReadLine()
            End While

        Next

        Newlines(1)

        ' Q1 part a.
        ' Part 1 - confirming users inputed depths with them.
        For i = 0 To 2
            Console.WriteLine("For distance in " & LCase(Trim(material(i))) & " you entered: " & depth(i) & " metres")
        Next

        Newlines(1)
        ' Part 2 - forming a table for the time taken by the bullet to reach the target.

        ' Designing  the table columns.  
        Console.Write(Space(2))
        Table_Heading(2, "Material", "Time taken (s)", "", "", "")

        ' Designing  the table rows.
        For i = 0 To 2
            Console.Write(Space(2) & material(i) & vbTab & "|" & vbTab)
            t = Time_Taken(depth(i), Average_speed(i))
            Console.WriteLine(Format(t, "0.000") & vbTab & vbTab & "|")
            total_time = total_time + t
        Next

        Newlines(1)

        ' Q1 part b.
        ' Average calculation.
        Console.WriteLine("The time taken by the bullet to reach the target on average was " & Format(Average(total_time, 3), "0.0000") & " seconds")
    End Sub
End Module

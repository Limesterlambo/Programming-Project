Module Module1
    Dim t

    '  Function To seperate your work out in the command prompt.
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

        ' If the table devision is off, modify the number of times "-" string is duplicated.
        Console.WriteLine(StrDup((Number_of_columnns * 17) + 7, "="))
        For I = 0 To 4
            If Number_of_columnns > I Then
                Console.Write(Space(2) & Headings(I) & vbTab & "|")
            Else
                Console.Write("")
            End If
        Next
        Console.WriteLine(vbNewLine & StrDup((Number_of_columnns * 17) + 7, "="))
    End Function

    ' Time = distance/velocity (rearranged from velocity = (distance/time), assuming SI units.)
    Function Calculate_Time_Taken(distance, speed) As Double
        Calculate_Time_Taken = (distance / speed)
    End Function

    ' Average calculator.
    Function Average(Total, number_of_data_samples) As Double
        Average = Total / number_of_data_samples
    End Function

    Sub Main()
        Dim material = {"Wood    ", "Steel   ", "Titanium"}
        Dim Average_speed = {323.5, 180.8, 100.3}
        Dim Average_Time_taken(5) As Double

        Dim depth(2)
        Dim j
        Dim Total_time_Mod_Combination As Double

        'Array to contain, Sorted from largest to smalllest, average time taken(s) of materials in each combination.
        Dim Sorted_Average_Time_Taken(5) As Double

        ' Q1 part c - fourcing the user to input positive doubles. 
        For i = 0 To 2
            Console.WriteLine("Please enter the depth of " & LCase(Trim(material(i))) & " you are testing on in meters:")
            depth(i) = Console.ReadLine

            ' Loop Making sure the assighned value is a number
            Int32.TryParse(depth(i), j)
            While j = False
                Console.WriteLine("ERROR: depth must be a numerical! " &
                    "Please re-enter the depth of " & LCase(Trim(material(i))) & " you are testing on in meters:")
                depth(i) = Console.ReadLine()
                Int32.TryParse(depth(i), j)
            End While
            depth(i) = CDbl(depth(i))

            ' Making sure the value is positive. (not including 0 as the average speed infers it spends time in the medium).
            While depth(i) <= 0
                ' If it isnt positve ie. depth<=0 , asking the user to reinput a value.
                Console.WriteLine("ERROR: depth must be a positive value! " &
                    "Please re-enter the depth of " & LCase(Trim(material(i))) & " you are testing on in meters:")
                depth(i) = Console.ReadLine()
            End While

        Next

        Newlines(1)

        ' Q1 part a.
        ' Part 1 - confirming users inputed depths with them.
        For i = 0 To 2
            Console.WriteLine("For distance in " & LCase(Trim(material(i))) & " you entered: " & depth(i) & " metres")
        Next


        ' Createing an arry containg all possible combination titles that will be needed for the table.
        Dim Combination_Titles() As String = {"Blanks and brown gunpowder", "Lead bullets and brown gunpowder", "Armor-piercing bullets and gunpowder",
            "Blanks and sulphur-free gunpowder", "Lead bullets and sulphur-free gunpowder", "Armor-piercing and sulphur-free gunpowder"}

        ' Creating three arrays, one for each material, to store the modified setup's average speeds. This array will be be used to calulate the times taken.
        Dim Mod_Average_Speed_Wood(2, 1), Mod_Average_Speed_Steel(2, 1), Mod_Average_Speed_Titanium(2, 1) As Double

        'should we have one array for time take or multiple again? how does the array sort function work?
        Dim Time_Taken_Wood(6), Time_Taken_Steel(6), Time_Taken_Titanium(6) As Double

        Dim l As Integer




        ' Array Bullet() elements correspond To 0 blank, 1 LeadBullets, 2 Armour-piercing Bullets.
        ' Array Gunpowder() elements Correspond to 0 Brown Gunpowder, 1 Sulphur-free Gunowder.
        Dim Wood_Bullets() As Integer = {-3, 20, -30}
        Dim Wood_Gunpowder() As Integer = {35, -10}
        Dim Steel_Bullets() As Integer = {-5, 12, 5}
        Dim Steel_Gunpowder() As Integer = {35, -10}
        Dim Titanium_Bullets() As Integer = {-12, 5, 30}
        Dim Titanium_Gunpowder() As Integer = {35, -10}

        ' Loop with nested Loop used to enter the variables for each combination of setup into the calculation of a new average time to be assigned into the array.
        For X = 0 To 1
            For I = 0 To 2
                ' Calculates the new average time taken by the effect of modifications, Mod_Average_Speed Array's fist rack = Bullet used and Second rack= Gunpowder used.
                Mod_Average_Speed_Wood(I, X) = Average_speed(0) * (1 + (Wood_Bullets(I) + Wood_Gunpowder(X)) / 100)
                Mod_Average_Speed_Steel(I, X) = Average_speed(1) * (1 + (Steel_Bullets(I) + Steel_Gunpowder(X)) / 100)
                Mod_Average_Speed_Titanium(I, X) = Average_speed(2) * (1 + (Titanium_Bullets(I) + Titanium_Gunpowder(X)) / 100)
            Next
        Next

        l = 0

        For X = 0 To 1
            For I = 0 To 2
                Time_Taken_Wood(l) = Calculate_Time_Taken(depth(0), Mod_Average_Speed_Wood(I, X))
                Time_Taken_Steel(l) = Calculate_Time_Taken(depth(1), Mod_Average_Speed_Steel(I, X))
                Time_Taken_Titanium(l) = Calculate_Time_Taken(depth(2), Mod_Average_Speed_Titanium(I, X))
                l = l + 1
            Next
        Next

        Newlines(2)
        'producing the tables: 6 tables, Overall Title, Material,(s)Timetaken(s)
        For I = 0 To 5

            Console.WriteLine(Space(1) & UCase(Combination_Titles(I)) & ":")

            Table_Heading(2, "Material", "Time taken (s)", "", "", "")  'Console.Write(Space(2)) at front 
            Console.WriteLine(Space(2) & material(0) & vbTab &
                              "|" & Space(2) & Format(Time_Taken_Wood(I), "0.000") & vbTab & vbTab & "|")
            Console.WriteLine(Space(2) & material(1) & vbTab &
                              "|" & Space(2) & Format(Time_Taken_Steel(I), "0.000") & vbTab & vbTab & "|")
            Console.WriteLine(Space(2) & material(2) & vbTab &
                              "|" & Space(2) & Format(Time_Taken_Titanium(I), "0.000") & vbTab & vbTab & "|")
            Console.WriteLine(StrDup((2 * 17) + 7, "="))

            Total_time_Mod_Combination = Time_Taken_Wood(I) + Time_Taken_Steel(I) + Time_Taken_Titanium(I)
            Average_Time_taken(I) = Average(Total_time_Mod_Combination, 3)
            Console.WriteLine(Space(1) & "Average time taken: " & Format(Average_Time_taken(I), "0.000") & " seconds")

            Newlines(2)
        Next


        Console.WriteLine()
        Array.Copy(Average_Time_taken, Sorted_Average_Time_Taken, 5)
        Array.Sort(Sorted_Average_Time_Taken)

        For I = 0 To 5
            For P = 0 To 5
                If Sorted_Average_Time_Taken(I) = Average_Time_taken(P) Then
                    Console.WriteLine(Space(2) & Combination_Titles(P) & vbTab &
                              "|" & Space(2) & Format(Sorted_Average_Time_Taken(I), "0.000") & vbTab & vbTab & "|")
                End If
            Next
        Next
        Newlines(1)





    End Sub
End Module
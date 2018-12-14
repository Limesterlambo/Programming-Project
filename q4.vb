Module Module1
    'This program has been designed to allow the scientist to see how the different types of gunpowder affect the flight of the bullet being fired.
    'K E Hutton, December 2018

    Dim distance
    Dim angle(999) As Double
    Dim Velocity
    Dim count, i, total_h
    Dim time(999)
    Dim height(999)
    Dim average_height
    Dim total_time
    Dim average_time
    Dim percentage_of_bullets
    Dim height_of_target
    Dim Velocity_material_specific = {323.5, 180.8, 100.3}

    'This function allows the user to see how random angles between the given minimum and maximum affect the time, accuracy and percentage.
    'They can input their required distance, velocity, material And gunpowder type for which this function will perform the calculations.
    Function Probability_of_Hits(angle_of_shot As Integer, height_of_target As Integer, distance As Integer, velocity As Double, gunpowder As String, Material As String)

        For i = 0 To 999
            angle(i) = Rnd() * angle_of_shot

            angle(i) = angle(i) * Math.PI / 180
        Next

        For i = 0 To 999
            time(i) = distance / velocity * Math.Cos(angle(i))
            height(i) = distance * Math.Tan(angle(i))
        Next

        count = 0
        percentage_of_bullets = 0
        average_time = 0

        For i = 0 To 999
            If height(i) < height_of_target Then
                total_time = total_time + time(i)
                count = count + 1
            End If
        Next

        average_time = total_time / count
        percentage_of_bullets = count / 1000 * 100

        Console.WriteLine(total_time & " is the total time of all the bullets fired from the gun in seconds.")
        Console.WriteLine(count & " shots fired were below 100 meters out of the 1000 fired.")
        Console.WriteLine(average_time & " is the average time it took for the bullet to hit the target.")
        Console.WriteLine(percentage_of_bullets & "% of bullets that hit the target for the " & gunpowder & " gunpowder shooting at " & Material & ".")

    End Function

    'This function creates the table for the results.
    Function Table_Heading(Number_of_columnns As Integer, Heading0 As String, Heading1 As String, Heading2 As String, Heading3 As String, Heading4 As String, Heading5 As String, Heading6 As String, Heading7 As String)

        Dim Headings(7) As String

        Headings(0) = Heading0
        Headings(1) = Heading1
        Headings(2) = Heading2
        Headings(3) = Heading3
        Headings(4) = Heading4
        Headings(5) = Heading5
        Headings(6) = Heading6
        Headings(7) = Heading7

        For i = 0 To 7
            If Number_of_columnns > i Then
                Console.Write(Headings(i) & Space(3) & "|" & Space(3))
            End If
        Next


        'If the table division is off, modify the number of times "-" string is repeated. 
        Console.WriteLine(vbNewLine & StrDup((Number_of_columnns * 17) + 40, "-"))

    End Function

    'This function allows the user to input different materials, distances, gunpowders and velocities and will output the accuracy and average time taken.
    Function Accuracy_Table(angle_of_shot As Integer, height_of_target As Integer, distance As Integer, velocity As Double, material As String)

        For i = 0 To 999
            angle(i) = Rnd() * angle_of_shot

            angle(i) = angle(i) * Math.PI / 180
        Next

        For i = 0 To 999
            time(i) = distance / velocity * Math.Cos(angle(i))

            height(i) = distance * Math.Tan(angle(i))
        Next

        count = 0
        percentage_of_bullets = 0
        average_time = 0

        For i = 0 To 999
            If height(i) < height_of_target Then
                total_time = total_time + time(i)
                count = count + 1
            End If
        Next

        average_time = total_time / count
        percentage_of_bullets = count / 1000 * 100



        For i = 0 To 0
            Console.WriteLine(material & Space(12) & distance & Space(20) & percentage_of_bullets & Space(15) & time(i))
        Next

    End Function
    Sub Main()

        'Values inputted to calculate the total time, average and percentage for the brown gunpowder.
        Probability_of_Hits(15, 100, 800, 323.5, "Brown", "Wood")
        Console.ReadLine()
        Probability_of_Hits(15, 100, 500, 180.8, "Brown", "Steel")
        Console.ReadLine()
        Probability_of_Hits(15, 100, 250, 100.3, "Brown", "Titanium")
        Console.ReadLine()

        'Values inputted to calculate the total time, average and percentage for the brown gunpowder.
        Probability_of_Hits(10, 100, 800, 323.5, "Sulphur-free", "Wood")
        Console.ReadLine()
        Probability_of_Hits(10, 100, 500, 180.8, "Sulphur-free", "Steel")
        Console.ReadLine()
        Probability_of_Hits(10, 100, 250, 100.3, "Sulphur-free", "Titanium")
        Console.ReadLine()

        'Brown gunpowder shooting at wood with various distances.
        Console.WriteLine("Brown gunpowder being tested at various distances shooting at Wood.")
        Console.WriteLine("Material ¦ Distance (Meters) ¦ Acurracy (As Percentage) ¦ Time (Seconds)")
        Accuracy_Table(15, 100, 900, 323.5, "Wood")
        Accuracy_Table(15, 100, 500, 323.5, "Wood")
        Accuracy_Table(15, 100, 100, 323.5, "Wood")
        Console.ReadLine()

        'Brown gunpowder shooting at steel with various distances.
        Console.WriteLine("Brown gunpowder being tested at various distances shooting at Steel.")
        Console.WriteLine("Material ¦ Distance (Meters) ¦ Acurracy (As Percentage) ¦ Time (Seconds)")
        Accuracy_Table(15, 100, 900, 180.8, "Steel")
        Accuracy_Table(15, 100, 500, 180.8, "Steel")
        Accuracy_Table(15, 100, 100, 180.8, "Steel")
        Console.ReadLine()

        'Brown gunpowder shooting at titanium with various distances.
        Console.WriteLine("Brown gunpowder being tested at various distances shooting at Titanium.")
        Console.WriteLine("Material ¦ Distance (Meters) ¦ Acurracy (As Percentage) ¦ Time (Seconds)")
        Accuracy_Table(15, 100, 900, 100.3, "Titanium")
        Accuracy_Table(15, 100, 500, 100.3, "Titanium")
        Accuracy_Table(15, 100, 100, 100.3, "Titanium")
        Console.ReadLine()

        'Sulphur-free gunpowder shooting at wood with various distances.
        Console.WriteLine("Sulphur-free gunpowder being tested at various distances shooting at Wood.")
        Console.WriteLine("Material ¦ Distance (Meters) ¦ Acurracy (As Percentage) ¦ Time (Seconds)")
        Accuracy_Table(10, 100, 900, 323.5, "Wood")
        Accuracy_Table(10, 100, 500, 323.5, "Wood")
        Accuracy_Table(10, 100, 100, 323.5, "Wood")
        Console.ReadLine()

        'Sulphur-free gunpowder shooting at steel with various distances.
        Console.WriteLine("Sulphur-free gunpowder being tested at various distances shooting at Steel.")
        Console.WriteLine("Material ¦ Distance (Meters) ¦ Acurracy (As Percentage) ¦ Time (Seconds)")
        Accuracy_Table(10, 100, 900, 180.8, "Steel")
        Accuracy_Table(10, 100, 500, 180.8, "Steel")
        Accuracy_Table(10, 100, 100, 180.8, "Steel")
        Console.ReadLine()

        'Sulphur-free gunpowder shooting at titanium with various distances.
        Console.WriteLine("Sulphur-free gunpowder being tested at various distances shooting at Titanium.")
        Console.WriteLine("Material ¦ Distance (Meters) ¦ Acurracy (As Percentage) ¦ Time (Seconds)")
        Accuracy_Table(10, 100, 900, 100.3, "Titanium")
        Accuracy_Table(10, 100, 500, 100.3, "Titanium")
        Accuracy_Table(10, 100, 100, 100.3, "Titanium")
        Console.ReadLine()

        'Results gained from the tests completed above.
        Console.WriteLine("From the tests completed above this shows that if the distance from the gun to the target is increased, the accuracy and the time taken increases also.")
        Console.WriteLine("With use of different materials they all gave different results, from the tests done it shows that when using the wood it has the smallest average time but is the least accurate " _
                          & "out of the three materials. Furthermore, the titanium took the longest average time but turned out to be the most accurate and the steel fell between the wood and the titanium.")
        Console.ReadLine()

    End Sub

End Module
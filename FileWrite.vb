Module Module1

    Sub Main()
        'Creating a directory for the text file (optional, could be useful for keeping track of multiple text files)
        My.Computer.FileSystem.CreateDirectory("FolderName")

        'Creating/writing to a text file, located within the directory we just created above
        My.Computer.FileSystem.WriteAllText("FolderName\TextFile.txt", "Text String" & vbNewLine, True)

        'The directory containing the file is located in the project folder: "[Project name]\obj\debug"
        'Note that this folder is only created after the solution is built at least once beforehand
    End Sub

End Module

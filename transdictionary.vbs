Class TransDictionary
' https://www.codeproject.com/Articles/51770/Simple-Persistent-Transactional-Dictionary-in-VBSc

    Dim dictionary,logFile
    Dim rollBk,rollFw,lfn,fso,clean
    
    Public Sub Class_Initialize
        Set Me.dictionary=CreateObject("Scripting.dictionary")
        Set Me.rollBk=CreateObject("Scripting.dictionary")
        Set Me.rollFw=CreateObject("Scripting.dictionary")
        Me.clean=TRUE
    End Sub
    
    ' This must be called immediately after the class
    ' is instantiated so that it can read and write its
    ' log file
    Public Sub LoadLog(logFileName)
        Me.lfn=logFileName
        Set Me.fso=CreateObject("Scripting.FileSystemObject")
        Set Me.logFile=Me.fso.OpenTextFile(Me.lfn,1,true)
        While Not Me.logFile.AtEndOfStream
            action Me.logFile.ReadLine()
        Wend
        Me.logFile.Close
        Set Me.logFile=Me.fso.OpenTextFile(Me.lfn,8,false)
    End Sub

    Private Sub Class_Terminate
        On Error Resume Next
        Me.logFile.Close
    End Sub
    
    ' This method takes the appropriate action given a row from
    ' a log file 
    Private Sub action(line)
        Dim row
        row=Split(line,",")
        If row(0)="S" Then
            internalSet Unescape(row(1)),Unescape(row(2))
        ElseIf row(0)="R" Then
            internalRemove Unescape(row(1))
        End If
    End Sub
    
    ' Adds a Remove record to the log file
    Private Sub addRemove(key)
        Me.rollBk.Add Me.rollBk.count,"S," & Escape(key) & _
                      "," & Escape(Me.dictionary.Item(key))
        Me.rollFw.Add Me.rollFw.count,"R," & Escape(key)
    End Sub
    
    ' Adds an Add record to the log file
    Private Sub addSet(key,value)
        Me.rollBk.Add Me.rollBk.count,"R," & Escape(key)
        Me.rollFw.Add Me.rollFw.count,"S," & Escape(key) & "," & Escape(value) 
    End Sub
    
    ' Sets a key,value pair in the internal dictionary. It either adds or replaces
    ' the pair according it if the key is already in the internal dictionary
    Private Sub internalSet(key,value)
        If Me.dictionary.Exists(key) Then Me.dictionary.Remove key
        Me.dictionary.Add key,value
        Me.clean=False
    End Sub
    
    ' Removes a key,value pair from the dictionary
    Private Sub internalRemove(key)
        If Me.dictionary.Exists(key) Then Me.dictionary.Remove key
        Me.clean=False
    End Sub
    
    ' This writes all the changes to the internal dictionary since the
    ' last commit to the log file.
    Public Sub Commit
        Dim i,c
        c=Me.rollFw.Count -1
        For i=0 To c
            Me.logFile.WriteLine Me.rollFw.Item(i)
            ' this is the only way to force a flush of the
            ' text stream object :(
            Me.logFile.Close
            Set Me.logFile=Me.fso.OpenTextFile(Me.lfn,8,false)
        Next 
        Me.rollFw.RemoveAll
        Me.rollBk.RemoveAll
        Me.clean=True
    End Sub
    
    ' This reverts the internal dictionary to the state it was when Commit
    ' was last called - or if Commit has never been called, to the state it
    ' was immediately after having read the log file for the first time
    Public Sub RollBack
        Dim i,c
        c=Me.rollBk.Count -1
        For i=c To 0 Step -1
            action Me.rollBk.Item(i)
        Next 
        Me.rollBk.RemoveAll
        Me.rollFw.RemoveAll
        Me.clean=true
    End Sub
    
    ' This creates a new log file which contains only records to 
    ' recreate the internal dictionary. This cannot be done unless
    ' the internal dictionary is clean (IE no changes since start
    ' or  the last commit). The resultant log file can be used as
    ' a direct replacement for the current log file and so this
    ' can be used to reduce the size and read performance hit of the
    ' log file next time the class is instantiated.
    Public Sub CreateCleanLog(newFileName)
        If Not Me.clean Then 
            Err.Raise -1,"Not Me.clean, commit or rollback first"
        End If
        Me.logFile.Close
        Dim olfn
        olfn=Me.lfn
        Me.lfn=newFileName
        Set Me.logFile=Me.fso.OpenTextFile(Me.lfn,2,true)
        Dim k
        For Each k In Me.dictionary.Keys
            addSet k,Me.dictionary.Item(k)
        Next
        Commit
        Me.logFile.Close
        Me.lfn=olfn
        Set Me.logFile=Me.fso.OpenTextFile(Me.lfn,8,false)        
    End Sub
    
    ' This method adds or replaces a key value pair in the internal
    ' dictionary. The change will not be reflected in the log file
    ' until a Commit is made. Rollback will remove the change unless
    ' a Commit is called and the change will not be persisted until a
    ' Commit is made.
    Public Sub SetValue(key,value)
        If Me.dictionary.Exists(key) Then
            addRemove key
        End If
        addSet key,value
        internalSet key,value
    End Sub

    ' This method removes all key value pairs from the internal
    ' dictionary. The change will not be reflected in the log file
    ' until a Commit is made. Rollback will remove the change unless
    ' a Commit is called and the change will not be persisted until a
    ' Commit is made.
    Public Sub RemoveAll
        Dim k
        For Each k In Me.dictionary.Keys
            Remove(k)
        Next
    End Sub
    
    ' This method removes a key value pair from the internal
    ' dictionary. The change will not be reflected in the log file
    ' until a Commit is made. Rollback will remove the change unless
    ' a Commit is called and the change will not be persisted until a
    ' Commit is made.
    Public Sub Remove(key)
        If Me.dictionary.Exists(key) Then 
            addRemove(key)
            internalRemove(key)
        End If
    End Sub
    
    ' This returns an array of all the values in the internal dictionary
    Public Function Items()
        Items=Me.dictionary.Items
    End Function
    
    ' This removes the value associated with passed key in the internal
    ' dictionary or NULL if the key is not present.
    Public Function Item(key)
        Item=Me.dictionary.Item(key)
    End Function
    
    ' This returns an array of all the keys in the internal dictionary
    Public Function Keys()
        Keys=Me.dictionary.Keys
    End Function
    
    ' This returns true if the passed key is in the internal dictionary
    ' and false otherwise.
    Public Function Exists(key)
        Exists=Me.dictionary.Exists(key)
    End Function
        
End Class
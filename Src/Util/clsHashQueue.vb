
Option Infer On

''' D'après le lien suivant :
''' http://stackoverflow.com/questions/823860/c-listt-contains-too-slow
''' <summary>
''' This is a class that mimics a queue, except the Contains() operation is O(1) 
'''  rather than O(n) thanks to an internal dictionary.
''' The dictionary remembers the hashcodes of the items that have been enqueued and dequeued.
''' Hashcode collisions are stored in a queue to maintain FIFO order.
''' </summary>
''' <typeparam name="T"></typeparam>
Public Class HashQueue(Of T) : Inherits Queue(Of T)

    'Public ReadOnly _hashes As Dictionary(Of Integer, Queue(Of T)) ' CA2104
    Public _hashes As Dictionary(Of Integer, Queue(Of T))

    Private ReadOnly _comp As IEqualityComparer(Of T)

    ' _hashes.Count doesn't always equal base.Count (due to collisions)
    Public Sub New(Optional comp As IEqualityComparer(Of T) = Nothing)
        MyBase.New()
        Me._comp = comp
        Me._hashes = New Dictionary(Of Integer, Queue(Of T))()
    End Sub

    Public Sub New(capacity%, Optional comp As IEqualityComparer(Of T) = Nothing)
        MyBase.New(capacity)
        Me._comp = comp
        Me._hashes = New Dictionary(Of Integer, Queue(Of T))(capacity)
    End Sub

    Public Sub New(collection As IEnumerable(Of T), _
        Optional comp As IEqualityComparer(Of T) = Nothing)

        MyBase.New(collection)
        Me._comp = comp

        Me._hashes = New Dictionary(Of Integer, Queue(Of T))(MyBase.Count)
        For Each item In collection
            Me.EnqueueDictionary(item)
        Next

    End Sub

    Public Shadows Sub Enqueue(item As T)
        MyBase.Enqueue(item)
        ' Add to queue
        Me.EnqueueDictionary(item)
    End Sub

    Private Sub EnqueueDictionary(item As T)
        Dim hash As Integer = If(Me._comp Is Nothing, item.GetHashCode(), _
            Me._comp.GetHashCode(item))
        Dim temp As Queue(Of T) = Nothing
        If Not Me._hashes.TryGetValue(hash, temp) Then
            temp = New Queue(Of T)()
            Me._hashes.Add(hash, temp)
        End If
        temp.Enqueue(item)
    End Sub

    Public Shadows Function Dequeue() As T
        Dim result As T = MyBase.Dequeue()
        ' Remove from queue
        Dim hash As Integer = If(Me._comp Is Nothing, result.GetHashCode(), _
            Me._comp.GetHashCode(result))
        Dim temp As Queue(Of T) = Nothing
        If Me._hashes.TryGetValue(hash, temp) Then
            temp.Dequeue()
            If temp.Count = 0 Then
                Me._hashes.Remove(hash)
            End If
        End If
        Return result
    End Function

    Public Shadows Function Contains(item As T) As Boolean
        ' This is O(1), whereas Queue.Contains is (n)
        Dim hash As Integer = If(Me._comp Is Nothing, item.GetHashCode(), _
            Me._comp.GetHashCode(item))
        Return Me._hashes.ContainsKey(hash)
    End Function

    Public Shadows Sub Clear()

        For Each item In Me._hashes.Values
            item.Clear()
        Next
        ' Clear collision lists
        Me._hashes.Clear()
        ' Clear dictionary
        MyBase.Clear()
        ' Clear queue

    End Sub

End Class
Attribute VB_Name = "modInternet"
Option Explicit

'------------------------------------------------------------------
' Purpose   : Dectects connection of the internet
'------------------------------------------------------------------
Public Function Connection(sckTemp As Winsock)
    On Error Resume Next
    If sckTemp.State <> sckClosed Then sckTemp.Close 'if connection is not closed then close it
    sckTemp.Connect Site, Port 'Tells where to connect to and the port
    Do Until sckTemp.State = sckConnected ' loop until connected
        DoEvents
        Select Case sckTemp.State
            Case sckConnected
                StatusBarMsg "Connected  ", 3
            Case sckConnecting
                StatusBarMsg "Connecting  ", 3
            Case sckResolvingHost
                StatusBarMsg "Resolving Host  ", 3
            Case sckHostResolved
                StatusBarMsg "Host Resolved  ", 3
            Case sckConnectionPending
                StatusBarMsg "Connection Pending  ", 3
            Case sckClosed
                StatusBarMsg "Not Connected  ", 3
            Case Else
                If sckTemp.State <> sckClosed Then sckTemp.Close
                StatusBarMsg "Error connecting to IMDB!  ", 3
                Exit Function
        End Select
    Loop
End Function
